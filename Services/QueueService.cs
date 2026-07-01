using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OutlookClassifierAddIn5.Data;
using OutlookClassifierAddIn5.ML;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassifierAddIn5.Services
{
    public sealed class QueueBuildOptions
    {
        public int DaysBack { get; set; } = 90;
        public int Cap { get; set; } = 500;
        public bool IgnoreFlagged { get; set; } = true;
        public bool UseIncremental { get; set; } = true;
        public bool ForceFolderRefresh { get; set; }
    }

    public class QueueService
    {
        private const string PrSenderSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";
        private const string PrHasAttachment = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";
        private const string PrFlagStatus = "http://schemas.microsoft.com/mapi/proptag/0x10900003";

        public sealed class Item
        {
            public string EntryId { get; }
            public string Subject { get; }
            public string From { get; }
            public string PredictedFolder { get; }
            public double Confidence { get; }
            public List<(string name, float p)> Top3 { get; }

            public Item(string entryId, string subject, string @from,
                string predictedFolder, double confidence,
                List<(string name, float p)> top3)
            {
                EntryId = entryId;
                Subject = subject;
                From = @from;
                PredictedFolder = predictedFolder;
                Confidence = confidence;
                Top3 = top3 ?? new List<(string name, float p)>();
            }
        }

        private sealed class InboxSnapshotRow
        {
            public string EntryId;
            public string Subject;
            public string FromRaw;
            public string FromFeature;
            public bool HasAttachments;
            public DateTime ReceivedUtc;

            public EmailRow ToEmailRow()
            {
                return new EmailRow
                {
                    Subject = Subject ?? string.Empty,
                    Body = string.Empty,
                    FromAddress = FromFeature ?? string.Empty,
                    SenderDomain = SenderResolutionService.ExtractDomain(FromFeature ?? string.Empty),
                    HasAttachments = HasAttachments,
                    Label = string.Empty
                };
            }
        }

        private sealed class QueueStateEntry
        {
            public string EntryId;
            public DateTime ReceivedUtc;
            public string ModelSignature;
            public string FolderVersion;
            public Item Item;
        }

        private readonly ModelService _ml;
        private readonly Queue<Item> _low = new Queue<Item>();
        private readonly List<Item> _medium = new List<Item>();
        private readonly List<Item> _high = new List<Item>();
        private readonly Dictionary<string, QueueStateEntry> _state =
            new Dictionary<string, QueueStateEntry>(StringComparer.Ordinal);

        public IEnumerable<Item> High { get { return _high; } }
        public IEnumerable<Item> Medium { get { return _medium; } }
        public IEnumerable<Item> Low { get { return _low; } }
        public FolderSnapshotService FolderSnapshot { get; private set; }

        public QueueService(ModelService ml, FeedbackStore store)
        {
            _ml = ml;
            FolderSnapshot = new FolderSnapshotService();
        }

        public bool IgnoreFlaggedForBatching { get; set; } = true;
        public string CurrentModelHash { get { return (_ml != null) ? _ml.ModelSignature : "none"; } }

        public async Task BuildQueuesAsync(Outlook.Application app)
        {
            await BuildQueuesAsync(app, new QueueBuildOptions(), CancellationToken.None);
        }

        public async Task BuildQueuesAsync(Outlook.Application app, int daysBack, CancellationToken ct)
        {
            await BuildQueuesAsync(app, new QueueBuildOptions
            {
                DaysBack = daysBack,
                Cap = 500,
                IgnoreFlagged = IgnoreFlaggedForBatching,
                UseIncremental = true
            }, ct);
        }

        public async Task BuildQueuesAsync(Outlook.Application app, QueueBuildOptions options, CancellationToken ct)
        {
            if (app == null || _ml == null) return;
            options = options ?? new QueueBuildOptions();

            var totalSw = Stopwatch.StartNew();
            AppLogger.Info("Queue build start. DaysBack=" + options.DaysBack + ", cap=" + options.Cap + ", incremental=" + options.UseIncremental + ".");

            var folderSnapshot = FolderSnapshot.Refresh(app, options.ForceFolderRefresh);
            ct.ThrowIfCancellationRequested();

            var snapshotRows = SnapshotInboxRows(app, options, ct);
            ct.ThrowIfCancellationRequested();

            var modelSignature = CurrentModelHash;
            var scoringInput = new List<InboxSnapshotRow>();
            var reused = new List<Item>();
            var snapshotIds = new HashSet<string>(StringComparer.Ordinal);

            foreach (var row in snapshotRows)
            {
                snapshotIds.Add(row.EntryId);

                QueueStateEntry existing;
                if (options.UseIncremental
                    && _state.TryGetValue(row.EntryId, out existing)
                    && existing.Item != null
                    && existing.ReceivedUtc == row.ReceivedUtc
                    && string.Equals(existing.ModelSignature, modelSignature, StringComparison.Ordinal)
                    && string.Equals(existing.FolderVersion, folderSnapshot.Version, StringComparison.Ordinal))
                {
                    reused.Add(existing.Item);
                }
                else
                {
                    scoringInput.Add(row);
                }
            }

            var scored = await ScoreRowsAsync(scoringInput, folderSnapshot, modelSignature, ct);
            ct.ThrowIfCancellationRequested();

            var allItems = reused.Concat(scored.Select(s => s.Item)).ToList();

            foreach (var missing in _state.Keys.Where(k => !snapshotIds.Contains(k)).ToList())
                _state.Remove(missing);

            foreach (var scoredRow in scored)
            {
                _state[scoredRow.EntryId] = scoredRow;
            }

            var high = new List<Item>();
            var medium = new List<Item>();
            var low = new Queue<Item>();

            foreach (var item in allItems.Where(i => i != null).OrderByDescending(i => i.Confidence))
            {
                if (item.Confidence >= 0.92) high.Add(item);
                else if (item.Confidence >= 0.70) medium.Add(item);
                else low.Enqueue(item);
            }

            medium.Sort((a, b) => Margin(a).CompareTo(Margin(b)));

            _high.Clear();
            _medium.Clear();
            _low.Clear();
            _high.AddRange(high);
            _medium.AddRange(medium);
            foreach (var item in low) _low.Enqueue(item);

            totalSw.Stop();
            AppLogger.Info("Queue build end. SnapshotRows=" + snapshotRows.Count +
                ", scored=" + scoringInput.Count +
                ", reused=" + reused.Count +
                ", high=" + _high.Count +
                ", medium=" + _medium.Count +
                ", low=" + _low.Count +
                ", elapsedMs=" + totalSw.ElapsedMilliseconds + ".");
        }

        public void InvalidateState()
        {
            _state.Clear();
        }

        public Item PeekLow() { return _low.Count > 0 ? _low.Peek() : null; }
        public async Task PopLowAsync() { if (_low.Count > 0) _low.Dequeue(); await Task.CompletedTask; }

        private List<InboxSnapshotRow> SnapshotInboxRows(Outlook.Application app, QueueBuildOptions options, CancellationToken ct)
        {
            var sw = Stopwatch.StartNew();
            var rows = new List<InboxSnapshotRow>(Math.Max(0, options.Cap));
            Outlook.Table table = null;

            try
            {
                var ns = app.Session;
                var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                var cutoffUtc = DateTime.UtcNow.AddDays(-Math.Max(1, options.DaysBack));
                var cutoffLocal = cutoffUtc.ToLocalTime().ToString("g", CultureInfo.CreateSpecificCulture("en-US"));
                var filter = "[MessageClass] = 'IPM.Note' AND [ReceivedTime] >= '" + cutoffLocal + "'";

                table = inbox.GetTable(filter, Outlook.OlTableContents.olUserItems);
                table.Columns.RemoveAll();
                table.Columns.Add("EntryID");
                table.Columns.Add("Subject");
                table.Columns.Add("SenderEmailAddress");
                TryAddColumn(table, PrSenderSmtpAddress);
                table.Columns.Add("ReceivedTime");
                table.Columns.Add(PrHasAttachment);
                table.Columns.Add(PrFlagStatus);

                try { table.Sort("[ReceivedTime]", Outlook.OlSortOrder.olDescending); }
                catch (Exception ex) { AppLogger.Warn("Queue table sort failed: " + ex.Message); }

                Outlook.Row r;
                while (!ct.IsCancellationRequested && rows.Count < Math.Max(1, options.Cap) && (r = table.GetNextRow()) != null)
                {
                    var flagStatus = SafeRowInt(r, PrFlagStatus);
                    if (options.IgnoreFlagged && flagStatus == 2)
                        continue;

                    var fromRaw = SafeRowString(r, "SenderEmailAddress");
                    var fromSmtp = SafeRowString(r, PrSenderSmtpAddress);
                    var featureSender = SenderResolutionService.ChooseFeatureAddress(fromSmtp, fromRaw);
                    var displaySender = !string.IsNullOrWhiteSpace(featureSender) && !SenderResolutionService.LooksLegacyDn(featureSender)
                        ? featureSender
                        : (fromRaw ?? string.Empty);

                    rows.Add(new InboxSnapshotRow
                    {
                        EntryId = SafeRowString(r, "EntryID") ?? string.Empty,
                        Subject = SafeRowString(r, "Subject") ?? string.Empty,
                        FromRaw = displaySender,
                        FromFeature = featureSender,
                        HasAttachments = SafeRowBool(r, PrHasAttachment),
                        ReceivedUtc = SafeRowDate(r, "ReceivedTime").ToUniversalTime()
                    });
                }
            }
            finally
            {
                if (table != null)
                {
                    try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(table); } catch { }
                }

                sw.Stop();
                AppLogger.Info("Inbox queue snapshot complete. Rows=" + rows.Count + ", elapsedMs=" + sw.ElapsedMilliseconds + ".");
            }

            return rows.Where(r => !string.IsNullOrEmpty(r.EntryId)).ToList();
        }

        private Task<List<QueueStateEntry>> ScoreRowsAsync(
            List<InboxSnapshotRow> rows,
            FolderSnapshot folderSnapshot,
            string modelSignature,
            CancellationToken ct)
        {
            if (rows == null || rows.Count == 0)
                return Task.FromResult(new List<QueueStateEntry>());

            return Task.Run(() =>
            {
                var sw = Stopwatch.StartNew();
                var result = new List<QueueStateEntry>();
                var modelRows = rows.Select(r => r.ToEmailRow()).ToList();
                var predictions = _ml.PredictBatch(modelRows).ToList();

                for (int i = 0; i < rows.Count && i < predictions.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    var snapshot = rows[i];
                    var prediction = predictions[i];
                    var item = BuildItem(snapshot, prediction.folder, prediction.top3, prediction.conf, folderSnapshot);
                    if (item == null) continue;

                    result.Add(new QueueStateEntry
                    {
                        EntryId = snapshot.EntryId,
                        ReceivedUtc = snapshot.ReceivedUtc,
                        ModelSignature = modelSignature,
                        FolderVersion = folderSnapshot.Version,
                        Item = item
                    });
                }

                sw.Stop();
                AppLogger.Info("Queue batch scoring complete. Rows=" + rows.Count + ", kept=" + result.Count + ", elapsedMs=" + sw.ElapsedMilliseconds + ".");
                return result;
            }, ct);
        }

        private static Item BuildItem(
            InboxSnapshotRow row,
            string predictedFolder,
            List<(string name, float p)> top3Raw,
            float confidence,
            FolderSnapshot folderSnapshot)
        {
            var top3Canon = new List<(string name, float p)>();
            foreach (var t in top3Raw ?? new List<(string name, float p)>())
            {
                var full = folderSnapshot.CanonicalizeToFull(t.name);
                if (!string.IsNullOrEmpty(full) && folderSnapshot.IsValidFilingFolder(full))
                    top3Canon.Add((full, t.p));
            }

            var bestFull = top3Canon.Count > 0
                ? top3Canon[0].name
                : folderSnapshot.CanonicalizeToFull(predictedFolder);

            var bestConf = top3Canon.Count > 0 ? top3Canon[0].p : confidence;
            if (string.IsNullOrEmpty(bestFull) || !folderSnapshot.IsValidFilingFolder(bestFull))
                return null;

            return new Item(
                row.EntryId,
                row.Subject ?? string.Empty,
                row.FromRaw ?? string.Empty,
                bestFull,
                bestConf,
                top3Canon);
        }

        private static double Margin(Item it) { return it.Top3.Count >= 2 ? (it.Top3[0].p - it.Top3[1].p) : 1.0; }

        private static string SafeRowString(Outlook.Row row, string name)
        {
            try { return row[name] as string; } catch { return string.Empty; }
        }

        private static int SafeRowInt(Outlook.Row row, string name)
        {
            try
            {
                object v = row[name];
                if (v is int) return (int)v;
                int parsed;
                if (v is string && int.TryParse((string)v, out parsed)) return parsed;
            }
            catch { }
            return 0;
        }

        private static DateTime SafeRowDate(Outlook.Row row, string name)
        {
            try
            {
                object v = row[name];
                if (v is DateTime) return (DateTime)v;
                DateTime parsed;
                if (v is string && DateTime.TryParse((string)v, out parsed)) return parsed;
            }
            catch { }
            return DateTime.MinValue;
        }

        private static bool SafeRowBool(Outlook.Row row, string name)
        {
            try
            {
                object v = row[name];
                if (v is bool) return (bool)v;
                if (v is int) return (int)v != 0;
                bool b;
                if (v is string && bool.TryParse((string)v, out b)) return b;
            }
            catch { }
            return false;
        }

        private static void TryAddColumn(Outlook.Table table, string name)
        {
            try { table.Columns.Add(name); }
            catch (Exception ex) { AppLogger.Warn("Could not add Outlook table column " + name + ": " + ex.Message); }
        }
    }
}
