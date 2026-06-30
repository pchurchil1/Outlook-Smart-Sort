using System.Collections.Generic;
using System.Linq;
using System;
using System.Threading;
using System.Globalization;
using System.Threading.Tasks;
using OutlookClassifierAddIn5.Data;
using OutlookClassifierAddIn5.ML;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace OutlookClassifierAddIn5.Services
{

    public class QueueService
    {
        private const string PrSenderSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";

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


        private readonly ModelService _ml;
        private readonly FeedbackStore _store;


        // replace target-typed new() with explicit types
        private readonly Queue<Item> _low = new Queue<Item>();
        private readonly List<Item> _medium = new List<Item>();
        private readonly List<Item> _high = new List<Item>();


        public IEnumerable<Item> High => _high;
        public IEnumerable<Item> Medium => _medium;
        public IEnumerable<Item> Low => _low;
        public QueueService(ModelService ml, FeedbackStore store) { _ml = ml; _store = store; }
        // Ignore messages that are currently flagged (active follow-up)
        public bool IgnoreFlaggedForBatching { get; set; } = true;

        public string CurrentModelHash { get { return (_ml != null) ? _ml.ModelSignature : "none"; } }

        public async Task BuildQueuesAsync(Outlook.Application app)
        {
            await BuildQueuesAsync(app, 90, CancellationToken.None); // default window 90d
        }

        //Set cap and daysBack as needed
        public async Task BuildQueuesAsync(Outlook.Application app, int daysBack, CancellationToken ct)
        {
            _low.Clear(); _medium.Clear(); _high.Clear();
            if (app == null) return;

            AppLogger.Info("Queue build start. DaysBack=" + daysBack + ".");

            // Snapshot valid filing folders
            var validPaths = new List<string>();
            var all = FolderMap.GetAllFolderPaths(app);
            if (all != null)
            {
                foreach (var p in all)
                    if (IsValidFilingFolder(p))
                        validPaths.Add(FolderPathNormalizer.Normalize(p));
            }

            var ns = app.Session;
            var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            var cutoffUtc = DateTime.UtcNow.AddDays(-Math.Max(1, daysBack));
            var enUS = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            var cutoffLocal = cutoffUtc.ToLocalTime().ToString("g", enUS);

            var filter = "[MessageClass] = 'IPM.Note' AND [ReceivedTime] >= '" + cutoffLocal + "'";
            var table = inbox.GetTable(filter, Outlook.OlTableContents.olUserItems);

            table.Columns.RemoveAll();
            table.Columns.Add("EntryID");
            table.Columns.Add("Subject");
            table.Columns.Add("SenderEmailAddress");
            TryAddColumn(table, PrSenderSmtpAddress);
            table.Columns.Add("ReceivedTime");
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"); // HasAttachment
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x10900003"); // FlagStatus

            try { table.Sort("[ReceivedTime]", Outlook.OlSortOrder.olDescending); } catch (Exception ex) { AppLogger.Warn("Queue table sort failed: " + ex.Message); }

            const int cap = 500;
            int count = 0;
            Outlook.Row r;
            while (!ct.IsCancellationRequested && count < cap && (r = table.GetNextRow()) != null)
            {
                var entryId = SafeRowString(r, "EntryID");
                var subject = SafeRowString(r, "Subject");
                var fromRaw = SafeRowString(r, "SenderEmailAddress");
                var fromSmtp = SafeRowString(r, PrSenderSmtpAddress);
                var featureSender = SenderResolutionService.ChooseFeatureAddress(fromSmtp, fromRaw);
                var hasAtt = SafeRowBool(r, "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B");

                int flagStatus = 0;
                try
                {
                    object fv = r["http://schemas.microsoft.com/mapi/proptag/0x10900003"];
                    if (fv is int) flagStatus = (int)fv;
                    else if (fv is string) { int tmp; if (int.TryParse((string)fv, out tmp)) flagStatus = tmp; }
                }
                    catch (Exception ex) { AppLogger.Warn("Queue flag read failed: " + ex.Message); }

                if (IgnoreFlaggedForBatching && flagStatus == 2) // olFlagMarked
                    continue;

                var modelRow = new EmailRow
                {
                    Subject = subject ?? string.Empty,
                    Body = string.Empty,                         // fast path: no Body
                    FromAddress = featureSender,
                    SenderDomain = SenderResolutionService.ExtractDomain(featureSender),
                    HasAttachments = hasAtt,
                    Label = string.Empty
                };

                var (folderRaw, top3Raw, confRaw) = _ml.PredictTop3(modelRow);

                var top3Canon = new List<(string name, float p)>();
                if (top3Raw != null)
                {
                    foreach (var t in top3Raw)
                    {
                        var full = CanonicalizeToFull(t.name, validPaths, app);
                        if (!string.IsNullOrEmpty(full))
                            top3Canon.Add((full, t.p));
                    }
                }

                var filtered = top3Canon.Where(t => IsValidFilingFolder(t.name)).ToList();

                string bestFull = filtered.Count > 0
                    ? filtered[0].name
                    : CanonicalizeToFull(folderRaw, validPaths, app);

                double bestConf = filtered.Count > 0 ? filtered[0].p : confRaw;

                if (string.IsNullOrEmpty(bestFull) || !IsValidFilingFolder(bestFull))
                    continue;

                var item = new Item(
                    entryId,
                    subject ?? string.Empty,
                    fromRaw ?? string.Empty,               // UI will pretty-print
                    bestFull,
                    bestConf,
                    filtered
                );

                if (bestConf >= 0.92) _high.Add(item);
                else if (bestConf >= 0.70) _medium.Add(item);
                else _low.Enqueue(item);

                count++;
            }

            _medium.Sort((a, b) => Margin(a).CompareTo(Margin(b)));
            AppLogger.Info("Queue build end. High=" + _high.Count + ", medium=" + _medium.Count + ", low=" + _low.Count + ".");
            await Task.CompletedTask;

        }

        // --- helpers ---

        private static bool IsValidFilingFolder(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            // Exclude all system/default folders
            if (FolderMap.IsSystemFolder(path)) return false;
            // Allow Inbox subfolders and any user-created folder
            return FolderMap.IsUnderInboxOrDeleted(path) || FolderMap.IsUserCreatedFolder(path);
        }

        private static string Normalize(string label)
        {
            if (string.IsNullOrWhiteSpace(label)) return string.Empty;
            return label.Replace("\\", "/").Trim().Trim('/');
        }

        private static string CanonicalizeToFull(string label, List<string> validPaths, Outlook.Application app)
        {
            if (string.IsNullOrWhiteSpace(label)) return string.Empty;
            var s = FolderPathNormalizer.Normalize(label);

            // Exact path match? done.
            if (s.IndexOf('/') >= 0)
            {
                var fullExact = validPaths.FirstOrDefault(p => string.Equals(p, s, System.StringComparison.OrdinalIgnoreCase));
                return fullExact ?? string.Empty;
            }

            // Leaf name -> unique full path?
            var matches = validPaths.Where(p =>
            {
                var parts = p.Split(new[] { '/' }, System.StringSplitOptions.RemoveEmptyEntries);
                var leaf = parts.Length > 0 ? parts[parts.Length - 1] : p;
                return string.Equals(leaf, s, System.StringComparison.OrdinalIgnoreCase);
            }).ToList();

            if (matches.Count == 1) return matches[0];

            // If ambiguous, prefer under the default store root if possible,
            // otherwise return empty to avoid wrong grouping/moves.
            try
            {
                var defaultRoot = app?.Session?.DefaultStore?.DisplayName;
                if (!string.IsNullOrEmpty(defaultRoot))
                {
                    var underDefault = matches.FirstOrDefault(p => p.StartsWith(defaultRoot + "/", System.StringComparison.OrdinalIgnoreCase));
                    if (!string.IsNullOrEmpty(underDefault))
                        return underDefault;
                }
            }
            catch (Exception ex) { AppLogger.Warn("Default store preference failed during canonicalization: " + ex.Message); }

            return string.Empty;
        }


        public Item PeekLow() => _low.Count > 0 ? _low.Peek() : null;
        public async Task PopLowAsync() { if (_low.Count > 0) _low.Dequeue(); await Task.CompletedTask; }

        private static double Margin(Item it) => it.Top3.Count >= 2 ? (it.Top3[0].p - it.Top3[1].p) : 1.0;


        private static bool IsInboxPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            var s = path.ToLowerInvariant();
            return s == "inbox" || s.EndsWith("/inbox");
        }


        private static string SafeRowString(Outlook.Row row, string name)
        {
            try { return row[name] as string; } catch { return string.Empty; }
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
