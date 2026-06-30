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

    // ---- Cache DTOs ----
    public sealed class QueueCache
    {
        public string FolderSnapshotHash;
        public DateTime BuiltUtc;

        public List<BatchItemDto> High = new List<BatchItemDto>();
        public List<BatchItemDto> Medium = new List<BatchItemDto>();
        public List<BatchItemDto> Low = new List<BatchItemDto>();
    }

    public sealed class BatchItemDto
    {
        public string EntryId;
        public string Subject;
        public string From;
        public string PredictedFullPath;
        public double Confidence;
        public List<Tuple<string, float>> Top3; // name, p
    }


    public class QueueService
    {
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

        // Simple "model hash" placeholder; if you later expose a real model signature, return it here.
        public string CurrentModelHash { get { return (_ml != null) ? "loaded" : "none"; } }

        // Compute a stable hash of folder full paths, to know when cache is stale
        public static string ComputeFolderSnapshotHash(IEnumerable<string> fullPaths)
        {
            var list = new List<string>();
            foreach (var p in fullPaths ?? new List<string>())
            {
                var s = (p ?? "").Replace('\\', '/').Trim();
                list.Add(s);
            }
            list.Sort(StringComparer.OrdinalIgnoreCase);
            var sAll = string.Join("\n", list.ToArray());
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = System.Text.Encoding.UTF8.GetBytes(sAll);
                var hash = sha.ComputeHash(bytes);
                return BitConverter.ToString(hash).Replace("-", "");
            }
        }

        // Persist/restore cache
        public QueueCache LoadCache(string path)
        {
            try
            {
                if (!System.IO.File.Exists(path)) return null;
                var json = System.IO.File.ReadAllText(path);
                return Newtonsoft.Json.JsonConvert.DeserializeObject<QueueCache>(json);
            }
            catch { return null; }
        }

        public void SaveCache(string path, QueueCache cache)
        {
            try
            {
                var dir = System.IO.Path.GetDirectoryName(path);
                if (!System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir);
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(cache);
                System.IO.File.WriteAllText(path, json);
            }
            catch { /* ignore */ }
        }

        // Build a cache snapshot from current in-memory items
        public QueueCache SnapshotToCache(IEnumerable<string> folderFullPaths)
        {
            var cache = new QueueCache();
            cache.FolderSnapshotHash = ComputeFolderSnapshotHash(folderFullPaths);
            cache.BuiltUtc = DateTime.UtcNow;

            Func<Item, BatchItemDto> toDto = it =>
            {
                var dto = new BatchItemDto();
                dto.EntryId = it.EntryId;
                dto.Subject = it.Subject;
                dto.From = it.From;
                dto.PredictedFullPath = it.PredictedFolder;
                dto.Confidence = it.Confidence;
                // Top3: convert value tuples to serializable tuples
                var list = new List<Tuple<string, float>>();
                foreach (var t in it.Top3)
                    list.Add(Tuple.Create(t.name, t.p));
                dto.Top3 = list;
                return dto;
            };

            foreach (var it in _high) cache.High.Add(toDto(it));
            foreach (var it in _medium) cache.Medium.Add(toDto(it));
            foreach (var it in _low) cache.Low.Add(toDto(it));

            return cache;
        }

        // Restore in-memory queues from cache (fast bind path)
        public void SetFromCache(QueueCache cache)
        {
            _high.Clear(); _medium.Clear(); _low.Clear();
            if (cache == null) return;

            Func<BatchItemDto, Item> fromDto = d =>
            {
                var top3 = new List<(string name, float p)>();
                if (d.Top3 != null)
                {
                    foreach (var t in d.Top3)
                        top3.Add((t.Item1, t.Item2));
                }
                return new Item(
                    d.EntryId ?? string.Empty,
                    d.Subject ?? string.Empty,
                    d.From ?? string.Empty,
                    d.PredictedFullPath ?? string.Empty,
                    d.Confidence,
                    top3
                );
            };

            foreach (var d in cache.High) _high.Add(fromDto(d));
            foreach (var d in cache.Medium) _medium.Add(fromDto(d));
            foreach (var d in cache.Low) _low.Enqueue(fromDto(d));
        }

        public async Task BuildQueuesAsync(Outlook.Application app)
        {
            await BuildQueuesAsync(app, 90, CancellationToken.None); // default window 90d
        }

        //Set cap and daysBack as needed
        public async Task BuildQueuesAsync(Outlook.Application app, int daysBack, CancellationToken ct)
        {
            _low.Clear(); _medium.Clear(); _high.Clear();
            if (app == null) return;

            // Snapshot valid filing folders
            var validPaths = new List<string>();
            var all = FolderMap.GetAllFolderPaths(app);
            if (all != null)
            {
                foreach (var p in all)
                    if (IsValidFilingFolder(p))
                        validPaths.Add(Normalize(p));
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
            table.Columns.Add("ReceivedTime");
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"); // HasAttachment
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x10900003"); // FlagStatus

            try { table.Sort("[ReceivedTime]", Outlook.OlSortOrder.olDescending); } catch { }

            const int cap = 500;
            int count = 0;
            Outlook.Row r;
            while (!ct.IsCancellationRequested && count < cap && (r = table.GetNextRow()) != null)
            {
                var entryId = SafeRowString(r, "EntryID");
                var subject = SafeRowString(r, "Subject");
                var fromRaw = SafeRowString(r, "SenderEmailAddress");
                var hasAtt = SafeRowBool(r, "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B");

                int flagStatus = 0;
                try
                {
                    object fv = r["http://schemas.microsoft.com/mapi/proptag/0x10900003"];
                    if (fv is int) flagStatus = (int)fv;
                    else if (fv is string) { int tmp; if (int.TryParse((string)fv, out tmp)) flagStatus = tmp; }
                }
                catch { }

                if (IgnoreFlaggedForBatching && flagStatus == 2) // olFlagMarked
                    continue;

                var modelRow = new EmailRow
                {
                    Subject = subject ?? string.Empty,
                    Body = string.Empty,                         // fast path: no Body
                    FromAddress = fromRaw ?? string.Empty,              // raw to match training
                    SenderDomain = ExtractDomain(fromRaw ?? string.Empty),
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
            var s = Normalize(label);

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
            catch { /* ignore and fall through */ }

            return string.Empty;
        }


        public Item PeekLow() => _low.Count > 0 ? _low.Peek() : null;
        public async Task PopLowAsync() { if (_low.Count > 0) _low.Dequeue(); await Task.CompletedTask; }

        private static string ExtractDomain(string addr)
        {
            int at = addr.IndexOf('@');
            if (at < 0)
                return string.Empty;


            return addr.Substring(at + 1).ToLowerInvariant();
        }
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

        private static bool IsActivelyFlagged(Outlook.MailItem m)
        {
            try
            {
                // FlagStatus covers most cases:
                //  - olFlagMarked   => user flagged (active)
                //  - olFlagComplete => completed (treat as not flagged)
                //  - olNoFlag       => not flagged
                var fs = m.FlagStatus;
                if (fs == Outlook.OlFlagStatus.olFlagMarked) return true;
                if (fs == Outlook.OlFlagStatus.olFlagComplete) return false;

                // Some stores use IsMarkedAsTask for follow-ups
                if (m.IsMarkedAsTask) return true;
            }
            catch { /* some providers don’t expose flags cleanly */ }
            return false;
        }

        // Resolve display + SMTP for a MailItem sender
        private static void ResolveSender(Outlook.MailItem m, out string display, out string smtp)
        {
            display = m?.SenderName ?? string.Empty;
            smtp = m?.SenderEmailAddress ?? string.Empty;

            try
            {
                // Already SMTP-looking? done.
                if (!string.IsNullOrEmpty(smtp) &&
                    !smtp.StartsWith("/O=", StringComparison.OrdinalIgnoreCase))
                    return;

                var sender = m?.Sender;
                if (sender != null)
                {
                    // Exchange internal/remote user
                    if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                        sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        var ex = sender.GetExchangeUser();
                        if (ex != null)
                        {
                            if (!string.IsNullOrEmpty(ex.PrimarySmtpAddress)) smtp = ex.PrimarySmtpAddress;
                            if (!string.IsNullOrEmpty(ex.Name)) display = ex.Name;
                        }
                    }

                    // Try generic SMTP MAPI property
                    if (string.IsNullOrEmpty(smtp))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        try
                        {
                            var pa = sender.PropertyAccessor;
                            var addr = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                            if (!string.IsNullOrEmpty(addr)) smtp = addr;
                        }
                        catch { /* ignore */ }
                    }

                    if (string.IsNullOrEmpty(display))
                        display = sender.Name ?? display;
                }

                // Keep empty instead of legacy DN if still unresolved
                if (string.IsNullOrEmpty(smtp) ||
                    smtp.StartsWith("/O=", StringComparison.OrdinalIgnoreCase))
                    smtp = string.Empty;
            }
            catch
            {
                // leave best effort
            }
        }

        private static string ComposeFromPretty(string display, string smtp)
        {
            if (!string.IsNullOrEmpty(smtp))
                return string.IsNullOrEmpty(display) ? smtp : (display + " <" + smtp + ">");
            return display ?? string.Empty;
        }
    }
}