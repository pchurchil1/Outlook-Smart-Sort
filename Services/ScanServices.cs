using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using OutlookClassifierAddIn5.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassifierAddIn5.Services
{
    public sealed class ScanService
    {
        private const string PrSenderSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";
        private const string PrInternetMessageId = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";

        private readonly Outlook.Application _app;
        private readonly FeedbackStore _store;

        public ScanService(Outlook.Application app, FeedbackStore store)
        {
            _app = app;
            _store = store;
        }

        public sealed class ScanOptions
        {
            public int CapPerFolder { get; set; } = 400;
            public int DaysBack { get; set; } = 360;
            public int BodySamplePerFolder { get; set; } = 20; // 0 to skip body entirely
            public int BodySnippetLength { get; set; } = 1000;
            public bool OnlyUnderInbox { get; set; } = false;   // avoid scanning non-mail stores
        }

        public async Task EnsureSeedTrainingDataAsync(ScanOptions opts = null, CancellationToken ct = default(CancellationToken))
        {
            if (opts == null) opts = new ScanOptions();
            AppLogger.Info("Training scan start. DaysBack=" + opts.DaysBack + ", capPerFolder=" + opts.CapPerFolder + ".");

            var ns = _app.Session;
            var stores = new List<Outlook.Store>();

            try
            {
                foreach (Outlook.Store s in ns.Stores) stores.Add(s);

                foreach (var store in stores)
                {
                    ct.ThrowIfCancellationRequested();

                    Outlook.MAPIFolder root = null;
                    Outlook.MAPIFolder start = null;

                    try
                    {
                        root = store.GetRootFolder();
                        start = SafeGetStartFolder(ns, store, opts.OnlyUnderInbox, root);
                        await WalkFolderTreeAsync(start, opts, ct);
                    }
                    catch (Exception ex)
                    {
                        AppLogger.Error(ex, "Store scan failed.");
                    }
                    finally
                    {
                        ReleaseCom(start);
                        ReleaseCom(root);
                    }
                }
            }
            finally
            {
                // release the Store COM objects we added to the list
                for (int i = 0; i < stores.Count; i++) ReleaseCom(stores[i]);
                AppLogger.Info("Training scan end.");
            }
        }

        private static Outlook.MAPIFolder SafeGetStartFolder(Outlook.NameSpace ns, Outlook.Store store, bool onlyInbox, Outlook.MAPIFolder fallback)
        {
            if (!onlyInbox) return fallback;
            try
            {
                var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                if (inbox != null && inbox.StoreID == store.StoreID) return inbox;
            }
            catch (Exception ex) { AppLogger.Warn("Could not resolve store inbox: " + ex.Message); }
            return fallback;
        }

        private async Task WalkFolderTreeAsync(Outlook.MAPIFolder folder, ScanOptions opts, CancellationToken ct)
        {
            if (folder == null) return;

            if (!IsSkippedFolder(folder))
            {
                bool isMailFolder = false;
                try { isMailFolder = (folder.DefaultItemType == Outlook.OlItemType.olMailItem); } catch (Exception ex) { AppLogger.Warn("Could not read folder item type: " + ex.Message); }
                if (isMailFolder)
                    await ScanMailFolderAsync(folder, opts, ct);
            }

            // Recurse into children
            foreach (Outlook.MAPIFolder child in folder.Folders)
            {
                try
                {
                    await WalkFolderTreeAsync(child, opts, ct);
                }
                finally
                {
                    ReleaseCom(child);
                }
            }
        }

        private async Task ScanMailFolderAsync(Outlook.MAPIFolder f, ScanOptions opts, CancellationToken ct)
        {
            Outlook.Table table = null;
            try
            {
                var cutoffUtc = DateTime.UtcNow.AddDays(-opts.DaysBack);
                var localCutoff = cutoffUtc.ToLocalTime().ToString("g", CultureInfo.CreateSpecificCulture("en-US"));

                string filter = "[MessageClass] = 'IPM.Note' AND [ReceivedTime] >= '" + localCutoff + "'";

                table = f.GetTable(filter, Outlook.OlTableContents.olUserItems);
                table.Columns.RemoveAll();
                table.Columns.Add("EntryID");
                table.Columns.Add("Subject");
                table.Columns.Add("SenderEmailAddress");
                TryAddColumn(table, PrSenderSmtpAddress);
                TryAddColumn(table, "ConversationID");
                TryAddColumn(table, PrInternetMessageId);
                table.Columns.Add("ReceivedTime");
                // HasAttachment (MAPI tag 0x0E1B000B) — faster/safer in Table than cracking MailItem
                table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B");

                try { table.Sort("[ReceivedTime]", Outlook.OlSortOrder.olDescending); } catch (Exception ex) { AppLogger.Warn("Scan table sort failed: " + ex.Message); }

                var batch = new List<SeedRow>(Math.Min(opts.CapPerFolder, 500));
                int taken = 0;
                string folderPath = SafeFolderPath(f);
                string storeId = SafeFolderStoreId(f);

                Outlook.Row row;
                while (taken < opts.CapPerFolder && (row = table.GetNextRow()) != null)
                {
                    ct.ThrowIfCancellationRequested();

                    string entryId = SafeRowString(row, "EntryID");
                    string subject = SafeRowString(row, "Subject");
                    string senderRaw = SafeRowString(row, "SenderEmailAddress");
                    string senderSmtp = SafeRowString(row, PrSenderSmtpAddress);
                    string sender = SenderResolutionService.ChooseFeatureAddress(senderSmtp, senderRaw);
                    string conversationId = SafeRowString(row, "ConversationID");
                    string internetMessageId = SafeRowString(row, PrInternetMessageId);
                    DateTime receivedLocal = SafeRowDate(row, "ReceivedTime");
                    bool hasAtt = SafeRowBool(row, "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B");
                    string domain = SenderResolutionService.ExtractDomain(sender);
                    DateTime receivedUtc = receivedLocal.ToUniversalTime();

                    var r = new SeedRow();
                    r.EntryId = entryId ?? string.Empty;
                    r.StoreId = storeId;
                    r.InternetMessageId = internetMessageId ?? string.Empty;
                    r.ConversationId = conversationId ?? string.Empty;
                    r.FolderPath = FolderPathNormalizer.Normalize(folderPath);
                    r.Subject = subject ?? string.Empty;
                    r.Body = string.Empty; // fill for a sample later
                    r.Sender = sender ?? string.Empty;
                    r.Domain = domain ?? string.Empty;
                    r.HasAttachments = hasAtt;
                    r.ReceivedUtc = receivedUtc;

                    batch.Add(r);
                    taken++;

                    if (batch.Count >= 400)
                    {
                        await _store.UpsertEmailsAsync(batch, ct);
                        batch.Clear();
                    }
                }

                if (batch.Count > 0)
                    await _store.UpsertEmailsAsync(batch, ct);

                // Optional: fetch and store bodies for a small newest sample
                if (opts.BodySamplePerFolder > 0 && taken > 0)
                    await PopulateBodiesForSampleAsync(f, opts.BodySamplePerFolder, opts.BodySnippetLength, ct);
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Folder scan failed: " + SafeFolderPath(f));
            }
            finally
            {
                ReleaseCom(table);
            }
        }

        private async Task PopulateBodiesForSampleAsync(Outlook.MAPIFolder f, int sampleCount, int bodySnippetLength, CancellationToken ct)
        {
            Outlook.Items items = null;
            try
            {
                items = f.Items;
                items.Sort("[ReceivedTime]", true);
                items = items.Restrict("[MessageClass] = 'IPM.Note'");

                var toWrite = new List<SeedBody>(sampleCount);
                int count = 0;

                foreach (object obj in items)
                {
                    var m = obj as Outlook.MailItem;
                    if (m == null) continue;

                    string body = TrimBody(m.Body, bodySnippetLength);
                    toWrite.Add(new SeedBody(m.EntryID, body));
                    count++;

                    ReleaseCom(m);
                    if (count >= sampleCount) break;
                }

                if (toWrite.Count > 0)
                    await _store.UpsertBodiesAsync(toWrite, ct);
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Body sample scan failed: " + SafeFolderPath(f));
            }
            finally
            {
                ReleaseCom(items);
            }
        }

        private static string SafeFolderPath(Outlook.MAPIFolder f)
        {
            try
            {
                if (!string.IsNullOrEmpty(f.FolderPath)) return FolderPathNormalizer.Normalize(f.FolderPath);
                return f.Name ?? string.Empty;
            }
            catch (Exception ex) { AppLogger.Warn("SafeFolderPath failed: " + ex.Message); return f.Name ?? string.Empty; }
        }

        private static string SafeFolderStoreId(Outlook.MAPIFolder f)
        {
            try { return f.StoreID ?? string.Empty; }
            catch (Exception ex) { AppLogger.Warn("SafeFolderStoreId failed: " + ex.Message); return string.Empty; }
        }

        private static string SafeRowString(Outlook.Row row, string name)
        {
            try { return row[name] as string; } catch { return string.Empty; }
        }

        private static DateTime SafeRowDate(Outlook.Row row, string name)
        {
            try
            {
                object v = row[name];
                DateTime dt;
                if (v is DateTime) return (DateTime)v;
                if (v is string && DateTime.TryParse((string)v, out dt)) return dt;
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

        private static bool IsSkippedFolder(Outlook.MAPIFolder f)
        {
            string name = null;
            try { name = f.Name == null ? null : f.Name.ToLowerInvariant(); } catch (Exception ex) { AppLogger.Warn("Could not read folder name: " + ex.Message); }
            if (string.IsNullOrEmpty(name)) return false;

            switch (name)
            {
                case "inbox":             // ADD THIS - skip root Inbox
                case "sent items":        // ADD THIS - skip Sent Items
                case "junk e-mail":
                case "junk":
                case "deleted items":
                case "drafts":
                case "outbox":
                case "rss feeds":
                case "conversation action settings":
                case "sync issues":
                case "conflicts":
                case "local failures":
                case "server failures":
                case "search folders":
                    return true;
                default:
                    return false;
            }
        }

        private static string TrimBody(string body, int maxLength)
        {
            if (string.IsNullOrEmpty(body)) return string.Empty;
            if (maxLength <= 0) return string.Empty;
            return body.Length > maxLength ? body.Substring(0, maxLength) : body;
        }

        private static void ReleaseCom(object o)
        {
            if (o == null) return;
            try { Marshal.FinalReleaseComObject(o); } catch { }
        }

        private static void TryAddColumn(Outlook.Table table, string name)
        {
            try { table.Columns.Add(name); }
            catch (Exception ex) { AppLogger.Warn("Could not add scan table column " + name + ": " + ex.Message); }
        }

        // DTOs (no records; C# 7.3-friendly)
        public sealed class SeedRow
        {
            public string EntryId { get; set; }
            public string StoreId { get; set; }
            public string InternetMessageId { get; set; }
            public string ConversationId { get; set; }
            public string FolderPath { get; set; }
            public string Subject { get; set; }
            public string Body { get; set; }
            public string Sender { get; set; }
            public string Domain { get; set; }
            public bool HasAttachments { get; set; }
            public DateTime ReceivedUtc { get; set; }
        }

        public sealed class SeedBody
        {
            public string EntryId { get; private set; }
            public string Body { get; private set; }
            public SeedBody(string entryId, string body)
            {
                EntryId = entryId ?? string.Empty;
                Body = body ?? string.Empty;
            }
        }
    }
}
