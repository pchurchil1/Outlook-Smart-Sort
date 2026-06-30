using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassifierAddIn5.Data
{
    public static class FolderMap
    {
        public static IEnumerable<string> GetAllFolderPaths(Outlook.Application app)
        {
            var ns = app.Session;
            foreach (Outlook.MAPIFolder root in ns.Folders)
            {
                foreach (var p in Walk(root, root.Name))
                {
                    // Only yield user-created folders that can be used for filing
                    if (IsUserCreatedFolder(p))
                        yield return p;
                }
            }
        }

        private static IEnumerable<string> Walk(Outlook.MAPIFolder f, string path)
        {
            // Yield only if this folder stores MailItem
            if (IsMailFolder(f))
                yield return path;

            // Recurse into children regardless (in case mail subfolders exist under mixed roots)
            foreach (Outlook.MAPIFolder c in f.Folders)
            {
                foreach (var p in Walk(c, path + "/" + c.Name))
                    yield return p;
            }
        }

        private static bool IsMailFolder(Outlook.MAPIFolder f)
        {
            try
            {
                return f.DefaultItemType == Outlook.OlItemType.olMailItem;
            }
            catch { return false; }
        }

        public static bool IsUnderInboxOrDeleted(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            var s = path.ToLowerInvariant();
            return s == "inbox" || s.EndsWith("/inbox")
                || s == "deleted items" || s.EndsWith("/deleted items") || s.Contains("/deleted items/");
        }

        public static bool IsInboxPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            var s = path.ToLowerInvariant();
            return s == "inbox" || s.EndsWith("/inbox");
        }

        public static Outlook.MAPIFolder ResolveFolderByPath(string path, Outlook.NameSpace ns)
        {
            var parts = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            Outlook.MAPIFolder cur = ns.Folders[parts[0]] as Outlook.MAPIFolder;
            for (int i = 1; i < parts.Length; i++)
                cur = cur.Folders[parts[i]] as Outlook.MAPIFolder;
            return cur;
        }

        public static bool IsUserCreatedFolder(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            
            // Exclude system/default folders that shouldn't be used for filing
            if (IsSystemFolder(path)) return false;
            
            // Allow Inbox subfolders (but not root Inbox)
            if (IsUnderInboxOrDeleted(path) && !IsInboxPath(path) && !IsDeletedItemsPath(path))
                return true;
            
            // Allow user-created folders that are outside the standard hierarchy
            if (!IsUnderInboxOrDeleted(path))
                return true;
            
            return false;
        }

        private static bool IsDeletedItemsPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            var lower = path.ToLowerInvariant();
            return lower == "deleted items" || 
                   lower.EndsWith("/deleted items") || 
                   lower.Contains("/deleted items/");
        }

        public static bool IsSystemFolder(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            // Normalize to forward slashes, trim outer slashes, lowercase
            var lower = path.Replace('\\', '/').Trim().Trim('/').ToLowerInvariant();
            if (lower.Length == 0) return false;

            // Segment helpers
            bool EqualsSeg(string name) => lower == name;
            bool EndsWithSeg(string name) => lower.EndsWith("/" + name);
            // "At or under" => the folder itself OR any subfolder of it
            bool AtOrUnder(string name) => EqualsSeg(name) || EndsWithSeg(name) || lower.Contains("/" + name + "/");

            // Inbox (root only; keep subfolders)
            if (EqualsSeg("inbox") || EndsWithSeg("inbox"))
                return true;

            // Core mail system folders (exclude folder AND its subfolders)
            if (AtOrUnder("sent items") || AtOrUnder("sent mail")) return true;
            if (AtOrUnder("deleted items") || AtOrUnder("trash")) return true;
            if (AtOrUnder("drafts")) return true;
            if (AtOrUnder("outbox")) return true;
            if (AtOrUnder("junk e-mail") || AtOrUnder("junk email") || AtOrUnder("junk") || AtOrUnder("spam")) return true;
            if (AtOrUnder("archive")) return true;

            // Special/auto folders (not real filing targets)
            if (AtOrUnder("rss feeds")) return true;
            if (AtOrUnder("search folders")) return true;
            if (AtOrUnder("conversation history") || AtOrUnder("conversation action settings")) return true; // legacy compat
            if (AtOrUnder("sync issues")) return true; // covers Conflicts/Local Failures/Server Failures beneath it
            if (AtOrUnder("suggested contacts")) return true;
            if (AtOrUnder("managed folders")) return true;
            if (AtOrUnder("all public folders") || AtOrUnder("public folders")) return true;

            // Non-mail module roots (exclude if they ever appear in the path list)
            if (AtOrUnder("calendar")) return true;
            if (AtOrUnder("contacts")) return true;
            if (AtOrUnder("tasks")) return true;
            if (AtOrUnder("notes")) return true;
            if (AtOrUnder("journal")) return true;
            if (AtOrUnder("yammer root")) return true;
            if (AtOrUnder("quick step settings")) return true;
            if (AtOrUnder("files")) return true;
            if (AtOrUnder("spam mail")) return true;
            if (AtOrUnder("social activity notifications")) return true;

            // Legacy/optional extras
            if (AtOrUnder("clutter")) return true;

            return false;
        }
    }
}