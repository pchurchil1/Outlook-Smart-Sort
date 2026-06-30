using System;
using System.Collections.Generic;
using System.Linq;
using OutlookClassifierAddIn5.Services;
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
            var s = FolderPathNormalizer.Normalize(path).ToLowerInvariant();
            return s == "inbox" || s.EndsWith("/inbox")
                || s == "deleted items" || s.EndsWith("/deleted items") || s.Contains("/deleted items/");
        }

        public static bool IsInboxPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            var s = FolderPathNormalizer.Normalize(path).ToLowerInvariant();
            return s == "inbox" || s.EndsWith("/inbox");
        }

        public static Outlook.MAPIFolder ResolveFolderByPath(string path, Outlook.NameSpace ns)
        {
            var parts = FolderPathNormalizer.Normalize(path).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return null;
            Outlook.MAPIFolder cur = ns.Folders[parts[0]] as Outlook.MAPIFolder;
            for (int i = 1; i < parts.Length; i++)
                cur = cur.Folders[parts[i]] as Outlook.MAPIFolder;
            return cur;
        }

        public static bool IsUserCreatedFolder(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;

            path = FolderPathNormalizer.Normalize(path);

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
            var lower = FolderPathNormalizer.Normalize(path).ToLowerInvariant();
            return lower == "deleted items" || 
                   lower.EndsWith("/deleted items") || 
                   lower.Contains("/deleted items/");
        }

        public static bool IsSystemFolder(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            return FolderPathNormalizer.IsExcludedSystemFolder(path);
        }
    }
}
