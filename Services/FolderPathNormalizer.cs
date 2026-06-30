using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookClassifierAddIn5.Services
{
    public static class FolderPathNormalizer
    {
        public static string Normalize(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            return path.Replace('\\', '/').Trim().Trim('/');
        }

        public static IEnumerable<string> NormalizeMany(IEnumerable<string> paths)
        {
            if (paths == null) yield break;
            foreach (var path in paths)
            {
                var normalized = Normalize(path);
                if (normalized.Length > 0) yield return normalized;
            }
        }

        public static string Leaf(string path)
        {
            var normalized = Normalize(path);
            var ix = normalized.LastIndexOf('/');
            return ix >= 0 ? normalized.Substring(ix + 1) : normalized;
        }

        public static string StripAccountRoot(string path)
        {
            var normalized = Normalize(path);
            var ix = normalized.IndexOf('/');
            return ix >= 0 ? normalized.Substring(ix + 1) : normalized;
        }

        public static bool IsRootInbox(string path)
        {
            var lower = Normalize(path).ToLowerInvariant();
            return lower == "inbox" || lower.EndsWith("/inbox");
        }

        public static bool IsDeletedItems(string path)
        {
            return AtOrUnder(Normalize(path).ToLowerInvariant(), "deleted items")
                || AtOrUnder(Normalize(path).ToLowerInvariant(), "trash");
        }

        public static bool IsExcludedSystemFolder(string path)
        {
            var lower = Normalize(path).ToLowerInvariant();
            if (lower.Length == 0) return false;

            if (IsRootInbox(lower)) return true;
            if (AtOrUnder(lower, "sent items") || AtOrUnder(lower, "sent mail")) return true;
            if (AtOrUnder(lower, "deleted items") || AtOrUnder(lower, "trash")) return true;
            if (AtOrUnder(lower, "drafts")) return true;
            if (AtOrUnder(lower, "outbox")) return true;
            if (AtOrUnder(lower, "junk e-mail") || AtOrUnder(lower, "junk email") || AtOrUnder(lower, "junk") || AtOrUnder(lower, "spam")) return true;
            if (AtOrUnder(lower, "archive")) return true;

            if (AtOrUnder(lower, "rss feeds")) return true;
            if (AtOrUnder(lower, "search folders")) return true;
            if (AtOrUnder(lower, "conversation history") || AtOrUnder(lower, "conversation action settings")) return true;
            if (AtOrUnder(lower, "sync issues")) return true;
            if (AtOrUnder(lower, "suggested contacts")) return true;
            if (AtOrUnder(lower, "managed folders")) return true;
            if (AtOrUnder(lower, "all public folders") || AtOrUnder(lower, "public folders")) return true;

            if (AtOrUnder(lower, "calendar")) return true;
            if (AtOrUnder(lower, "contacts")) return true;
            if (AtOrUnder(lower, "tasks")) return true;
            if (AtOrUnder(lower, "notes")) return true;
            if (AtOrUnder(lower, "journal")) return true;
            if (AtOrUnder(lower, "yammer root")) return true;
            if (AtOrUnder(lower, "quick step settings")) return true;
            if (AtOrUnder(lower, "files")) return true;
            if (AtOrUnder(lower, "spam mail")) return true;
            if (AtOrUnder(lower, "social activity notifications")) return true;
            if (AtOrUnder(lower, "clutter")) return true;

            return false;
        }

        public static bool IsCanonicalDuplicateLeaf(string candidate, IEnumerable<string> knownFullPaths)
        {
            var leaf = Leaf(candidate);
            if (leaf.Length == 0 || knownFullPaths == null) return false;

            return knownFullPaths
                .Select(Leaf)
                .Count(x => string.Equals(x, leaf, StringComparison.OrdinalIgnoreCase)) > 1;
        }

        private static bool AtOrUnder(string lower, string segment)
        {
            return lower == segment
                || lower.EndsWith("/" + segment)
                || lower.Contains("/" + segment + "/");
        }
    }
}
