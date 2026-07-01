using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using OutlookClassifierAddIn5.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookClassifierAddIn5.Services
{
    public sealed class FolderSnapshotService
    {
        private readonly object _gate = new object();
        private readonly TimeSpan _ttl = TimeSpan.FromMinutes(5);
        private FolderSnapshot _current;

        public FolderSnapshot Current
        {
            get
            {
                lock (_gate)
                {
                    return _current;
                }
            }
        }

        public FolderSnapshot Refresh(Outlook.Application app, bool force)
        {
            if (app == null) return FolderSnapshot.Empty;

            lock (_gate)
            {
                if (!force && _current != null && DateTime.UtcNow - _current.CreatedUtc < _ttl)
                    return _current;
            }

            var sw = Stopwatch.StartNew();
            var paths = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var p in FolderMap.GetAllFolderPaths(app) ?? Enumerable.Empty<string>())
            {
                var normalized = FolderPathNormalizer.Normalize(p);
                if (normalized.Length == 0) continue;
                if (FolderMap.IsSystemFolder(normalized)) continue;
                if (!FolderMap.IsUserCreatedFolder(normalized)) continue;
                if (seen.Add(normalized)) paths.Add(normalized);
            }

            var defaultRoot = string.Empty;
            try { defaultRoot = app.Session.DefaultStore.DisplayName ?? string.Empty; }
            catch (Exception ex) { AppLogger.Warn("Could not read default store name for folder snapshot: " + ex.Message); }

            var snapshot = new FolderSnapshot(paths, defaultRoot);
            sw.Stop();
            AppLogger.Info("Folder snapshot refreshed. Paths=" + paths.Count + ", elapsedMs=" + sw.ElapsedMilliseconds + ".");

            lock (_gate)
            {
                _current = snapshot;
                return _current;
            }
        }

        public void Invalidate()
        {
            lock (_gate)
            {
                _current = null;
            }
        }
    }

    public sealed class FolderSnapshot
    {
        public static readonly FolderSnapshot Empty = new FolderSnapshot(new List<string>(), string.Empty);

        private readonly Dictionary<string, string> _exact;
        private readonly Dictionary<string, List<string>> _leaf;
        private readonly Dictionary<string, List<string>> _display;

        public FolderSnapshot(IEnumerable<string> fullPaths, string defaultStoreRoot)
        {
            FullPaths = FolderPathNormalizer.NormalizeMany(fullPaths)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToList();

            DefaultStoreRoot = FolderPathNormalizer.Normalize(defaultStoreRoot);
            CreatedUtc = DateTime.UtcNow;
            Version = CreatedUtc.Ticks.ToString();

            _exact = FullPaths.ToDictionary(p => p, p => p, StringComparer.OrdinalIgnoreCase);
            _leaf = FullPaths.GroupBy(FolderPathNormalizer.Leaf, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);
            _display = FullPaths.GroupBy(FolderPathNormalizer.StripAccountRoot, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);
        }

        public List<string> FullPaths { get; private set; }
        public string DefaultStoreRoot { get; private set; }
        public DateTime CreatedUtc { get; private set; }
        public string Version { get; private set; }

        public bool IsValidFilingFolder(string path)
        {
            var normalized = FolderPathNormalizer.Normalize(path);
            return normalized.Length > 0 && _exact.ContainsKey(normalized);
        }

        public string CanonicalizeToFull(string label)
        {
            var s = FolderPathNormalizer.Normalize(label);
            if (s.Length == 0) return string.Empty;

            string exact;
            if (_exact.TryGetValue(s, out exact)) return exact;

            List<string> matches;
            if (_leaf.TryGetValue(s, out matches) && matches.Count == 1)
                return matches[0];

            if (_display.TryGetValue(s, out matches) && matches.Count == 1)
                return matches[0];

            if (_leaf.TryGetValue(s, out matches) && matches.Count > 1)
            {
                var preferred = PreferDefaultStore(matches);
                if (!string.IsNullOrEmpty(preferred)) return preferred;
            }

            if (_display.TryGetValue(s, out matches) && matches.Count > 1)
            {
                var preferred = PreferDefaultStore(matches);
                if (!string.IsNullOrEmpty(preferred)) return preferred;
            }

            return string.Empty;
        }

        private string PreferDefaultStore(List<string> matches)
        {
            if (string.IsNullOrEmpty(DefaultStoreRoot)) return string.Empty;
            return matches.FirstOrDefault(p => p.StartsWith(DefaultStoreRoot + "/", StringComparison.OrdinalIgnoreCase)) ?? string.Empty;
        }
    }
}
