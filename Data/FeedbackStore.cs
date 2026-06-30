using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Dapper;
using OutlookClassifierAddIn5.ML;
using OutlookClassifierAddIn5.Services; // for ScanService.SeedRow / SeedBody

namespace OutlookClassifierAddIn5.Data
{
    public class FeedbackStore
    {
        private readonly string _dbPath;
        private readonly string _connStr;

        public FeedbackStore()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookClassifier");
            Directory.CreateDirectory(dir);
            _dbPath = Path.Combine(dir, "store.sqlite");
            _connStr = $"Data Source={_dbPath};Version=3;";
        }

        public void Initialize()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                conn.Open();

                // Pragmas suitable for a local, single-user store
                try
                {
                    conn.Execute("PRAGMA journal_mode=WAL;");
                    conn.Execute("PRAGMA synchronous=NORMAL;");
                    conn.Execute("PRAGMA temp_store=MEMORY;");
                    conn.Execute("PRAGMA foreign_keys=ON;");
                }
                catch { /* pragma failures are non-fatal */ }

                // Core tables
                conn.Execute(@"
                CREATE TABLE IF NOT EXISTS Emails (
                  EntryId TEXT PRIMARY KEY,
                  Folder TEXT,
                  Subject TEXT,
                  Body TEXT,
                  FromAddress TEXT,
                  SenderDomain TEXT,
                  HasAttachments INTEGER,
                  ReceivedUtc TEXT
                );");

                conn.Execute(@"
                CREATE TABLE IF NOT EXISTS Decisions (
                  EntryId TEXT PRIMARY KEY,
                  Predicted TEXT,
                  Chosen TEXT,
                  Confidence REAL,
                  DecidedUtc TEXT
                );");

                conn.Execute(@"CREATE TABLE IF NOT EXISTS Meta (Key TEXT PRIMARY KEY, Value TEXT);");

                // Helpful indexes
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_Folder ON Emails(Folder);");
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_SenderDomain ON Emails(SenderDomain);");
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_ReceivedUtc ON Emails(ReceivedUtc);");
            }
        }

        // --------------------------------------------------------------------
        // Single-row API (kept for compatibility with existing call sites)
        // --------------------------------------------------------------------
        public async Task UpsertEmailAsync(
            string entryId,
            string folder,
            string subject,
            string body,
            string fromAddr,
            string domain,
            bool hasAtt,
            DateTime receivedUtc)
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();

                using (var tx = conn.BeginTransaction())
                {
                    // Body can be very large; avoid unnecessary overwrites
                    var bodyParam = string.IsNullOrEmpty(body) ? null : body;

                    // Try UPDATE first
                    var updated = await conn.ExecuteAsync(@"
                    UPDATE Emails
                       SET Folder         = @Folder,
                           Subject        = COALESCE(@Subject, Subject),
                           Body           = COALESCE(@Body, Body),
                           FromAddress    = COALESCE(@FromAddress, FromAddress),
                           SenderDomain   = COALESCE(@SenderDomain, SenderDomain),
                           HasAttachments = COALESCE(@HasAttachments, HasAttachments),
                           ReceivedUtc    = COALESCE(@ReceivedUtc, ReceivedUtc)
                     WHERE EntryId        = @EntryId;",
                        new
                        {
                            EntryId = entryId,
                            Folder = folder,
                            Subject = subject,
                            Body = bodyParam,
                            FromAddress = fromAddr,
                            SenderDomain = domain,
                            HasAttachments = hasAtt ? 1 : 0,
                            ReceivedUtc = receivedUtc.ToString("o")
                        }, tx);

                    // Insert if not present
                    if (updated == 0)
                    {
                        await conn.ExecuteAsync(@"
                        INSERT INTO Emails
                          (EntryId, Folder, Subject, Body, FromAddress, SenderDomain, HasAttachments, ReceivedUtc)
                        VALUES
                          (@EntryId, @Folder, @Subject, @Body, @FromAddress, @SenderDomain, @HasAttachments, @ReceivedUtc);",
                            new
                            {
                                EntryId = entryId,
                                Folder = folder,
                                Subject = subject,
                                Body = bodyParam,
                                FromAddress = fromAddr,
                                SenderDomain = domain,
                                HasAttachments = hasAtt ? 1 : 0,
                                ReceivedUtc = receivedUtc.ToString("o")
                            }, tx);
                    }

                    tx.Commit();
                }
            }
        }

        // Legacy/basic loader (no filters) — kept for compatibility
        public async Task<IEnumerable<EmailRow>> LoadTrainingRowsAsync()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                return await conn.QueryAsync<EmailRow>(@"
                    SELECT
                      IFNULL(Subject,'')         AS Subject,
                      IFNULL(Body,'')            AS Body,
                      IFNULL(FromAddress,'')     AS FromAddress,
                      IFNULL(SenderDomain,'')    AS SenderDomain,
                      COALESCE(HasAttachments,0) AS HasAttachments,
                      Folder                     AS Label
                    FROM Emails");
            }
        }

        public async Task LogDecisionAsync(string entryId, string predicted, string chosen, double conf)
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();

                using (var tx = conn.BeginTransaction())
                {
                    await conn.ExecuteAsync(@"
                        INSERT OR REPLACE INTO Decisions
                          (EntryId, Predicted, Chosen, Confidence, DecidedUtc)
                        VALUES
                          (@EntryId, @Predicted, @Chosen, @Confidence, @DecidedUtc);",
                        new
                        {
                            EntryId = entryId,
                            Predicted = predicted,
                            Chosen = chosen,
                            Confidence = conf,
                            DecidedUtc = DateTime.UtcNow.ToString("o")
                        }, tx);

                    // Optionally keep Emails.Folder aligned with Chosen
                    if (!string.IsNullOrWhiteSpace(chosen))
                    {
                        await conn.ExecuteAsync(
                            "UPDATE Emails SET Folder = @Chosen WHERE EntryId = @EntryId;",
                            new { Chosen = chosen, EntryId = entryId }, tx);
                    }

                    tx.Commit();
                }
            }
        }

        // --------------------------------------------------------------------
        // Batched APIs used by ScanService
        // --------------------------------------------------------------------
        public Task UpsertEmailsAsync(IEnumerable<ScanService.SeedRow> rows, CancellationToken ct)
        {
            return Task.Run(() =>
            {
                using (var conn = new SQLiteConnection(_connStr))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    using (var update = conn.CreateCommand())
                    using (var insert = conn.CreateCommand())
                    {
                        update.CommandText = @"
UPDATE Emails
   SET Folder         = @Folder,
       Subject        = COALESCE(@Subject, Subject),
       Body           = COALESCE(@Body, Body),
       FromAddress    = COALESCE(@FromAddress, FromAddress),
       SenderDomain   = COALESCE(@SenderDomain, SenderDomain),
       HasAttachments = COALESCE(@HasAttachments, HasAttachments),
       ReceivedUtc    = COALESCE(@ReceivedUtc, ReceivedUtc)
 WHERE EntryId        = @EntryId;";

                        insert.CommandText = @"
INSERT INTO Emails
  (EntryId, Folder, Subject, Body, FromAddress, SenderDomain, HasAttachments, ReceivedUtc)
VALUES
  (@EntryId, @Folder, @Subject, @Body, @FromAddress, @SenderDomain, @HasAttachments, @ReceivedUtc);";

                        var u_id = update.CreateParameter(); u_id.ParameterName = "@EntryId"; update.Parameters.Add(u_id);
                        var u_fld = update.CreateParameter(); u_fld.ParameterName = "@Folder"; update.Parameters.Add(u_fld);
                        var u_sub = update.CreateParameter(); u_sub.ParameterName = "@Subject"; update.Parameters.Add(u_sub);
                        var u_bdy = update.CreateParameter(); u_bdy.ParameterName = "@Body"; update.Parameters.Add(u_bdy);
                        var u_frm = update.CreateParameter(); u_frm.ParameterName = "@FromAddress"; update.Parameters.Add(u_frm);
                        var u_dom = update.CreateParameter(); u_dom.ParameterName = "@SenderDomain"; update.Parameters.Add(u_dom);
                        var u_ha = update.CreateParameter(); u_ha.ParameterName = "@HasAttachments"; update.Parameters.Add(u_ha);
                        var u_rtc = update.CreateParameter(); u_rtc.ParameterName = "@ReceivedUtc"; update.Parameters.Add(u_rtc);

                        var i_id = insert.CreateParameter(); i_id.ParameterName = "@EntryId"; insert.Parameters.Add(i_id);
                        var i_fld = insert.CreateParameter(); i_fld.ParameterName = "@Folder"; insert.Parameters.Add(i_fld);
                        var i_sub = insert.CreateParameter(); i_sub.ParameterName = "@Subject"; insert.Parameters.Add(i_sub);
                        var i_bdy = insert.CreateParameter(); i_bdy.ParameterName = "@Body"; insert.Parameters.Add(i_bdy);
                        var i_frm = insert.CreateParameter(); i_frm.ParameterName = "@FromAddress"; insert.Parameters.Add(i_frm);
                        var i_dom = insert.CreateParameter(); i_dom.ParameterName = "@SenderDomain"; insert.Parameters.Add(i_dom);
                        var i_ha = insert.CreateParameter(); i_ha.ParameterName = "@HasAttachments"; insert.Parameters.Add(i_ha);
                        var i_rtc = insert.CreateParameter(); i_rtc.ParameterName = "@ReceivedUtc"; insert.Parameters.Add(i_rtc);

                        foreach (var r in rows)
                        {
                            ct.ThrowIfCancellationRequested();

                            var entryId = r.EntryId ?? string.Empty;
                            var folder = r.FolderPath ?? string.Empty;   // <-- was r.Folder
                            var subject = r.Subject ?? string.Empty;
                            object body = string.IsNullOrEmpty(r.Body) ? (object)DBNull.Value : r.Body;
                            var from = r.Sender ?? string.Empty;       // <-- was r.FromAddress
                            var domain = r.Domain ?? string.Empty;
                            var ha = r.HasAttachments ? 1 : 0;
                            var rutc = r.ReceivedUtc.ToString("o");

                            u_id.Value = entryId; u_fld.Value = folder; u_sub.Value = subject;
                            u_bdy.Value = body; u_frm.Value = from; u_dom.Value = domain;
                            u_ha.Value = ha; u_rtc.Value = rutc;

                            var updated = update.ExecuteNonQuery();
                            if (updated == 0)
                            {
                                i_id.Value = entryId; i_fld.Value = folder; i_sub.Value = subject;
                                i_bdy.Value = (body is DBNull) ? (object)DBNull.Value : body;
                                i_frm.Value = from; i_dom.Value = domain;
                                i_ha.Value = ha; i_rtc.Value = rutc;

                                insert.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                    }
                }
            }, ct);
        }

        public Task UpsertBodiesAsync(IEnumerable<ScanService.SeedBody> bodies, CancellationToken ct)
        {
            return Task.Run(() =>
            {
                using (var conn = new SQLiteConnection(_connStr))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
UPDATE Emails
   SET Body = COALESCE(@Body, Body)
 WHERE EntryId = @EntryId;";

                        var pId = cmd.CreateParameter(); pId.ParameterName = "@EntryId"; cmd.Parameters.Add(pId);
                        var pBody = cmd.CreateParameter(); pBody.ParameterName = "@Body"; cmd.Parameters.Add(pBody);

                        foreach (var b in bodies)
                        {
                            ct.ThrowIfCancellationRequested();
                            pId.Value = b.EntryId ?? string.Empty;
                            pBody.Value = (object)(b.Body ?? string.Empty) ?? DBNull.Value;
                            cmd.ExecuteNonQuery();
                        }

                        tx.Commit();
                    }
                }
            }, ct);
        }

        // --------------------------------------------------------------------
        // Training data loader with hygiene controls
        // --------------------------------------------------------------------
        public async Task<IEnumerable<EmailRow>> LoadTrainingRowsAsync(int? recentDays = 180, bool excludeDeletedItems = true)
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();

                var whereClauses = new List<string>();

                if (excludeDeletedItems)
                {
                    whereClauses.Add(@"(
                        LOWER(Folder) NOT LIKE '%deleted items%' AND 
                        LOWER(Folder) != 'deleted items'
                    )");
                }

                // Always exclude Inbox root and common system folders
                whereClauses.Add(@"(
                    LOWER(Folder) != 'inbox' AND
                    LOWER(Folder) NOT LIKE '%/inbox' AND
                    LOWER(Folder) NOT LIKE '%sent items%' AND 
                    LOWER(Folder) != 'sent items' AND
                    LOWER(Folder) NOT LIKE '%outbox%' AND 
                    LOWER(Folder) != 'outbox' AND
                    LOWER(Folder) NOT LIKE '%drafts%' AND 
                    LOWER(Folder) != 'drafts' AND
                    LOWER(Folder) NOT LIKE '%junk%' AND
                    LOWER(Folder) NOT LIKE '%calendar%' AND
                    LOWER(Folder) NOT LIKE '%contacts%' AND
                    LOWER(Folder) NOT LIKE '%tasks%' AND
                    LOWER(Folder) NOT LIKE '%notes%' AND
                    LOWER(Folder) NOT LIKE '%journal%'
                )");

                if (recentDays.HasValue && recentDays.Value > 0)
                {
                    var cutoff = DateTime.UtcNow.AddDays(-recentDays.Value).ToString("o");
                    whereClauses.Add($"(ReceivedUtc IS NULL OR ReceivedUtc >= '{cutoff}')");
                }

                var whereClause = whereClauses.Count > 0
                    ? ("WHERE " + string.Join(" AND ", whereClauses))
                    : string.Empty;

                var sql = $@"
                    SELECT
                        IFNULL(Subject,'')         AS Subject,
                        IFNULL(Body,'')            AS Body,
                        IFNULL(FromAddress,'')     AS FromAddress,
                        IFNULL(SenderDomain,'')    AS SenderDomain,
                        COALESCE(HasAttachments,0) AS HasAttachments,
                        Folder                     AS Label
                    FROM Emails
                    {whereClause}
                    ORDER BY ReceivedUtc DESC";

                return await conn.QueryAsync<EmailRow>(sql);
            }
        }

        // --------------------------------------------------------------------
        // Cleanups & metadata
        // --------------------------------------------------------------------
        public async Task<int> CleanupDeletedItemsAsync()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();

                using (var tx = conn.BeginTransaction())
                {
                    var deletedCount = await conn.ExecuteAsync(@"
                        DELETE FROM Emails 
                        WHERE LOWER(Folder) LIKE '%deleted items%' 
                           OR LOWER(Folder) = 'deleted items'", transaction: tx);

                    tx.Commit();
                    return deletedCount;
                }
            }
        }

        public async Task<string> GetMetaValueAsync(string key)
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                return await conn.ExecuteScalarAsync<string>(
                    "SELECT Value FROM Meta WHERE Key = @Key",
                    new { Key = key });
            }
        }

        public async Task SetMetaValueAsync(string key, string value)
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                await conn.ExecuteAsync(
                    "INSERT OR REPLACE INTO Meta (Key, Value) VALUES (@Key, @Value)",
                    new { Key = key, Value = value });
            }
        }

        public async Task<int> CleanupSystemFoldersAsync()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();

                using (var tx = conn.BeginTransaction())
                {
                    var deletedCount = await conn.ExecuteAsync(@"
                        DELETE FROM Emails 
                        WHERE LOWER(Folder) != 'inbox' AND (
                           LOWER(Folder) LIKE '%/inbox' OR
                           LOWER(Folder) LIKE '%sent items%' OR
                           LOWER(Folder) = 'sent items' OR
                           LOWER(Folder) LIKE '%outbox%' OR
                           LOWER(Folder) = 'outbox' OR
                           LOWER(Folder) LIKE '%drafts%' OR
                           LOWER(Folder) = 'drafts' OR
                           LOWER(Folder) LIKE '%junk%' OR
                           LOWER(Folder) LIKE '%calendar%' OR
                           LOWER(Folder) LIKE '%contacts%' OR
                           LOWER(Folder) LIKE '%tasks%' OR
                           LOWER(Folder) LIKE '%notes%' OR
                           LOWER(Folder) LIKE '%journal%')", transaction: tx);

                    tx.Commit();
                    return deletedCount;
                }
            }
        }

        /// <summary>
        /// Purge training rows whose Folder no longer exists in Outlook.
        /// Pass the current verified set of Outlook folder paths (case-insensitive).
        /// </summary>
        public async Task<int> CleanupNonexistentFoldersAsync(IEnumerable<string> validFolderPaths)
        {
            if (validFolderPaths == null) return 0;

            // Normalize & dedupe
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in validFolderPaths)
            {
                if (!string.IsNullOrWhiteSpace(p))
                {
                    var trimmed = p.Trim();
                    if (trimmed.Length > 0) paths.Add(trimmed);
                }
            }
            if (paths.Count == 0) return 0;

            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    // Temp table to hold the valid set
                    using (var drop = conn.CreateCommand())
                    {
                        drop.Transaction = tx;
                        drop.CommandText = "DROP TABLE IF EXISTS ValidFolders;";
                        drop.ExecuteNonQuery();
                    }

                    using (var create = conn.CreateCommand())
                    {
                        create.Transaction = tx;
                        create.CommandText = "CREATE TEMP TABLE ValidFolders(Path TEXT PRIMARY KEY);";
                        create.ExecuteNonQuery();
                    }

                    using (var insert = conn.CreateCommand())
                    {
                        insert.Transaction = tx;
                        insert.CommandText = "INSERT OR IGNORE INTO ValidFolders(Path) VALUES (@p);";
                        var pp = insert.CreateParameter();
                        pp.ParameterName = "@p";
                        insert.Parameters.Add(pp);

                        foreach (var path in paths)
                        {
                            pp.Value = path;
                            insert.ExecuteNonQuery();
                        }
                    }

                    // Case-insensitive comparison via LOWER(...)
                    var deleted = await conn.ExecuteAsync(
                        @"DELETE FROM Emails
                          WHERE Folder IS NOT NULL
                            AND LOWER(Folder) NOT IN (SELECT LOWER(Path) FROM ValidFolders);",
                        transaction: tx);

                    try
                    {
                        using (var drop2 = conn.CreateCommand())
                        {
                            drop2.Transaction = tx;
                            drop2.CommandText = "DROP TABLE IF EXISTS ValidFolders;";
                            drop2.ExecuteNonQuery();
                        }
                    }
                    catch { /* non-fatal */ }

                    tx.Commit();
                    return deleted;
                }
            }
        }

        public async Task<int> NormalizeFolderLabelsAsync(IEnumerable<string> validFullPaths)
        {
            if (validFullPaths == null) return 0;
            var fulls = new List<string>(validFullPaths);
            var lookup = fulls
                .GroupBy(p => p.Split('/').Last(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

            using (var conn = new System.Data.SQLite.SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    // 1) Update unambiguous leaf-only rows to their unique full path
                    // (we’ll do this row-by-row in C# for clarity)
                    var rows = await conn.QueryAsync<(string EntryId, string Folder)>(
                        "SELECT EntryId, Folder FROM Emails WHERE Folder NOT LIKE '%/%';",
                        transaction: tx);

                    int updated = 0, deleted = 0;
                    foreach (var row in rows)
                    {
                        var leaf = row.Folder?.Trim();
                        if (string.IsNullOrEmpty(leaf)) continue;

                        if (lookup.TryGetValue(leaf, out var candidates) && candidates.Count == 1)
                        {
                            // Unambiguous -> update to full path
                            await conn.ExecuteAsync(
                                "UPDATE Emails SET Folder = @Full WHERE EntryId = @Id;",
                                new { Full = candidates[0], Id = row.EntryId }, tx);
                            updated++;
                        }
                        else
                        {
                            // Ambiguous or unknown -> remove to avoid polluting labels
                            await conn.ExecuteAsync(
                                "DELETE FROM Emails WHERE EntryId = @Id;", new { Id = row.EntryId }, tx);
                            deleted++;
                        }
                    }

                    // 2) Normalize backslashes to forward slashes
                    await conn.ExecuteAsync(
                        "UPDATE Emails SET Folder = REPLACE(Folder, '\\\\', '/');", transaction: tx);

                    tx.Commit();
                    return updated + deleted;
                }
            }
        }
    }
}