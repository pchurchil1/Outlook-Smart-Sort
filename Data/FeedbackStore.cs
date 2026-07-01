using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Dapper;
using Newtonsoft.Json;
using OutlookClassifierAddIn5.ML;
using OutlookClassifierAddIn5.Services;

namespace OutlookClassifierAddIn5.Data
{
    public sealed class FeedbackExample
    {
        public string OldEntryId { get; set; }
        public string NewEntryId { get; set; }
        public string StoreId { get; set; }
        public string InternetMessageId { get; set; }
        public string ConversationId { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string FromAddress { get; set; }
        public string SenderDomain { get; set; }
        public bool HasAttachments { get; set; }
        public string ChosenFolder { get; set; }
        public string PredictedFolder { get; set; }
        public double Confidence { get; set; }
        public DateTime CreatedUtc { get; set; }
        public DateTime? ReceivedUtc { get; set; }
        public string Source { get; set; }
    }

    public sealed class DecisionRecord
    {
        public string EntryId { get; set; }
        public string OldEntryId { get; set; }
        public string NewEntryId { get; set; }
        public string StoreId { get; set; }
        public string Predicted { get; set; }
        public string Chosen { get; set; }
        public double Confidence { get; set; }
        public DateTime DecidedUtc { get; set; }
        public string Source { get; set; }
    }

    public class FeedbackStore
    {
        private const int CurrentSchemaVersion = 4;
        private const int DefaultBodySnippetLength = 1000;

        private readonly string _dbPath;
        private readonly string _connStr;

        public FeedbackStore()
            : this(null)
        {
        }

        public FeedbackStore(string dbPath)
        {
            _dbPath = string.IsNullOrWhiteSpace(dbPath)
                ? Path.Combine(AppLogger.DataDirectory, "store.sqlite")
                : dbPath;
            var dir = Path.GetDirectoryName(_dbPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            _connStr = "Data Source=" + _dbPath + ";Version=3;";
        }

        public string DatabasePath
        {
            get { return _dbPath; }
        }

        public string DataFolder
        {
            get { return AppLogger.DataDirectory; }
        }

        public void Initialize()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                conn.Open();
                ExecutePragmas(conn);

                using (var tx = conn.BeginTransaction())
                {
                    conn.Execute(@"CREATE TABLE IF NOT EXISTS Meta (Key TEXT PRIMARY KEY, Value TEXT NOT NULL);", transaction: tx);
                    tx.Commit();
                }

                try
                {
                    var version = GetSchemaVersion(conn);
                    AppLogger.Info("Database schema version before migrations: " + version);

                    if (version < 1) Migrate0To1(conn);
                    if (version < 2) Migrate1To2(conn);
                    if (version < 3) Migrate2To3(conn);
                    if (version < 4) Migrate3To4(conn);

                    AppLogger.Info("Database schema ready at version " + CurrentSchemaVersion + ".");
                }
                catch (Exception ex)
                {
                    AppLogger.Error(ex, "Schema migration failed.");
                    throw;
                }
            }
        }

        private static void ExecutePragmas(SQLiteConnection conn)
        {
            try
            {
                conn.Execute("PRAGMA journal_mode=WAL;");
                conn.Execute("PRAGMA synchronous=NORMAL;");
                conn.Execute("PRAGMA temp_store=MEMORY;");
                conn.Execute("PRAGMA foreign_keys=ON;");
            }
            catch (Exception ex)
            {
                AppLogger.Warn("SQLite pragma setup failed: " + ex.Message);
            }
        }

        private static int GetSchemaVersion(SQLiteConnection conn)
        {
            var value = conn.ExecuteScalar<string>("SELECT Value FROM Meta WHERE Key = 'SchemaVersion';");
            int version;
            return int.TryParse(value, out version) ? version : 0;
        }

        private static void SetSchemaVersion(SQLiteConnection conn, SQLiteTransaction tx, int version)
        {
            conn.Execute(
                "INSERT OR REPLACE INTO Meta (Key, Value) VALUES ('SchemaVersion', @Value);",
                new { Value = version.ToString() },
                tx);
        }

        private static void Migrate0To1(SQLiteConnection conn)
        {
            using (var tx = conn.BeginTransaction())
            {
                AppLogger.Info("Applying schema migration 0 -> 1.");

                conn.Execute(@"
CREATE TABLE IF NOT EXISTS Emails (
  EntryId TEXT PRIMARY KEY,
  StoreId TEXT,
  InternetMessageId TEXT,
  ConversationId TEXT,
  NormalizedSubjectHash TEXT,
  Folder TEXT,
  Subject TEXT,
  Body TEXT,
  FromAddress TEXT,
  SenderDomain TEXT,
  HasAttachments INTEGER,
  ReceivedUtc TEXT
);", transaction: tx);

                conn.Execute(@"
CREATE TABLE IF NOT EXISTS Decisions (
  EntryId TEXT PRIMARY KEY,
  OldEntryId TEXT,
  NewEntryId TEXT,
  StoreId TEXT,
  Predicted TEXT,
  Chosen TEXT,
  Confidence REAL,
  DecidedUtc TEXT,
  Source TEXT
);", transaction: tx);

                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_Folder ON Emails(Folder);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_SenderDomain ON Emails(SenderDomain);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_ReceivedUtc ON Emails(ReceivedUtc);", transaction: tx);

                SetSchemaVersion(conn, tx, 1);
                tx.Commit();
            }
        }

        private static void Migrate1To2(SQLiteConnection conn)
        {
            using (var tx = conn.BeginTransaction())
            {
                AppLogger.Info("Applying schema migration 1 -> 2.");

                AddColumnIfMissing(conn, tx, "Emails", "StoreId", "TEXT");
                AddColumnIfMissing(conn, tx, "Emails", "InternetMessageId", "TEXT");
                AddColumnIfMissing(conn, tx, "Emails", "ConversationId", "TEXT");
                AddColumnIfMissing(conn, tx, "Emails", "NormalizedSubjectHash", "TEXT");

                AddColumnIfMissing(conn, tx, "Decisions", "OldEntryId", "TEXT");
                AddColumnIfMissing(conn, tx, "Decisions", "NewEntryId", "TEXT");
                AddColumnIfMissing(conn, tx, "Decisions", "StoreId", "TEXT");
                AddColumnIfMissing(conn, tx, "Decisions", "Source", "TEXT");

                conn.Execute(@"
CREATE TABLE IF NOT EXISTS FeedbackExamples (
  Id INTEGER PRIMARY KEY AUTOINCREMENT,
  OldEntryId TEXT,
  NewEntryId TEXT,
  StoreId TEXT,
  InternetMessageId TEXT,
  ConversationId TEXT,
  NormalizedSubjectHash TEXT,
  Subject TEXT,
  Body TEXT,
  FromAddress TEXT,
  SenderDomain TEXT,
  HasAttachments INTEGER,
  ChosenFolder TEXT NOT NULL,
  PredictedFolder TEXT,
  Confidence REAL,
  CreatedUtc TEXT NOT NULL,
  ReceivedUtc TEXT,
  Source TEXT NOT NULL
);", transaction: tx);

                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_FeedbackExamples_ChosenFolder ON FeedbackExamples(ChosenFolder);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_FeedbackExamples_NewEntryId ON FeedbackExamples(NewEntryId);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_FeedbackExamples_InternetMessageId ON FeedbackExamples(InternetMessageId);", transaction: tx);

                SetSchemaVersion(conn, tx, 2);
                tx.Commit();
            }
        }

        private static void Migrate2To3(SQLiteConnection conn)
        {
            using (var tx = conn.BeginTransaction())
            {
                AppLogger.Info("Applying schema migration 2 -> 3.");
                NormalizeStoredFolderPaths(conn, tx);
                SetSchemaVersion(conn, tx, 3);
                tx.Commit();
            }
        }

        private static void Migrate3To4(SQLiteConnection conn)
        {
            using (var tx = conn.BeginTransaction())
            {
                AppLogger.Info("Applying schema migration 3 -> 4.");
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_InternetMessageId ON Emails(InternetMessageId);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_ConversationId ON Emails(ConversationId);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_Emails_NormalizedSubjectHash ON Emails(NormalizedSubjectHash);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_FeedbackExamples_CreatedUtc ON FeedbackExamples(CreatedUtc);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_FeedbackExamples_ReceivedUtc ON FeedbackExamples(ReceivedUtc);", transaction: tx);
                conn.Execute(@"CREATE INDEX IF NOT EXISTS IX_FeedbackExamples_ConversationId ON FeedbackExamples(ConversationId);", transaction: tx);
                SetSchemaVersion(conn, tx, 4);
                tx.Commit();
            }
        }

        private static void AddColumnIfMissing(SQLiteConnection conn, SQLiteTransaction tx, string table, string column, string definition)
        {
            if (ColumnExists(conn, table, column)) return;
            conn.Execute("ALTER TABLE " + table + " ADD COLUMN " + column + " " + definition + ";", transaction: tx);
        }

        private static bool ColumnExists(SQLiteConnection conn, string table, string column)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "PRAGMA table_info(" + table + ");";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var name = reader["name"] as string;
                        if (string.Equals(name, column, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }

            return false;
        }

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
                    var normalizedFolder = FolderPathNormalizer.Normalize(folder);
                    var normalizedSubjectHash = OutlookMailSnapshotService.NormalizedSubjectHash(subject);
                    var bodyParam = string.IsNullOrEmpty(body) ? null : body;

                    var updated = await conn.ExecuteAsync(@"
UPDATE Emails
   SET Folder                = @Folder,
       Subject               = COALESCE(@Subject, Subject),
       Body                  = COALESCE(@Body, Body),
       FromAddress           = COALESCE(@FromAddress, FromAddress),
       SenderDomain          = COALESCE(@SenderDomain, SenderDomain),
       HasAttachments        = COALESCE(@HasAttachments, HasAttachments),
       ReceivedUtc           = COALESCE(@ReceivedUtc, ReceivedUtc),
       NormalizedSubjectHash = COALESCE(@NormalizedSubjectHash, NormalizedSubjectHash)
 WHERE EntryId               = @EntryId;",
                        new
                        {
                            EntryId = entryId ?? string.Empty,
                            Folder = normalizedFolder,
                            Subject = subject ?? string.Empty,
                            Body = bodyParam,
                            FromAddress = fromAddr ?? string.Empty,
                            SenderDomain = string.IsNullOrWhiteSpace(domain) ? SenderResolutionService.ExtractDomain(fromAddr) : domain,
                            HasAttachments = hasAtt ? 1 : 0,
                            ReceivedUtc = receivedUtc.ToString("o"),
                            NormalizedSubjectHash = normalizedSubjectHash
                        }, tx);

                    if (updated == 0)
                    {
                        await conn.ExecuteAsync(@"
INSERT INTO Emails
  (EntryId, Folder, Subject, Body, FromAddress, SenderDomain, HasAttachments, ReceivedUtc, NormalizedSubjectHash)
VALUES
  (@EntryId, @Folder, @Subject, @Body, @FromAddress, @SenderDomain, @HasAttachments, @ReceivedUtc, @NormalizedSubjectHash);",
                            new
                            {
                                EntryId = entryId ?? string.Empty,
                                Folder = normalizedFolder,
                                Subject = subject ?? string.Empty,
                                Body = bodyParam,
                                FromAddress = fromAddr ?? string.Empty,
                                SenderDomain = string.IsNullOrWhiteSpace(domain) ? SenderResolutionService.ExtractDomain(fromAddr) : domain,
                                HasAttachments = hasAtt ? 1 : 0,
                                ReceivedUtc = receivedUtc.ToString("o"),
                                NormalizedSubjectHash = normalizedSubjectHash
                            }, tx);
                    }

                    tx.Commit();
                }
            }
        }

        public async Task<IEnumerable<EmailRow>> LoadTrainingRowsAsync()
        {
            var rows = await LoadTrainingExamplesAsync(null, false);
            return rows.Select(ToEmailRow).ToList();
        }

        public async Task<IEnumerable<EmailRow>> LoadTrainingRowsAsync(int? recentDays = 180, bool excludeDeletedItems = true)
        {
            var rows = await LoadTrainingExamplesAsync(recentDays, excludeDeletedItems);
            return rows.Select(ToEmailRow).ToList();
        }

        public async Task<List<TrainingExample>> LoadTrainingExamplesAsync(int? recentDays = 180, bool excludeSystemFolders = true, bool includeBody = true)
        {
            var sw = Stopwatch.StartNew();
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();

                var cutoff = recentDays.HasValue && recentDays.Value > 0
                    ? DateTime.UtcNow.AddDays(-recentDays.Value).ToString("o")
                    : null;

                var bodySelect = includeBody ? "IFNULL(Body,'') AS Body" : "'' AS Body";
                var emailSystemFilter = excludeSystemFolders ? (" AND " + SystemFolderSqlPredicate("Folder")) : string.Empty;
                var feedbackSystemFilter = excludeSystemFolders ? (" AND " + SystemFolderSqlPredicate("ChosenFolder")) : string.Empty;

                var emailSql = @"
SELECT
  IFNULL(Subject,'') AS Subject,
  " + bodySelect + @",
  IFNULL(FromAddress,'') AS FromAddress,
  IFNULL(SenderDomain,'') AS SenderDomain,
  COALESCE(HasAttachments,0) AS HasAttachments,
  IFNULL(Folder,'') AS Label,
  ReceivedUtc AS ReceivedUtcText,
  'Emails' AS Source
FROM Emails
WHERE Folder IS NOT NULL
  AND (@Cutoff IS NULL OR ReceivedUtc IS NULL OR ReceivedUtc >= @Cutoff)" + emailSystemFilter + @";";

                var feedbackSql = @"
SELECT
  IFNULL(Subject,'') AS Subject,
  " + bodySelect + @",
  IFNULL(FromAddress,'') AS FromAddress,
  IFNULL(SenderDomain,'') AS SenderDomain,
  COALESCE(HasAttachments,0) AS HasAttachments,
  IFNULL(ChosenFolder,'') AS Label,
  ReceivedUtc AS ReceivedUtcText,
  IFNULL(Source,'Feedback') AS Source,
  CreatedUtc AS CreatedUtcText
FROM FeedbackExamples
WHERE ChosenFolder IS NOT NULL
  AND (@Cutoff IS NULL OR ReceivedUtc IS NULL OR ReceivedUtc >= @Cutoff OR CreatedUtc >= @Cutoff)" + feedbackSystemFilter + @";";

                var records = new List<TrainingRecord>();
                records.AddRange(await conn.QueryAsync<TrainingRecord>(emailSql, new { Cutoff = cutoff }));
                records.AddRange(await conn.QueryAsync<TrainingRecord>(feedbackSql, new { Cutoff = cutoff }));

                var examples = new List<TrainingExample>();
                foreach (var record in records)
                {
                    var label = FolderPathNormalizer.Normalize(record.Label);
                    if (label.Length == 0) continue;
                    if (excludeSystemFolders && FolderPathNormalizer.IsExcludedSystemFolder(label)) continue;

                    var from = record.FromAddress ?? string.Empty;
                    var domain = string.IsNullOrWhiteSpace(record.SenderDomain)
                        ? SenderResolutionService.ExtractDomain(from)
                        : record.SenderDomain;

                    examples.Add(new TrainingExample
                    {
                        Subject = record.Subject ?? string.Empty,
                        Body = record.Body ?? string.Empty,
                        FromAddress = from,
                        SenderDomain = domain,
                        HasAttachments = record.HasAttachments != 0,
                        Label = label,
                        ReceivedUtc = ParseUtc(record.ReceivedUtcText) ?? ParseUtc(record.CreatedUtcText),
                        Source = record.Source ?? string.Empty
                    });
                }

                var result = examples
                    .OrderByDescending(r => r.ReceivedUtc.HasValue ? r.ReceivedUtc.Value : DateTime.MinValue)
                    .ToList();
                sw.Stop();
                AppLogger.Info("Training rows loaded. Rows=" + result.Count + ", includeBody=" + includeBody + ", elapsedMs=" + sw.ElapsedMilliseconds + ".");
                return result;
            }
        }

        public async Task LogDecisionAsync(string entryId, string predicted, string chosen, double conf)
        {
            var decision = new DecisionRecord
            {
                EntryId = entryId ?? string.Empty,
                OldEntryId = entryId ?? string.Empty,
                NewEntryId = entryId ?? string.Empty,
                Predicted = predicted ?? string.Empty,
                Chosen = chosen ?? string.Empty,
                Confidence = conf,
                DecidedUtc = DateTime.UtcNow,
                Source = "Decision"
            };

            await RecordApprovalsAsync(null, new[] { decision }, CancellationToken.None);
        }

        public Task UpsertEmailsAsync(IEnumerable<ScanService.SeedRow> rows, CancellationToken ct)
        {
            var materialized = (rows ?? Enumerable.Empty<ScanService.SeedRow>()).ToList();
            return Task.Run(() =>
            {
                var sw = Stopwatch.StartNew();
                using (var conn = new SQLiteConnection(_connStr))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    using (var update = conn.CreateCommand())
                    using (var insert = conn.CreateCommand())
                    {
                        update.Transaction = tx;
                        insert.Transaction = tx;

                        update.CommandText = @"
UPDATE Emails
   SET StoreId               = COALESCE(@StoreId, StoreId),
       InternetMessageId     = COALESCE(@InternetMessageId, InternetMessageId),
       ConversationId        = COALESCE(@ConversationId, ConversationId),
       NormalizedSubjectHash = COALESCE(@NormalizedSubjectHash, NormalizedSubjectHash),
       Folder                = @Folder,
       Subject               = COALESCE(@Subject, Subject),
       Body                  = COALESCE(@Body, Body),
       FromAddress           = COALESCE(@FromAddress, FromAddress),
       SenderDomain          = COALESCE(@SenderDomain, SenderDomain),
       HasAttachments        = COALESCE(@HasAttachments, HasAttachments),
       ReceivedUtc           = COALESCE(@ReceivedUtc, ReceivedUtc)
 WHERE EntryId               = @EntryId;";

                        insert.CommandText = @"
INSERT INTO Emails
  (EntryId, StoreId, InternetMessageId, ConversationId, NormalizedSubjectHash, Folder, Subject, Body, FromAddress, SenderDomain, HasAttachments, ReceivedUtc)
VALUES
  (@EntryId, @StoreId, @InternetMessageId, @ConversationId, @NormalizedSubjectHash, @Folder, @Subject, @Body, @FromAddress, @SenderDomain, @HasAttachments, @ReceivedUtc);";

                        AddParameters(update, "@EntryId", "@StoreId", "@InternetMessageId", "@ConversationId", "@NormalizedSubjectHash", "@Folder", "@Subject", "@Body", "@FromAddress", "@SenderDomain", "@HasAttachments", "@ReceivedUtc");
                        AddParameters(insert, "@EntryId", "@StoreId", "@InternetMessageId", "@ConversationId", "@NormalizedSubjectHash", "@Folder", "@Subject", "@Body", "@FromAddress", "@SenderDomain", "@HasAttachments", "@ReceivedUtc");

                        foreach (var r in materialized)
                        {
                            ct.ThrowIfCancellationRequested();

                            var entryId = r.EntryId ?? string.Empty;
                            var folder = FolderPathNormalizer.Normalize(r.FolderPath);
                            var subject = r.Subject ?? string.Empty;
                            var from = r.Sender ?? string.Empty;
                            var domain = string.IsNullOrWhiteSpace(r.Domain) ? SenderResolutionService.ExtractDomain(from) : r.Domain;

                            SetParameterValues(update, entryId, r.StoreId, r.InternetMessageId, r.ConversationId,
                                OutlookMailSnapshotService.NormalizedSubjectHash(subject), folder, subject,
                                string.IsNullOrEmpty(r.Body) ? (object)DBNull.Value : r.Body, from, domain,
                                r.HasAttachments ? 1 : 0, r.ReceivedUtc.ToString("o"));

                            var updated = update.ExecuteNonQuery();
                            if (updated == 0)
                            {
                                SetParameterValues(insert, entryId, r.StoreId, r.InternetMessageId, r.ConversationId,
                                    OutlookMailSnapshotService.NormalizedSubjectHash(subject), folder, subject,
                                    string.IsNullOrEmpty(r.Body) ? (object)DBNull.Value : r.Body, from, domain,
                                    r.HasAttachments ? 1 : 0, r.ReceivedUtc.ToString("o"));
                                insert.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                    }
                }
                sw.Stop();
                AppLogger.Info("Email batch upsert complete. Rows=" + materialized.Count + ", elapsedMs=" + sw.ElapsedMilliseconds + ".");
            }, ct);
        }

        public Task UpsertBodiesAsync(IEnumerable<ScanService.SeedBody> bodies, CancellationToken ct)
        {
            var materialized = (bodies ?? Enumerable.Empty<ScanService.SeedBody>()).ToList();
            return Task.Run(() =>
            {
                var sw = Stopwatch.StartNew();
                using (var conn = new SQLiteConnection(_connStr))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"
UPDATE Emails
   SET Body = COALESCE(@Body, Body)
 WHERE EntryId = @EntryId;";

                        var pId = cmd.CreateParameter(); pId.ParameterName = "@EntryId"; cmd.Parameters.Add(pId);
                        var pBody = cmd.CreateParameter(); pBody.ParameterName = "@Body"; cmd.Parameters.Add(pBody);

                        foreach (var b in materialized)
                        {
                            ct.ThrowIfCancellationRequested();
                            pId.Value = b.EntryId ?? string.Empty;
                            pBody.Value = string.IsNullOrEmpty(b.Body) ? (object)DBNull.Value : b.Body;
                            cmd.ExecuteNonQuery();
                        }

                        tx.Commit();
                    }
                }
                sw.Stop();
                AppLogger.Info("Body batch upsert complete. Rows=" + materialized.Count + ", elapsedMs=" + sw.ElapsedMilliseconds + ".");
            }, ct);
        }

        public Task RecordApprovalAsync(FeedbackExample feedback, DecisionRecord decision, CancellationToken ct)
        {
            return RecordApprovalsAsync(
                feedback == null ? null : new[] { feedback },
                decision == null ? null : new[] { decision },
                ct);
        }

        public Task RecordApprovalsAsync(IEnumerable<FeedbackExample> feedbackExamples, IEnumerable<DecisionRecord> decisions, CancellationToken ct)
        {
            var feedbackList = (feedbackExamples ?? Enumerable.Empty<FeedbackExample>()).ToList();
            var decisionList = (decisions ?? Enumerable.Empty<DecisionRecord>()).ToList();
            return Task.Run(() =>
            {
                var sw = Stopwatch.StartNew();
                using (var conn = new SQLiteConnection(_connStr))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    {
                        foreach (var decision in decisionList)
                        {
                            ct.ThrowIfCancellationRequested();
                            var chosen = FolderPathNormalizer.Normalize(decision.Chosen);
                            var predicted = FolderPathNormalizer.Normalize(decision.Predicted);
                            var newId = decision.NewEntryId ?? decision.EntryId ?? string.Empty;
                            var oldId = decision.OldEntryId ?? decision.EntryId ?? string.Empty;

                            conn.Execute(@"
INSERT OR REPLACE INTO Decisions
  (EntryId, OldEntryId, NewEntryId, StoreId, Predicted, Chosen, Confidence, DecidedUtc, Source)
VALUES
  (@EntryId, @OldEntryId, @NewEntryId, @StoreId, @Predicted, @Chosen, @Confidence, @DecidedUtc, @Source);",
                                new
                                {
                                    EntryId = string.IsNullOrEmpty(newId) ? oldId : newId,
                                    OldEntryId = oldId,
                                    NewEntryId = newId,
                                    StoreId = decision.StoreId ?? string.Empty,
                                    Predicted = predicted,
                                    Chosen = chosen,
                                    Confidence = decision.Confidence,
                                    DecidedUtc = (decision.DecidedUtc == default(DateTime) ? DateTime.UtcNow : decision.DecidedUtc).ToString("o"),
                                    Source = decision.Source ?? "Decision"
                                }, tx);

                            if (!string.IsNullOrWhiteSpace(chosen))
                            {
                                conn.Execute("UPDATE Emails SET Folder = @Chosen WHERE EntryId = @EntryId;",
                                    new { Chosen = chosen, EntryId = newId }, tx);
                                if (!string.Equals(oldId, newId, StringComparison.Ordinal))
                                {
                                    conn.Execute("UPDATE Emails SET Folder = @Chosen WHERE EntryId = @EntryId;",
                                        new { Chosen = chosen, EntryId = oldId }, tx);
                                }
                            }
                        }

                        foreach (var feedback in feedbackList)
                        {
                            ct.ThrowIfCancellationRequested();
                            InsertFeedbackExample(conn, tx, feedback);
                        }

                        tx.Commit();
                    }
                }
                sw.Stop();
                AppLogger.Info("Approval records written. Feedback=" + feedbackList.Count + ", decisions=" + decisionList.Count + ", elapsedMs=" + sw.ElapsedMilliseconds + ".");
            }, ct);
        }

        public async Task<int> CleanupDeletedItemsAsync()
        {
            return await CleanupSystemFoldersAsync();
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
                    new { Key = key, Value = value ?? string.Empty });
            }
        }

        public async Task<int> CleanupSystemFoldersAsync()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    var deleted = DeleteSystemLabels(conn, tx, "Emails", "EntryId", "Folder");
                    deleted += DeleteSystemLabels(conn, tx, "FeedbackExamples", "Id", "ChosenFolder");
                    tx.Commit();
                    return deleted;
                }
            }
        }

        public async Task<int> CleanupNonexistentFoldersAsync(IEnumerable<string> validFolderPaths)
        {
            var paths = new HashSet<string>(
                FolderPathNormalizer.NormalizeMany(validFolderPaths),
                StringComparer.OrdinalIgnoreCase);

            if (paths.Count == 0) return 0;

            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    var deleted = DeleteLabelsNotInSet(conn, tx, "Emails", "EntryId", "Folder", paths);
                    deleted += DeleteLabelsNotInSet(conn, tx, "FeedbackExamples", "Id", "ChosenFolder", paths);
                    tx.Commit();
                    return deleted;
                }
            }
        }

        public async Task<int> NormalizeFolderLabelsAsync(IEnumerable<string> validFullPaths)
        {
            var normalizedFullPaths = FolderPathNormalizer.NormalizeMany(validFullPaths).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var lookup = normalizedFullPaths
                .GroupBy(FolderPathNormalizer.Leaf, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    var changed = 0;
                    changed += NormalizeLabelColumn(conn, tx, "Emails", "EntryId", "Folder", lookup);
                    changed += NormalizeLabelColumn(conn, tx, "FeedbackExamples", "Id", "ChosenFolder", lookup);
                    changed += NormalizeLabelColumn(conn, tx, "FeedbackExamples", "Id", "PredictedFolder", lookup);
                    changed += NormalizeLabelColumn(conn, tx, "Decisions", "EntryId", "Predicted", lookup);
                    changed += NormalizeLabelColumn(conn, tx, "Decisions", "EntryId", "Chosen", lookup);
                    tx.Commit();
                    return changed;
                }
            }
        }

        public async Task<int> ClearTrainingDatabaseAsync()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    var count = await conn.ExecuteAsync("DELETE FROM Emails;", transaction: tx);
                    count += await conn.ExecuteAsync("DELETE FROM FeedbackExamples;", transaction: tx);
                    tx.Commit();
                    AppLogger.Warn("Training database cleared by user action.");
                    return count;
                }
            }
        }

        public async Task<int> ClearDecisionAndFeedbackHistoryAsync()
        {
            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    var count = await conn.ExecuteAsync("DELETE FROM Decisions;", transaction: tx);
                    count += await conn.ExecuteAsync("DELETE FROM FeedbackExamples;", transaction: tx);
                    tx.Commit();
                    AppLogger.Warn("Decision and feedback history cleared by user action.");
                    return count;
                }
            }
        }

        public async Task<int> GetBodySnippetLengthAsync()
        {
            var value = await GetMetaValueAsync("BodySnippetLength");
            int length;
            return int.TryParse(value, out length) && length >= 0 ? length : DefaultBodySnippetLength;
        }

        public Task SetBodySnippetLengthAsync(int length)
        {
            if (length < 0) length = 0;
            if (length > 5000) length = 5000;
            return SetMetaValueAsync("BodySnippetLength", length.ToString());
        }

        public async Task<string> ExportDiagnosticsAsync(string targetPath)
        {
            if (string.IsNullOrWhiteSpace(targetPath))
                targetPath = Path.Combine(AppLogger.DataDirectory, "diagnostics-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + ".json");

            using (var conn = new SQLiteConnection(_connStr))
            {
                await conn.OpenAsync();
                var diagnostics = new
                {
                    ExportedUtc = DateTime.UtcNow.ToString("o"),
                    DataFolder = AppLogger.DataDirectory,
                    SchemaVersion = await conn.ExecuteScalarAsync<string>("SELECT Value FROM Meta WHERE Key = 'SchemaVersion';"),
                    EmailRows = await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM Emails;"),
                    FeedbackRows = await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM FeedbackExamples;"),
                    DecisionRows = await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM Decisions;"),
                    LabelCount = await conn.ExecuteScalarAsync<int>("SELECT COUNT(DISTINCT Folder) FROM Emails WHERE Folder IS NOT NULL;"),
                    FeedbackLabelCount = await conn.ExecuteScalarAsync<int>("SELECT COUNT(DISTINCT ChosenFolder) FROM FeedbackExamples WHERE ChosenFolder IS NOT NULL;")
                };

                File.WriteAllText(targetPath, JsonConvert.SerializeObject(diagnostics, Formatting.Indented));
                return targetPath;
            }
        }

        private static void InsertFeedbackExample(SQLiteConnection conn, SQLiteTransaction tx, FeedbackExample feedback)
        {
            if (feedback == null) return;

            var chosen = FolderPathNormalizer.Normalize(feedback.ChosenFolder);
            if (string.IsNullOrEmpty(chosen)) return;

            var predicted = FolderPathNormalizer.Normalize(feedback.PredictedFolder);
            var subject = feedback.Subject ?? string.Empty;
            var from = feedback.FromAddress ?? string.Empty;
            var domain = string.IsNullOrWhiteSpace(feedback.SenderDomain)
                ? SenderResolutionService.ExtractDomain(from)
                : feedback.SenderDomain;

            conn.Execute(@"
INSERT INTO FeedbackExamples
  (OldEntryId, NewEntryId, StoreId, InternetMessageId, ConversationId, NormalizedSubjectHash, Subject, Body,
   FromAddress, SenderDomain, HasAttachments, ChosenFolder, PredictedFolder, Confidence, CreatedUtc, ReceivedUtc, Source)
VALUES
  (@OldEntryId, @NewEntryId, @StoreId, @InternetMessageId, @ConversationId, @NormalizedSubjectHash, @Subject, @Body,
   @FromAddress, @SenderDomain, @HasAttachments, @ChosenFolder, @PredictedFolder, @Confidence, @CreatedUtc, @ReceivedUtc, @Source);",
                new
                {
                    OldEntryId = feedback.OldEntryId ?? string.Empty,
                    NewEntryId = feedback.NewEntryId ?? string.Empty,
                    StoreId = feedback.StoreId ?? string.Empty,
                    InternetMessageId = feedback.InternetMessageId ?? string.Empty,
                    ConversationId = feedback.ConversationId ?? string.Empty,
                    NormalizedSubjectHash = OutlookMailSnapshotService.NormalizedSubjectHash(subject),
                    Subject = subject,
                    Body = feedback.Body ?? string.Empty,
                    FromAddress = from,
                    SenderDomain = domain,
                    HasAttachments = feedback.HasAttachments ? 1 : 0,
                    ChosenFolder = chosen,
                    PredictedFolder = predicted,
                    Confidence = feedback.Confidence,
                    CreatedUtc = (feedback.CreatedUtc == default(DateTime) ? DateTime.UtcNow : feedback.CreatedUtc).ToString("o"),
                    ReceivedUtc = feedback.ReceivedUtc.HasValue ? feedback.ReceivedUtc.Value.ToUniversalTime().ToString("o") : null,
                    Source = string.IsNullOrWhiteSpace(feedback.Source) ? "Feedback" : feedback.Source
                }, tx);
        }

        private static string SystemFolderSqlPredicate(string column)
        {
            var normalized = "LOWER(REPLACE(" + column + ", '\\\\', '/'))";
            return "(" +
                   normalized + " <> 'inbox' AND " +
                   normalized + " <> 'sent items' AND " +
                   normalized + " <> 'sent mail' AND " +
                   normalized + " <> 'deleted items' AND " +
                   normalized + " <> 'trash' AND " +
                   normalized + " <> 'drafts' AND " +
                   normalized + " <> 'outbox' AND " +
                   normalized + " <> 'junk' AND " +
                   normalized + " <> 'junk e-mail' AND " +
                   normalized + " <> 'junk email' AND " +
                   normalized + " <> 'spam' AND " +
                   normalized + " <> 'archive' AND " +
                   normalized + " <> 'calendar' AND " +
                   normalized + " <> 'contacts' AND " +
                   normalized + " <> 'tasks' AND " +
                   normalized + " <> 'notes' AND " +
                   normalized + " <> 'journal' AND " +
                   normalized + " NOT LIKE '%/inbox' AND " +
                   normalized + " NOT LIKE '%/sent items%' AND " +
                   normalized + " NOT LIKE '%/sent mail%' AND " +
                   normalized + " NOT LIKE '%/deleted items%' AND " +
                   normalized + " NOT LIKE '%/trash%' AND " +
                   normalized + " NOT LIKE '%/drafts%' AND " +
                   normalized + " NOT LIKE '%/outbox%' AND " +
                   normalized + " NOT LIKE '%/junk%' AND " +
                   normalized + " NOT LIKE '%/junk e-mail%' AND " +
                   normalized + " NOT LIKE '%/junk email%' AND " +
                   normalized + " NOT LIKE '%/spam%' AND " +
                   normalized + " NOT LIKE '%/archive%' AND " +
                   normalized + " NOT LIKE '%/calendar%' AND " +
                   normalized + " NOT LIKE '%/contacts%' AND " +
                   normalized + " NOT LIKE '%/tasks%' AND " +
                   normalized + " NOT LIKE '%/notes%' AND " +
                   normalized + " NOT LIKE '%/journal%'" +
                   ")";
        }

        private static void NormalizeStoredFolderPaths(SQLiteConnection conn, SQLiteTransaction tx)
        {
            NormalizeLabelColumn(conn, tx, "Emails", "EntryId", "Folder", null);
            NormalizeLabelColumn(conn, tx, "FeedbackExamples", "Id", "ChosenFolder", null);
            NormalizeLabelColumn(conn, tx, "FeedbackExamples", "Id", "PredictedFolder", null);
            NormalizeLabelColumn(conn, tx, "Decisions", "EntryId", "Predicted", null);
            NormalizeLabelColumn(conn, tx, "Decisions", "EntryId", "Chosen", null);
        }

        private static int NormalizeLabelColumn(
            SQLiteConnection conn,
            SQLiteTransaction tx,
            string table,
            string idColumn,
            string folderColumn,
            Dictionary<string, List<string>> leafLookup)
        {
            var rows = conn.Query<FolderRecord>(
                "SELECT " + idColumn + " AS Id, " + folderColumn + " AS Folder FROM " + table + " WHERE " + folderColumn + " IS NOT NULL;",
                transaction: tx).ToList();

            var changed = 0;
            foreach (var row in rows)
            {
                var normalized = FolderPathNormalizer.Normalize(row.Folder);
                if (leafLookup != null && normalized.IndexOf('/') < 0)
                {
                    List<string> candidates;
                    if (leafLookup.TryGetValue(normalized, out candidates) && candidates.Count == 1)
                        normalized = candidates[0];
                }

                if (!string.Equals(row.Folder ?? string.Empty, normalized, StringComparison.Ordinal))
                {
                    conn.Execute(
                        "UPDATE " + table + " SET " + folderColumn + " = @Folder WHERE " + idColumn + " = @Id;",
                        new { Folder = normalized, Id = row.Id },
                        tx);
                    changed++;
                }
            }

            return changed;
        }

        private static int DeleteSystemLabels(SQLiteConnection conn, SQLiteTransaction tx, string table, string idColumn, string folderColumn)
        {
            var rows = conn.Query<FolderRecord>(
                "SELECT " + idColumn + " AS Id, " + folderColumn + " AS Folder FROM " + table + " WHERE " + folderColumn + " IS NOT NULL;",
                transaction: tx).ToList();

            var deleted = 0;
            foreach (var row in rows)
            {
                if (!FolderPathNormalizer.IsExcludedSystemFolder(row.Folder)) continue;
                conn.Execute("DELETE FROM " + table + " WHERE " + idColumn + " = @Id;", new { Id = row.Id }, tx);
                deleted++;
            }

            return deleted;
        }

        private static int DeleteLabelsNotInSet(SQLiteConnection conn, SQLiteTransaction tx, string table, string idColumn, string folderColumn, HashSet<string> validPaths)
        {
            var rows = conn.Query<FolderRecord>(
                "SELECT " + idColumn + " AS Id, " + folderColumn + " AS Folder FROM " + table + " WHERE " + folderColumn + " IS NOT NULL;",
                transaction: tx).ToList();

            var deleted = 0;
            foreach (var row in rows)
            {
                var normalized = FolderPathNormalizer.Normalize(row.Folder);
                if (normalized.Length == 0 || validPaths.Contains(normalized)) continue;
                conn.Execute("DELETE FROM " + table + " WHERE " + idColumn + " = @Id;", new { Id = row.Id }, tx);
                deleted++;
            }

            return deleted;
        }

        private static void AddParameters(SQLiteCommand cmd, params string[] names)
        {
            foreach (var name in names)
            {
                var p = cmd.CreateParameter();
                p.ParameterName = name;
                cmd.Parameters.Add(p);
            }
        }

        private static void SetParameterValues(SQLiteCommand cmd, params object[] values)
        {
            for (var i = 0; i < values.Length; i++)
                cmd.Parameters[i].Value = values[i] ?? DBNull.Value;
        }

        private static EmailRow ToEmailRow(TrainingExample r)
        {
            return new EmailRow
            {
                Subject = r.Subject ?? string.Empty,
                Body = r.Body ?? string.Empty,
                FromAddress = r.FromAddress ?? string.Empty,
                SenderDomain = r.SenderDomain ?? string.Empty,
                HasAttachments = r.HasAttachments,
                Label = FolderPathNormalizer.Normalize(r.Label)
            };
        }

        private static DateTime? ParseUtc(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            DateTime parsed;
            if (!DateTime.TryParse(value, null, System.Globalization.DateTimeStyles.RoundtripKind, out parsed))
                return null;
            return parsed.Kind == DateTimeKind.Utc ? parsed : parsed.ToUniversalTime();
        }

        private sealed class TrainingRecord
        {
            public string Subject { get; set; }
            public string Body { get; set; }
            public string FromAddress { get; set; }
            public string SenderDomain { get; set; }
            public int HasAttachments { get; set; }
            public string Label { get; set; }
            public string ReceivedUtcText { get; set; }
            public string CreatedUtcText { get; set; }
            public string Source { get; set; }
        }

        private sealed class FolderRecord
        {
            public object Id { get; set; }
            public string Folder { get; set; }
        }
    }
}
