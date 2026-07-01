using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.ML.Transforms.Text;
using Microsoft.ML.Trainers;
using Newtonsoft.Json;
using OutlookClassifierAddIn5.Services;

namespace OutlookClassifierAddIn5.ML
{
    public sealed class EmailRow
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string FromAddress { get; set; }
        public string SenderDomain { get; set; }
        public bool HasAttachments { get; set; }
        public string Label { get; set; }
    }

    public sealed class TrainingExample
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string FromAddress { get; set; }
        public string SenderDomain { get; set; }
        public bool HasAttachments { get; set; }
        public string Label { get; set; }
        public DateTime? ReceivedUtc { get; set; }
        public string Source { get; set; }
    }

    public sealed class Prediction
    {
        [ColumnName("PredictedLabel")]
        public string Folder { get; set; }

        public float[] Score { get; set; }
    }

    public sealed class ModelMetadata
    {
        public int ModelVersion { get; set; }
        public string TrainedUtc { get; set; }
        public int TrainingRows { get; set; }
        public int LabelCount { get; set; }
        public int BodyCap { get; set; }
        public int RecentDays { get; set; }
        public string AppVersion { get; set; }
        public int FeaturePipelineVersion { get; set; }
        public double? Top1Accuracy { get; set; }
        public double? Top3Accuracy { get; set; }
        public double? AccuracyAt70 { get; set; }
        public double? AccuracyAt92 { get; set; }
    }

    public sealed class EvaluationMetrics
    {
        public bool Ran { get; set; }
        public string SkipReason { get; set; }
        public int TrainingRows { get; set; }
        public int ValidationRows { get; set; }
        public int LabelCount { get; set; }
        public double Top1Accuracy { get; set; }
        public double Top3Accuracy { get; set; }
        public double? AccuracyAt70 { get; set; }
        public double? AccuracyAt92 { get; set; }
        public int CoverageAt70 { get; set; }
        public int CoverageAt92 { get; set; }
        public List<string> WorstConfusions { get; set; }
        public List<string> FoldersWithTooFewExamples { get; set; }

        public EvaluationMetrics()
        {
            WorstConfusions = new List<string>();
            FoldersWithTooFewExamples = new List<string>();
        }
    }

    public sealed class TrainingResult
    {
        public int TrainingRows { get; set; }
        public int LabelCount { get; set; }
        public ModelMetadata Metadata { get; set; }
        public EvaluationMetrics Evaluation { get; set; }

        public string ToStatusSummary()
        {
            if (Evaluation != null && Evaluation.Ran)
            {
                return string.Format(
                    "Training complete: {0} rows, {1} folders, top-3 accuracy {2:P1}",
                    TrainingRows,
                    LabelCount,
                    Evaluation.Top3Accuracy);
            }

            return string.Format("Training complete: {0} rows, {1} folders", TrainingRows, LabelCount);
        }
    }

    public class ModelService
    {
        public const int ModelVersion = 1;
        public const int FeaturePipelineVersion = 1;
        public const int BodyCap = 1000;
        public const int DefaultRecentDays = 180;

        private readonly MLContext _ml = new MLContext(0);
        private ITransformer _model;
        private DataViewSchema _schema;
        private PredictionEngine<EmailRow, Prediction> _engine;
        private string[] _classNames;
        private readonly object _predictLock = new object();

        public ModelMetadata LastMetadata { get; private set; }
        public TrainingResult LastTrainingResult { get; private set; }

        public bool IsLoaded
        {
            get { return _model != null; }
        }

        public string ModelSignature
        {
            get
            {
                var meta = LastMetadata;
                if (meta == null) return IsLoaded ? "loaded-without-metadata" : "none";
                return meta.ModelVersion + ":" + meta.FeaturePipelineVersion + ":" + meta.TrainedUtc + ":" + meta.TrainingRows + ":" + meta.LabelCount;
            }
        }

        public void Train(IEnumerable<EmailRow> rows, out DataViewSchema schema)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            var examples = rows.Select(r => new TrainingExample
            {
                Subject = r.Subject,
                Body = r.Body,
                FromAddress = r.FromAddress,
                SenderDomain = r.SenderDomain,
                HasAttachments = r.HasAttachments,
                Label = r.Label
            }).ToList();

            TrainWithEvaluation(examples, DefaultRecentDays, out schema);
        }

        public TrainingResult TrainWithEvaluation(IEnumerable<TrainingExample> examples, int recentDays, out DataViewSchema schema)
        {
            if (examples == null) throw new ArgumentNullException(nameof(examples));

            var totalSw = System.Diagnostics.Stopwatch.StartNew();
            var prepareSw = System.Diagnostics.Stopwatch.StartNew();
            var prepared = PrepareExamples(examples);
            prepareSw.Stop();
            var rows = prepared.Select(p => p.Row).ToList();
            var labelCount = rows.Select(r => r.Label).Distinct(StringComparer.OrdinalIgnoreCase).Count();

            if (labelCount < 2)
                throw new InvalidOperationException("Model training requires at least two different folders after cleanup.");

            AppLogger.Info("Model train start. Rows=" + rows.Count + ", labels=" + labelCount + ".");

            var evalSw = System.Diagnostics.Stopwatch.StartNew();
            var evaluation = EvaluateIfPossible(prepared);
            evalSw.Stop();

            var data = _ml.Data.LoadFromEnumerable(rows);
            var pipeline = BuildPipeline();

            try
            {
                var fitSw = System.Diagnostics.Stopwatch.StartNew();
                _model = pipeline.Fit(data);
                fitSw.Stop();
                _schema = data.Schema;
                schema = _schema;
                BuildEngineAndClassNames();

                LastMetadata = new ModelMetadata
                {
                    ModelVersion = ModelVersion,
                    TrainedUtc = DateTime.UtcNow.ToString("o"),
                    TrainingRows = rows.Count,
                    LabelCount = labelCount,
                    BodyCap = BodyCap,
                    RecentDays = recentDays,
                    AppVersion = GetAppVersion(),
                    FeaturePipelineVersion = FeaturePipelineVersion,
                    Top1Accuracy = evaluation != null && evaluation.Ran ? (double?)evaluation.Top1Accuracy : null,
                    Top3Accuracy = evaluation != null && evaluation.Ran ? (double?)evaluation.Top3Accuracy : null,
                    AccuracyAt70 = evaluation == null ? null : evaluation.AccuracyAt70,
                    AccuracyAt92 = evaluation == null ? null : evaluation.AccuracyAt92
                };

                LastTrainingResult = new TrainingResult
                {
                    TrainingRows = rows.Count,
                    LabelCount = labelCount,
                    Metadata = LastMetadata,
                    Evaluation = evaluation
                };

                LogEvaluation(evaluation);
                totalSw.Stop();
                AppLogger.Info("Model train end. " + LastTrainingResult.ToStatusSummary() +
                    ", prepareMs=" + prepareSw.ElapsedMilliseconds +
                    ", evalMs=" + evalSw.ElapsedMilliseconds +
                    ", fitMs=" + fitSw.ElapsedMilliseconds +
                    ", totalMs=" + totalSw.ElapsedMilliseconds + ".");

                return LastTrainingResult;
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "ML.NET training failed.");
                throw;
            }
        }

        public (string folder, List<(string name, float p)> top3, float conf) PredictTop3(EmailRow row)
        {
            if (_model == null) throw new InvalidOperationException("Model not trained");
            if (_engine == null || _classNames == null) BuildEngineAndClassNames();

            Prediction pred;
            lock (_predictLock)
            {
                pred = _engine.Predict(SanitizePredictionRow(row));
            }

            var ranked = RankScores(pred.Score).Take(3).ToList();
            return (pred.Folder, ranked, ranked.Count > 0 ? ranked[0].p : 0f);
        }

        public IEnumerable<(string folder, List<(string name, float p)> top3, float conf)>
            PredictBatch(IEnumerable<EmailRow> rows)
        {
            if (_model == null) throw new InvalidOperationException("Model not trained");
            if (rows == null) yield break;
            if (_engine == null || _classNames == null) BuildEngineAndClassNames();

            var sanitized = rows.Select(SanitizePredictionRow).ToList();
            var dv = _ml.Data.LoadFromEnumerable(sanitized);
            var scored = _model.Transform(dv);

            var pred = scored.GetColumn<string>("PredictedLabel").ToArray();
            var scores = scored.GetColumn<VBuffer<float>>("Score").ToArray();

            for (int i = 0; i < scores.Length; i++)
            {
                var dense = scores[i].DenseValues().ToArray();
                var ranked = RankScores(dense).Take(3).ToList();
                yield return (pred[i], ranked, ranked.Count > 0 ? ranked[0].p : 0f);
            }
        }

        public void Save(string path)
        {
            if (_model == null) throw new InvalidOperationException("Model not trained");
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            _ml.Model.Save(_model, _schema, path);

            var metadata = LastMetadata ?? new ModelMetadata
            {
                ModelVersion = ModelVersion,
                TrainedUtc = DateTime.UtcNow.ToString("o"),
                TrainingRows = 0,
                LabelCount = 0,
                BodyCap = BodyCap,
                RecentDays = DefaultRecentDays,
                AppVersion = GetAppVersion(),
                FeaturePipelineVersion = FeaturePipelineVersion
            };

            File.WriteAllText(GetMetadataPath(path), JsonConvert.SerializeObject(metadata, Formatting.Indented));
            AppLogger.Info("Model saved: " + path);
        }

        public void Load(string path)
        {
            if (!TryLoad(path))
                throw new InvalidOperationException("Model could not be loaded.");
        }

        public bool TryLoad(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    return false;

                var metadataPath = GetMetadataPath(path);
                ModelMetadata metadata = null;
                if (File.Exists(metadataPath))
                {
                    metadata = JsonConvert.DeserializeObject<ModelMetadata>(File.ReadAllText(metadataPath));
                    if (!IsMetadataCompatible(metadata))
                    {
                        AppLogger.Warn("Model metadata is incompatible. Model will not be loaded.");
                        return false;
                    }
                }
                else
                {
                    AppLogger.Warn("Model metadata is missing. Loading model for compatibility, but automation remains disabled.");
                }

                _model = _ml.Model.Load(path, out _schema);
                LastMetadata = metadata;
                BuildEngineAndClassNames();
                AppLogger.Info("Model load success: " + path);
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Model load failed.");
                return false;
            }
        }

        public bool CanAutoApprove(double threshold, out string reason)
        {
            reason = string.Empty;

            if (!IsLoaded)
            {
                reason = "No model is loaded.";
                return false;
            }

            if (LastMetadata == null)
            {
                reason = "The loaded model has no metadata or evaluation.";
                return false;
            }

            if (!LastMetadata.Top3Accuracy.HasValue || !LastMetadata.AccuracyAt92.HasValue)
            {
                reason = "Auto-approve requires post-training evaluation metrics.";
                return false;
            }

            if (LastMetadata.Top3Accuracy.Value < 0.80)
            {
                reason = "Top-3 validation accuracy is below 80%.";
                return false;
            }

            if (threshold >= 0.92 && LastMetadata.AccuracyAt92.Value < 0.90)
            {
                reason = "High-confidence validation accuracy is below 90%.";
                return false;
            }

            return true;
        }

        private IEstimator<ITransformer> BuildPipeline()
        {
            var subjectPipe =
                _ml.Transforms.Text.NormalizeText(
                        outputColumnName: "s_norm",
                        inputColumnName: nameof(EmailRow.Subject),
                        caseMode: TextNormalizingEstimator.CaseMode.Lower,
                        keepDiacritics: false,
                        keepPunctuations: false,
                        keepNumbers: true)
                  .Append(_ml.Transforms.Text.TokenizeIntoWords(
                        outputColumnName: "s_tok",
                        inputColumnName: "s_norm"))
                  .Append(_ml.Transforms.Text.ProduceHashedWordBags(
                        outputColumnName: "fSubject",
                        inputColumnName: "s_tok",
                        ngramLength: 2,
                        useAllLengths: true,
                        numberOfBits: 16))
                  .Append(_ml.Transforms.Text.TokenizeIntoCharactersAsKeys(
                        outputColumnName: "s_chars",
                        inputColumnName: "s_norm"))
                  .Append(_ml.Transforms.Text.ProduceNgrams(
                        outputColumnName: "fSubjectChar",
                        inputColumnName: "s_chars",
                        ngramLength: 3,
                        useAllLengths: false,
                        maximumNgramsCount: 16000,
                        weighting: NgramExtractingEstimator.WeightingCriteria.Tf));

            var fromPipe =
                _ml.Transforms.Text.NormalizeText(
                        outputColumnName: "fa_norm",
                        inputColumnName: nameof(EmailRow.FromAddress),
                        caseMode: TextNormalizingEstimator.CaseMode.Lower,
                        keepDiacritics: false,
                        keepPunctuations: false,
                        keepNumbers: true)
                  .Append(_ml.Transforms.Text.TokenizeIntoWords(
                        outputColumnName: "fa_tok",
                        inputColumnName: "fa_norm"))
                  .Append(_ml.Transforms.Text.ProduceHashedWordBags(
                        outputColumnName: "fFrom",
                        inputColumnName: "fa_tok",
                        ngramLength: 2,
                        useAllLengths: true,
                        numberOfBits: 15))
                  .Append(_ml.Transforms.Text.TokenizeIntoCharactersAsKeys(
                        outputColumnName: "fa_chars",
                        inputColumnName: "fa_norm"))
                  .Append(_ml.Transforms.Text.ProduceNgrams(
                        outputColumnName: "fFromChar",
                        inputColumnName: "fa_chars",
                        ngramLength: 3,
                        useAllLengths: false,
                        maximumNgramsCount: 16000,
                        weighting: NgramExtractingEstimator.WeightingCriteria.Tf));

            var bodyPipe =
                _ml.Transforms.Text.NormalizeText(
                        outputColumnName: "b_norm",
                        inputColumnName: nameof(EmailRow.Body),
                        caseMode: TextNormalizingEstimator.CaseMode.Lower,
                        keepDiacritics: false,
                        keepPunctuations: false,
                        keepNumbers: false)
                  .Append(_ml.Transforms.Text.TokenizeIntoWords(
                        outputColumnName: "b_tok",
                        inputColumnName: "b_norm"))
                  .Append(_ml.Transforms.Text.ProduceHashedWordBags(
                        outputColumnName: "fBody",
                        inputColumnName: "b_tok",
                        ngramLength: 1,
                        useAllLengths: true,
                        numberOfBits: 16));

            var sdca = _ml.MulticlassClassification.Trainers.SdcaMaximumEntropy(
                new SdcaMaximumEntropyMulticlassTrainer.Options
                {
                    MaximumNumberOfIterations = 100
                });

            return subjectPipe
                .Append(bodyPipe)
                .Append(fromPipe)
                .Append(_ml.Transforms.Categorical.OneHotHashEncoding("fDomain", nameof(EmailRow.SenderDomain)))
                .Append(_ml.Transforms.Conversion.ConvertType(
                    outputColumnName: "HasAttachmentsF",
                    inputColumnName: nameof(EmailRow.HasAttachments),
                    outputKind: DataKind.Single))
                .Append(_ml.Transforms.Concatenate("Features",
                    new[] { "fSubject", "fSubjectChar", "fFrom", "fFromChar", "fBody", "fDomain", "HasAttachmentsF" }))
                .Append(_ml.Transforms.Conversion.MapValueToKey("Label"))
                .AppendCacheCheckpoint(_ml)
                .Append(sdca)
                .Append(_ml.Transforms.Conversion.MapKeyToValue("PredictedLabel"));
        }

        private List<PreparedTrainingExample> PrepareExamples(IEnumerable<TrainingExample> examples)
        {
            var list = examples.Select(r =>
            {
                var from = r.FromAddress ?? string.Empty;
                var domain = string.IsNullOrWhiteSpace(r.SenderDomain)
                    ? SenderResolutionService.ExtractDomain(from)
                    : r.SenderDomain;

                var body = r.Body ?? string.Empty;
                if (body.Length > BodyCap) body = body.Substring(0, BodyCap);

                return new PreparedTrainingExample
                {
                    Row = new EmailRow
                    {
                        Subject = r.Subject ?? string.Empty,
                        Body = body,
                        FromAddress = from,
                        SenderDomain = domain,
                        HasAttachments = r.HasAttachments,
                        Label = FolderPathNormalizer.Normalize(r.Label)
                    },
                    ReceivedUtc = r.ReceivedUtc,
                    Source = r.Source ?? string.Empty
                };
            })
            .Where(r => !string.IsNullOrWhiteSpace(r.Row.Label))
            .Where(r => !FolderPathNormalizer.IsExcludedSystemFolder(r.Row.Label))
            .ToList();

            var counts = list.GroupBy(r => r.Row.Label, StringComparer.OrdinalIgnoreCase)
                             .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);

            list = list.Where(r => counts[r.Row.Label] >= 2).ToList();
            return ClassBalancedSample(list, 2000);
        }

        private EvaluationMetrics EvaluateIfPossible(List<PreparedTrainingExample> prepared)
        {
            var metrics = new EvaluationMetrics();
            var labelCounts = prepared.GroupBy(r => r.Row.Label, StringComparer.OrdinalIgnoreCase)
                                      .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);

            metrics.FoldersWithTooFewExamples = labelCounts
                .Where(kv => kv.Value < 5)
                .OrderBy(kv => kv.Value)
                .Take(10)
                .Select(kv => kv.Key + " (" + kv.Value + ")")
                .ToList();

            if (prepared.Count < 50)
            {
                metrics.SkipReason = "Not enough rows for holdout evaluation.";
                return metrics;
            }

            var labels = prepared.Select(r => r.Row.Label).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            if (labels < 2)
            {
                metrics.SkipReason = "Not enough labels for holdout evaluation.";
                return metrics;
            }

            var withDates = prepared.Count(r => r.ReceivedUtc.HasValue);
            var ordered = withDates >= prepared.Count / 2
                ? prepared.OrderBy(r => r.ReceivedUtc.HasValue ? r.ReceivedUtc.Value : DateTime.MinValue).ToList()
                : prepared.ToList();

            var validationCount = Math.Max(10, (int)Math.Round(ordered.Count * 0.20));
            if (ordered.Count - validationCount < 20)
            {
                metrics.SkipReason = "Not enough training rows after holdout split.";
                return metrics;
            }

            var train = ordered.Take(ordered.Count - validationCount).ToList();
            var validation = ordered.Skip(ordered.Count - validationCount).ToList();
            var trainLabels = new HashSet<string>(train.Select(r => r.Row.Label), StringComparer.OrdinalIgnoreCase);
            validation = validation.Where(r => trainLabels.Contains(r.Row.Label)).ToList();

            if (validation.Count < 10 || trainLabels.Count < 2)
            {
                metrics.SkipReason = "Holdout labels were not represented in training split.";
                return metrics;
            }

            try
            {
                var trainData = _ml.Data.LoadFromEnumerable(train.Select(r => r.Row));
                var validationRows = validation.Select(r => r.Row).ToList();
                var validationData = _ml.Data.LoadFromEnumerable(validationRows);
                var model = BuildPipeline().Fit(trainData);
                var scored = model.Transform(validationData);

                var predicted = scored.GetColumn<string>("PredictedLabel").ToArray();
                var scores = scored.GetColumn<VBuffer<float>>("Score").ToArray();
                var classNames = GetClassNames(scored.Schema);

                int top1 = 0;
                int top3 = 0;
                int at70 = 0;
                int ok70 = 0;
                int at92 = 0;
                int ok92 = 0;
                var confusions = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                for (var i = 0; i < validationRows.Count; i++)
                {
                    var truth = validationRows[i].Label;
                    var pred = predicted[i];
                    var dense = scores[i].DenseValues().ToArray();
                    var ranked = dense
                        .Select((score, idx) => new { Name = idx < classNames.Length ? classNames[idx] : string.Empty, Score = score })
                        .OrderByDescending(x => x.Score)
                        .Take(3)
                        .ToList();

                    var conf = ranked.Count == 0 ? 0f : ranked[0].Score;

                    if (string.Equals(truth, pred, StringComparison.OrdinalIgnoreCase))
                    {
                        top1++;
                    }
                    else
                    {
                        var key = truth + " -> " + pred;
                        int count;
                        confusions[key] = confusions.TryGetValue(key, out count) ? count + 1 : 1;
                    }

                    if (ranked.Any(x => string.Equals(x.Name, truth, StringComparison.OrdinalIgnoreCase)))
                        top3++;

                    if (conf >= 0.70f)
                    {
                        at70++;
                        if (string.Equals(truth, pred, StringComparison.OrdinalIgnoreCase)) ok70++;
                    }

                    if (conf >= 0.92f)
                    {
                        at92++;
                        if (string.Equals(truth, pred, StringComparison.OrdinalIgnoreCase)) ok92++;
                    }
                }

                metrics.Ran = true;
                metrics.TrainingRows = train.Count;
                metrics.ValidationRows = validationRows.Count;
                metrics.LabelCount = trainLabels.Count;
                metrics.Top1Accuracy = top1 / (double)validationRows.Count;
                metrics.Top3Accuracy = top3 / (double)validationRows.Count;
                metrics.CoverageAt70 = at70;
                metrics.CoverageAt92 = at92;
                metrics.AccuracyAt70 = at70 > 0 ? (double?)ok70 / at70 : null;
                metrics.AccuracyAt92 = at92 > 0 ? (double?)ok92 / at92 : null;
                metrics.WorstConfusions = confusions
                    .OrderByDescending(kv => kv.Value)
                    .Take(5)
                    .Select(kv => kv.Key + " (" + kv.Value + ")")
                    .ToList();
            }
            catch (Exception ex)
            {
                metrics.Ran = false;
                metrics.SkipReason = "Evaluation failed: " + ex.Message;
                AppLogger.Error(ex, "Post-training evaluation failed.");
            }

            return metrics;
        }

        private List<(string name, float p)> RankScores(float[] scores)
        {
            if (scores == null) return new List<(string name, float p)>();
            return scores
                .Select((p, i) => (name: _classNames != null && i < _classNames.Length ? _classNames[i] : string.Empty, p: p))
                .OrderByDescending(t => t.p)
                .ToList();
        }

        private static EmailRow SanitizePredictionRow(EmailRow row)
        {
            row = row ?? new EmailRow();
            var body = row.Body ?? string.Empty;
            if (body.Length > BodyCap) body = body.Substring(0, BodyCap);
            var from = row.FromAddress ?? string.Empty;

            return new EmailRow
            {
                Subject = row.Subject ?? string.Empty,
                Body = body,
                FromAddress = from,
                SenderDomain = string.IsNullOrWhiteSpace(row.SenderDomain) ? SenderResolutionService.ExtractDomain(from) : row.SenderDomain,
                HasAttachments = row.HasAttachments,
                Label = row.Label ?? string.Empty
            };
        }

        private static List<PreparedTrainingExample> ClassBalancedSample(IEnumerable<PreparedTrainingExample> src, int perLabelCap)
        {
            var caps = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var keep = new List<PreparedTrainingExample>();
            foreach (var r in src)
            {
                int used;
                if (!caps.TryGetValue(r.Row.Label, out used)) used = 0;
                if (used >= perLabelCap) continue;
                caps[r.Row.Label] = used + 1;
                keep.Add(r);
            }
            return keep;
        }

        private void BuildEngineAndClassNames()
        {
            if (_model == null) return;

            if (_engine != null)
            {
                try { _engine.Dispose(); } catch (Exception ex) { AppLogger.Warn("Prediction engine dispose failed: " + ex.Message); }
                _engine = null;
            }

            _engine = _ml.Model.CreatePredictionEngine<EmailRow, Prediction>(_model);
            _classNames = GetClassNames(_engine.OutputSchema);
        }

        private static string[] GetClassNames(DataViewSchema schema)
        {
            var slot = default(VBuffer<ReadOnlyMemory<char>>);
            schema[nameof(Prediction.Score)].GetSlotNames(ref slot);
            return slot.DenseValues().Select(s => s.ToString()).ToArray();
        }

        private static bool IsMetadataCompatible(ModelMetadata metadata)
        {
            if (metadata == null) return false;
            return metadata.ModelVersion == ModelVersion
                && metadata.FeaturePipelineVersion == FeaturePipelineVersion
                && metadata.BodyCap == BodyCap;
        }

        private static string GetMetadataPath(string modelPath)
        {
            return Path.ChangeExtension(modelPath, ".meta.json");
        }

        private static string GetAppVersion()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            return version == null ? "1.0.0" : version.ToString();
        }

        private static void LogEvaluation(EvaluationMetrics evaluation)
        {
            if (evaluation == null)
            {
                AppLogger.Warn("No evaluation metrics produced.");
                return;
            }

            if (!evaluation.Ran)
            {
                AppLogger.Warn("Training evaluation skipped: " + evaluation.SkipReason);
                return;
            }

            AppLogger.Info(string.Format(
                "Evaluation: train={0}, validate={1}, labels={2}, top1={3:P2}, top3={4:P2}, acc>=0.70={5}, acc>=0.92={6}.",
                evaluation.TrainingRows,
                evaluation.ValidationRows,
                evaluation.LabelCount,
                evaluation.Top1Accuracy,
                evaluation.Top3Accuracy,
                evaluation.AccuracyAt70.HasValue ? evaluation.AccuracyAt70.Value.ToString("P2") : "n/a",
                evaluation.AccuracyAt92.HasValue ? evaluation.AccuracyAt92.Value.ToString("P2") : "n/a"));

            if (evaluation.WorstConfusions.Count > 0)
                AppLogger.Info("Worst confused folders: " + string.Join("; ", evaluation.WorstConfusions.ToArray()));

            if (evaluation.FoldersWithTooFewExamples.Count > 0)
                AppLogger.Info("Folders with few examples: " + string.Join("; ", evaluation.FoldersWithTooFewExamples.ToArray()));
        }

        private sealed class PreparedTrainingExample
        {
            public EmailRow Row { get; set; }
            public DateTime? ReceivedUtc { get; set; }
            public string Source { get; set; }
        }
    }
}
