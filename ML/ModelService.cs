using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.ML.Transforms.Text;
using Microsoft.ML.Trainers;

namespace OutlookClassifierAddIn5.ML
{
    // Training/prediction row
    public sealed class EmailRow
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string FromAddress { get; set; }
        public string SenderDomain { get; set; }
        public bool HasAttachments { get; set; }

        // For training only; can be null at predict time
        public string Label { get; set; }
    }

    // Predicted output schema
    public sealed class Prediction
    {
        [ColumnName("PredictedLabel")]
        public string Folder { get; set; }

        public float[] Score { get; set; }
    }

    public class ModelService
    {
        private readonly MLContext _ml = new MLContext(0);
        private ITransformer _model;
        private DataViewSchema _schema;

        // Cached single-row engine (for interactive UI) + class names cache
        private PredictionEngine<EmailRow, Prediction> _engine;
        private string[] _classNames;
        private readonly object _predictLock = new object();

        /// <summary>
        /// Train (or retrain) the multiclass model from in-memory rows.
        /// Applies data hygiene, caps body length, class-balanced sampling,
        /// and uses a balanced, fast feature set.
        /// </summary>
        public void Train(IEnumerable<EmailRow> rows, out DataViewSchema schema)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            // 1) Sanitize + defensive copies
            var list = rows.Select(r => new EmailRow
            {
                Subject = r.Subject ?? string.Empty,
                Body = r.Body ?? string.Empty,
                FromAddress = r.FromAddress ?? string.Empty,
                SenderDomain = r.SenderDomain ?? string.Empty,
                HasAttachments = r.HasAttachments,
                Label = string.IsNullOrWhiteSpace(r.Label) ? null : r.Label
            })
            .Where(r => !string.IsNullOrWhiteSpace(r.Label))
            .ToList();

            // 2) Guard: never learn to "file to Inbox" (root inbox only)
            list = list.Where(r => !IsInboxPath(r.Label)).ToList();

            // 3) Cap body length to reduce featurization cost
            const int BodyCap = 1000;
            for (int i = 0; i < list.Count; i++)
            {
                var b = list[i].Body;
                if (!string.IsNullOrEmpty(b) && b.Length > BodyCap)
                    list[i].Body = b.Substring(0, BodyCap);
            }

            // 4) Drop labels with < 2 examples to avoid degenerate classes
            var labelCounts = list.GroupBy(r => r.Label)
                                  .ToDictionary(g => g.Key, g => g.Count());
            list = list.Where(r => labelCounts[r.Label] >= 2).ToList();

            // 5) Class-balanced sampling: keep at most N examples per label (newest-first assumed)
            list = ClassBalancedSample(list, perLabelCap: 2000);

            var distinctLabels = list.GroupBy(r => r.Label).Count();
            if (distinctLabels < 2)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Model training requires at least two different folders (labels) after cleanup.\n" +
                    "Try broadening the training window or confirming folder names.",
                    "Not enough training data");
                throw new InvalidOperationException("Insufficient label diversity for training.");
            }

            // 6) IDataView
            var data = _ml.Data.LoadFromEnumerable(list);

            // 7) Build balanced, fixed-size text features (version-safe explicit pipeline)

            // SUBJECT: words (uni+bi, keep numbers) + tiny char 3-grams (dictionary-based)
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
                        ngramLength: 2,           // uni+bi
                        useAllLengths: true,
                        numberOfBits: 16))        // 65,536 dims (hashed)
                  .Append(_ml.Transforms.Text.TokenizeIntoCharactersAsKeys(
                        outputColumnName: "s_chars",
                        inputColumnName: "s_norm"))
                  .Append(_ml.Transforms.Text.ProduceNgrams(           // <-- portable alt to ProduceHashedNGrams
                        outputColumnName: "fSubjectChar",
                        inputColumnName: "s_chars",
                        ngramLength: 3,
                        useAllLengths: false,
                        maximumNgramsCount: 16000,                       // cap dictionary size
                        weighting: NgramExtractingEstimator.WeightingCriteria.Tf));

            // FROM: words (uni+bi, keep numbers) + tiny char 3-grams (dictionary-based)
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
                        numberOfBits: 15))        // 32,768 dims (hashed)
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

            // BODY: words (unigrams only), trimmed, no char grams
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

            // 8) Trainer with bounded iterations for predictable runtime
            var sdca = _ml.MulticlassClassification.Trainers.SdcaMaximumEntropy(
                new SdcaMaximumEntropyMulticlassTrainer.Options
                {
                    MaximumNumberOfIterations = 100, // balanced speed/accuracy
                    // L2Regularization = 1e-4f,     // optional: small L2 can help generalization
                });

            var pipeline =
                subjectPipe
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

            // 9) Fit
            try
            {
                _model = pipeline.Fit(data);
                _schema = data.Schema;
                schema = _schema;

                // Reset caches for prediction
                BuildEngineAndClassNames();
            }
            catch (Exception ex)
            {
                var root = ex;
                while (root.InnerException != null) root = root.InnerException;

                System.Windows.Forms.MessageBox.Show(
                    "ML.NET training failed:\n\n" + root.Message + "\n\n" + root.StackTrace,
                    "Training error");
                throw;
            }
        }

        /// <summary>
        /// Single-item prediction for interactive UI.
        /// Returns (folder, top3 (name, prob), top1 probability).
        /// </summary>
        public (string folder, List<(string name, float p)> top3, float conf) PredictTop3(EmailRow row)
        {
            if (_model == null) throw new InvalidOperationException("Model not trained");
            if (_engine == null || _classNames == null) BuildEngineAndClassNames();

            Prediction pred;
            lock (_predictLock)
            {
                pred = _engine.Predict(row);
            }

            var ranked = pred.Score
                .Select((p, i) => Tuple.Create(_classNames[i], p))
                .OrderByDescending(t => t.Item2)
                .Take(3)
                .Select(t => (t.Item1, t.Item2))
                .ToList();

            return (pred.Folder, ranked, ranked.Count > 0 ? ranked[0].Item2 : 0f);
        }

        /// <summary>
        /// High-throughput batch prediction for queue building.
        /// </summary>
        public IEnumerable<(string folder, List<(string name, float p)> top3, float conf)>
            PredictBatch(IEnumerable<EmailRow> rows)
        {
            if (_model == null) throw new InvalidOperationException("Model not trained");
            if (rows == null) yield break;

            if (_engine == null || _classNames == null) BuildEngineAndClassNames();

            var dv = _ml.Data.LoadFromEnumerable(rows);
            var scored = _model.Transform(dv);

            var pred = scored.GetColumn<string>("PredictedLabel").ToArray();
            var scores = scored.GetColumn<VBuffer<float>>("Score").ToArray();

            for (int i = 0; i < scores.Length; i++)
            {
                var dense = scores[i].DenseValues().ToArray();
                var ranked = dense
                    .Select((p, idx) => new { Name = _classNames[idx], P = p })
                    .OrderByDescending(x => x.P)
                    .Take(3)
                    .Select(x => (x.Name, x.P))
                    .ToList();

                yield return (pred[i], ranked, ranked.Count > 0 ? ranked[0].P : 0f);
            }
        }

        public void Save(string path)
        {
            if (_model == null) throw new InvalidOperationException("Model not trained");
            _ml.Model.Save(_model, _schema, path);
        }

        public void Load(string path)
        {
            _model = _ml.Model.Load(path, out _schema);
            BuildEngineAndClassNames();
        }

        // ----------------------- helpers -----------------------

        private static bool IsInboxPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            var s = path.Replace('\\', '/').Trim().Trim('/').ToLowerInvariant();
            return s == "inbox" || s.EndsWith("/inbox");
        }

        /// <summary>
        /// Keep at most perLabelCap rows per label, in input order.
        /// Assumes input is roughly newest-first (as provided by the DB).
        /// </summary>
        private static List<EmailRow> ClassBalancedSample(IEnumerable<EmailRow> src, int perLabelCap)
        {
            var caps = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var keep = new List<EmailRow>();
            foreach (var r in src)
            {
                if (string.IsNullOrWhiteSpace(r.Label)) continue;
                int used;
                if (!caps.TryGetValue(r.Label, out used)) used = 0;
                if (used >= perLabelCap) continue;
                caps[r.Label] = used + 1;
                keep.Add(r);
            }
            return keep;
        }

        private void BuildEngineAndClassNames()
        {
            if (_model == null) return;

            if (_engine != null)
            {
                try { _engine.Dispose(); } catch { }
                _engine = null;
            }

            _engine = _ml.Model.CreatePredictionEngine<EmailRow, Prediction>(_model);

            // Cache class names (Score slot names)
            var slot = default(VBuffer<ReadOnlyMemory<char>>);
            _engine.OutputSchema[nameof(Prediction.Score)].GetSlotNames(ref slot);
            _classNames = slot.DenseValues().Select(s => s.ToString()).ToArray();
        }
    }
}