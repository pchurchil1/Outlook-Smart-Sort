// Testing/SyntheticRunner.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using OutlookClassifierAddIn5.ML;
using OutlookClassifierAddIn5.Testing;

namespace OutlookClassifierAddIn5.Testing
{
    public static class SyntheticRunner
    {
        public static void Run()
        {
            var folders = SyntheticData.BuildFolderSet(120, 4, "Primary");
            const int TOTAL = 120_000;
            var rows = SyntheticData.GenerateRows(TOTAL, folders, bodyCap: 1000);

            int trainN = (int)(rows.Count * 0.8);
            var trainRows = rows.Take(trainN).ToList();
            var testRows  = rows.Skip(trainN).ToList();

            var ms = new ModelService();
            Microsoft.ML.DataViewSchema schema;

            var sw = Stopwatch.StartNew();
            ms.Train(trainRows, out schema);
            sw.Stop();
            Debug.WriteLine($"[BENCH] Train {trainRows.Count} rows in {sw.ElapsedMilliseconds} ms");

            sw.Restart();
            var preds = ms.PredictBatch(testRows).ToList();
            sw.Stop();
            Debug.WriteLine($"[BENCH] Predict {testRows.Count} rows in {sw.ElapsedMilliseconds} ms");

            int top1 = 0, top3 = 0; double conf = 0;
            for (int i = 0; i < testRows.Count; i++)
            {
                var truth = testRows[i].Label;
                var p = preds[i];
                if (string.Equals(p.folder, truth, StringComparison.OrdinalIgnoreCase)) top1++;
                if (p.top3.Any(t => string.Equals(t.name, truth, StringComparison.OrdinalIgnoreCase))) top3++;
                conf += p.conf;
            }
            Debug.WriteLine($"[BENCH] Top1={top1/(double)testRows.Count:P2}, Top3={top3/(double)testRows.Count:P2}, MeanConf={conf/(double)testRows.Count:0.000}");
        }
    }
}
