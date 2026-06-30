// Testing/SyntheticData.cs
using System;
using System.Collections.Generic;
using System.Linq;
using OutlookClassifierAddIn5.ML; // <-- so we can use EmailRow

namespace OutlookClassifierAddIn5.Testing
{
    public static class SyntheticData
    {
        // Builds N folders, including duplicate leafs under 2024/2025 to test disambiguation.
        public static List<string> BuildFolderSet(int numBaseTeams = 40, int yearsBack = 3, string account = "Account")
        {
            var result = new List<string>();

            // Top-level user folders
            for (int i = 0; i < numBaseTeams; i++)
            {
                var name = $"Team{i:00}";
                result.Add($"{account}/Projects/{name}");
            }

            // Year/PRM Team dup leaves: 2023/PRM Team, 2024/PRM Team, ...
            int currentYear = DateTime.Now.Year;
            for (int y = currentYear - yearsBack + 1; y <= currentYear; y++)
            {
                result.Add($"{account}/Inbox/{y}/PRM Team");
                result.Add($"{account}/Inbox/{y}/Q{(y % 4) + 1}/PRM Team"); // deeper dup
            }

            return result;
        }

        // Long-tail label sampler (Zipf-ish): a few folders popular, many rare.
        private static int SampleFolderIndex(int k, Random rng)
        {
            var weights = new double[k];
            double sum = 0;
            for (int i = 0; i < k; i++) { sum += (weights[i] = 1.0 / Math.Pow(i + 1, 1.1)); }
            double r = rng.NextDouble() * sum, c = 0;
            for (int i = 0; i < k; i++) { c += weights[i]; if (c >= r) return i; }
            return k - 1;
        }

        private static string RandomFrom(Random rng, params string[] arr) => arr[rng.Next(arr.Length)];

        public static List<EmailRow> GenerateRows(int total, List<string> fullFolderPaths, int bodyCap = 1000)
        {
            var rng = new Random(42);
            var rows = new List<EmailRow>(total);

            string[] actions = { "Update", "Invoice", "Action required", "Reminder", "Minutes", "Proposal", "SOW", "NDA" };
            string[] projects = { "PRM", "ERP", "CRM", "MIGRATION", "QBR", "Roadmap", "ADC", "Ops", "SEV" };
            string[] vendors = { "Contoso", "Fabrikam", "Northwind", "AdventureWorks", "Wingtip", "Tailspin" };
            string[] domains = { "contoso.com", "fabrikam.com", "northwind.example", "partner.example", "alerts.example" };

            for (int i = 0; i < total; i++)
            {
                int labelIdx = SampleFolderIndex(fullFolderPaths.Count, rng);
                string label = fullFolderPaths[labelIdx];

                string vendor = RandomFrom(rng, vendors);
                string proj = RandomFrom(rng, projects);
                string act = RandomFrom(rng, actions);
                int po = 10000 + rng.Next(90000);
                string quarter = "Q" + (1 + rng.Next(4));
                string fy = "FY" + (DateTime.Now.Year % 100);

                var subject = $"{vendor} {proj} {act} {quarter} {fy} PO {po}";
                var from = $"{proj.ToLower()}_{rng.Next(200):D3}@{RandomFrom(rng, domains)}";

                var body = $"Hi team,\n{act} for {proj}. PO #{po}. Please review by {quarter} {fy}.\nThanks,\n{vendor}";
                if (body.Length > bodyCap) body = body.Substring(0, bodyCap);

                rows.Add(new EmailRow
                {
                    Subject = subject,
                    Body = body,
                    FromAddress = from,
                    SenderDomain = from.Substring(from.IndexOf('@') + 1),
                    HasAttachments = rng.NextDouble() < 0.3,
                    Label = label
                });
            }

            return rows;
        }
    }
}
