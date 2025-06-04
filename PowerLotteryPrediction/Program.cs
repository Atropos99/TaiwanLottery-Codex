using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using ClosedXML.Excel;

namespace PowerLotteryPrediction
{
    enum AnalysisMethod
    {
        Frequency = 1,
        RecencyWeighted = 2
    }

    class LotteryRecord
    {
        public int[] MainNumbers { get; set; } = Array.Empty<int>();
        public int? SpecialNumber { get; set; }
    }

    class Program
    {
        const int MinMain = 1;
        const int MaxMain = 38;
        const int MinSpecial = 1;
        const int MaxSpecial = 8;
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: dotnet run <excel-file>");
                return;
            }
            string excelPath = args[0];
            List<LotteryRecord> records = LoadRecords(excelPath, 100);
            Console.WriteLine("Select analysis method:");
            Console.WriteLine("1. Frequency based probability");
            Console.WriteLine("2. Recency weighted probability");
            Console.Write("Enter choice: ");
            if (!int.TryParse(Console.ReadLine(), out int choice) || !Enum.IsDefined(typeof(AnalysisMethod), choice))
            {
                Console.WriteLine("Invalid choice");
                return;
            }
            AnalysisMethod method = (AnalysisMethod)choice;
            var mainProbs = method switch
            {
                AnalysisMethod.RecencyWeighted => RecencyWeightedMain(records),
                _ => FrequencyMain(records)
            };
            var specialProbs = method switch
            {
                AnalysisMethod.RecencyWeighted => RecencyWeightedSpecial(records),
                _ => FrequencySpecial(records)
            };
            Console.WriteLine("Main number probabilities:");
            foreach (var kv in mainProbs.OrderBy(k => k.Key))
            {
                Console.WriteLine($"Number {kv.Key}: {kv.Value:P2}");
            }
            Console.WriteLine();
            Console.WriteLine("Special number probabilities:");
            foreach (var kv in specialProbs.OrderBy(k => k.Key))
            {
                Console.WriteLine($"Number {kv.Key}: {kv.Value:P2}");
            }

            // Predict numbers based on highest probabilities
            var predictedMain = mainProbs.OrderByDescending(kv => kv.Value)
                                         .Take(6)
                                         .Select(kv => kv.Key)
                                         .ToArray();
            int predictedSpecial = specialProbs.OrderByDescending(kv => kv.Value)
                                               .First().Key;

            Console.WriteLine();
            Console.WriteLine("Predicted main numbers: " + string.Join(", ", predictedMain));
            Console.WriteLine($"Predicted special number: {predictedSpecial}");
        }

        static List<LotteryRecord> LoadRecords(string path, int count)
        {
            var results = new List<LotteryRecord>();
            using var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheets.First();
            // assume headers at row 1, data from row 2
            int row = 2;
            while (ws.Row(row).CellsUsed().Any() && results.Count < count)
            {
                var numbers = new int[6];
                for (int i = 0; i < 6; i++)
                {
                    int val = ws.Cell(row, i + 1).GetValue<int>();
                    if (val < MinMain || val > MaxMain)
                        throw new InvalidDataException($"Main number {val} out of range at row {row}, column {i + 1}");
                    numbers[i] = val;
                }
                int? special = null;
                if (ws.Cell(row, 7).TryGetValue<int>(out int sp))
                {
                    if (sp < MinSpecial || sp > MaxSpecial)
                        throw new InvalidDataException($"Special number {sp} out of range at row {row}");
                    special = sp;
                }
                results.Add(new LotteryRecord { MainNumbers = numbers, SpecialNumber = special });
                row++;
            }
            return results;
        }

        static Dictionary<int, double> FrequencyMain(List<LotteryRecord> records)
        {
            var counts = Enumerable.Range(MinMain, MaxMain - MinMain + 1)
                                   .ToDictionary(n => n, _ => 0);
            foreach (var r in records)
            {
                foreach (var n in r.MainNumbers)
                {
                    counts[n] += 1;
                }
            }
            int totalNumbers = records.Count * records[0].MainNumbers.Length;
            return counts.ToDictionary(kv => kv.Key, kv => kv.Value / (double)totalNumbers);
        }

        static Dictionary<int, double> FrequencySpecial(List<LotteryRecord> records)
        {
            var counts = Enumerable.Range(MinSpecial, MaxSpecial - MinSpecial + 1)
                                   .ToDictionary(n => n, _ => 0);
            foreach (var r in records)
            {
                if (r.SpecialNumber.HasValue)
                    counts[r.SpecialNumber.Value] += 1;
            }
            int totalNumbers = records.Count(r => r.SpecialNumber.HasValue);
            if (totalNumbers == 0) return counts.ToDictionary(kv => kv.Key, _ => 0d);
            return counts.ToDictionary(kv => kv.Key, kv => kv.Value / (double)totalNumbers);
        }

        static Dictionary<int, double> RecencyWeightedMain(List<LotteryRecord> records)
        {
            var scores = Enumerable.Range(MinMain, MaxMain - MinMain + 1)
                                   .ToDictionary(n => n, _ => 0d);
            double weight = 1.0;
            double decay = 0.95; // more recent draws weigh more
            foreach (var r in Enumerable.Reverse(records))
            {
                foreach (var n in r.MainNumbers)
                {
                    scores[n] += weight;
                }
                weight *= decay;
            }
            double totalWeight = scores.Values.Sum();
            if (totalWeight == 0) return scores.ToDictionary(kv => kv.Key, _ => 0d);
            return scores.ToDictionary(kv => kv.Key, kv => kv.Value / totalWeight);
        }

        static Dictionary<int, double> RecencyWeightedSpecial(List<LotteryRecord> records)
        {
            var scores = Enumerable.Range(MinSpecial, MaxSpecial - MinSpecial + 1)
                                   .ToDictionary(n => n, _ => 0d);
            double weight = 1.0;
            double decay = 0.95;
            foreach (var r in Enumerable.Reverse(records))
            {
                if (r.SpecialNumber.HasValue)
                    scores[r.SpecialNumber.Value] += weight;
                weight *= decay;
            }
            double totalWeight = scores.Values.Sum();
            if (totalWeight == 0) return scores.ToDictionary(kv => kv.Key, _ => 0d);
            return scores.ToDictionary(kv => kv.Key, kv => kv.Value / totalWeight);
        }
    }
}

