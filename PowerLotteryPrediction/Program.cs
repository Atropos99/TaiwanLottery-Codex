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
        RecencyWeighted = 2,
        Last30Frequency = 3,
        Last10Frequency = 4,
        Hybrid = 5
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
                Console.WriteLine("使用方法：dotnet run <excel檔案>");
                return;
            }
            string excelPath = args[0];
            List<LotteryRecord> records = LoadRecords(excelPath, 100);
            Console.WriteLine("請選擇分析方法:");
            Console.WriteLine("1. 依出現頻率計算機率");
            Console.WriteLine("2. 依近期加權計算機率");
            Console.WriteLine("3. 最近30期頻率");
            Console.WriteLine("4. 最近10期頻率");
            Console.WriteLine("5. 綜合(頻率+加權)計算機率");
            Console.Write("輸入選項: ");
            if (!int.TryParse(Console.ReadLine(), out int choice) || !Enum.IsDefined(typeof(AnalysisMethod), choice))
            {
                Console.WriteLine("無效的選項");
                return;
            }
            AnalysisMethod method = (AnalysisMethod)choice;
            var mainProbs = method switch
            {
                AnalysisMethod.Frequency => FrequencyMain(records),
                AnalysisMethod.RecencyWeighted => RecencyWeightedMain(records),
                AnalysisMethod.Last30Frequency => RecentFrequencyMain(records, 30),
                AnalysisMethod.Last10Frequency => RecentFrequencyMain(records, 10),
                AnalysisMethod.Hybrid => HybridMain(records),
                _ => FrequencyMain(records)
            };
            var specialProbs = method switch
            {
                AnalysisMethod.Frequency => FrequencySpecial(records),
                AnalysisMethod.RecencyWeighted => RecencyWeightedSpecial(records),
                AnalysisMethod.Last30Frequency => RecentFrequencySpecial(records, 30),
                AnalysisMethod.Last10Frequency => RecentFrequencySpecial(records, 10),
                AnalysisMethod.Hybrid => HybridSpecial(records),
                _ => FrequencySpecial(records)
            };
            Console.WriteLine("主號機率:");
            foreach (var kv in mainProbs.OrderBy(k => k.Key))
            {
                Console.WriteLine($"號碼 {kv.Key}: {kv.Value:P2}");
            }
            Console.WriteLine();
            Console.WriteLine("特別號機率:");
            foreach (var kv in specialProbs.OrderBy(k => k.Key))
            {
                Console.WriteLine($"號碼 {kv.Key}: {kv.Value:P2}");
            }
            // Predict numbers based on highest probabilities
            var predictedMain = mainProbs.OrderByDescending(kv => kv.Value)
                                         .Take(6)
                                         .Select(kv => kv.Key)
                                         .ToArray();
            int predictedSpecial = specialProbs.OrderByDescending(kv => kv.Value)
                                               .First().Key;

            // Numbers with the lowest probabilities
            var leastMain = mainProbs.OrderBy(kv => kv.Value)
                                    .Take(6)
                                    .Select(kv => kv.Key)
                                    .ToArray();
            int leastSpecial = specialProbs.OrderBy(kv => kv.Value)
                                           .First().Key;

            Console.WriteLine();
            Console.WriteLine("預測主號: " + string.Join(", ", predictedMain));
            Console.WriteLine($"預測特別號: {predictedSpecial}");
            Console.WriteLine("機率最低主號: " + string.Join(", ", leastMain));
            Console.WriteLine($"機率最低特別號: {leastSpecial}");
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
                        throw new InvalidDataException($"主號 {val} 在第 {row} 列第 {i + 1} 欄超出範圍");
                    numbers[i] = val;
                }
                int? special = null;
                if (ws.Cell(row, 7).TryGetValue<int>(out int sp))
                {
                    if (sp < MinSpecial || sp > MaxSpecial)
                        throw new InvalidDataException($"特別號 {sp} 在第 {row} 列超出範圍");
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

        static Dictionary<int, double> RecentFrequencyMain(List<LotteryRecord> records, int recent)
        {
            var slice = records.TakeLast(recent).ToList();
            return FrequencyMain(slice);
        }

        static Dictionary<int, double> RecentFrequencySpecial(List<LotteryRecord> records, int recent)
        {
            var slice = records.TakeLast(recent).ToList();
            return FrequencySpecial(slice);
        }

        static Dictionary<int, double> HybridMain(List<LotteryRecord> records)
        {
            var freq = FrequencyMain(records);
            var recency = RecencyWeightedMain(records);
            return freq.ToDictionary(kv => kv.Key, kv => (kv.Value + recency[kv.Key]) / 2);
        }

        static Dictionary<int, double> HybridSpecial(List<LotteryRecord> records)
        {
            var freq = FrequencySpecial(records);
            var recency = RecencyWeightedSpecial(records);
            return freq.ToDictionary(kv => kv.Key, kv => (kv.Value + recency[kv.Key]) / 2);
        }
    }
}

