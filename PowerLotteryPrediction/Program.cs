using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using ClosedXML.Excel;

// 這支程式用來讀取台灣大樂透開獎資料，
// 並以不同的統計方法計算各號碼再次開出的機率。

namespace PowerLotteryPrediction
{
    /// <summary>
    /// 提供不同的分析方式選項。
    /// </summary>
    enum AnalysisMethod
    {
        /// <summary>計算所有歷史資料中每個號碼的出現頻率</summary>
        Frequency = 1,
        /// <summary>根據近期期數給予較高權重計算機率</summary>
        RecencyWeighted = 2,
        /// <summary>僅統計最近30期的出現頻率</summary>
        Last30Frequency = 3,
        /// <summary>僅統計最近10期的出現頻率</summary>
        Last10Frequency = 4,
        /// <summary>綜合頻率與近期加權兩種方法</summary>
        Hybrid = 5,
        /// <summary>利用簡易 AR(1) 時間序列預測</summary>
        TimeSeries = 6
    }

    /// <summary>
    /// 封裝單一期開獎資料的資料結構。
    /// </summary>
    class LotteryRecord
    {
        /// <summary>六個主號碼</summary>
        public int[] MainNumbers { get; set; } = Array.Empty<int>();

        /// <summary>特別號，可能沒有開出</summary>
        public int? SpecialNumber { get; set; }
    }

    /// <summary>
    /// 程式主體，包含讀取資料與各種分析方法。
    /// </summary>
    class Program
    {
        // 主號的最小與最大值
        const int MinMain = 1;
        const int MaxMain = 38;
        // 特別號的最小與最大值
        const int MinSpecial = 1;
        const int MaxSpecial = 8;
        /// <summary>
        /// 程式進入點，讀取資料並依使用者指定的方式計算號碼機率。
        /// </summary>
        static void Main(string[] args)
        {
            // 未提供參數時顯示使用方式
            if (args.Length == 0)
            {
                Console.WriteLine("使用方法：dotnet run <excel檔案>");
                return;
            }

            // 讀取第一個參數作為 Excel 檔路徑
            string excelPath = args[0];
            // 取得最近 100 期的開獎資料
            List<LotteryRecord> records = LoadRecords(excelPath, 100);

            // 讓使用者選擇欲採用的分析方法
            // 列出所有可使用的分析方式讓使用者選擇
            Console.WriteLine("請選擇分析方法:");
            Console.WriteLine("1. 依出現頻率計算機率");
            Console.WriteLine("2. 依近期加權計算機率");
            Console.WriteLine("3. 最近30期頻率");
            Console.WriteLine("4. 最近10期頻率");
            Console.WriteLine("5. 綜合(頻率+加權)計算機率");
            Console.WriteLine("6. 時間序列分析(ARIMA)");
            Console.Write("輸入選項: ");
            // 讀取並驗證選項
            if (!int.TryParse(Console.ReadLine(), out int choice) || !Enum.IsDefined(typeof(AnalysisMethod), choice))
            {
                Console.WriteLine("無效的選項");
                return;
            }
            // 將輸入的數字轉成列舉型別
            AnalysisMethod method = (AnalysisMethod)choice;

            // 依選擇的方法計算主號及特別號的機率分布
            var mainProbs = method switch
            {
                AnalysisMethod.Frequency => FrequencyMain(records),
                AnalysisMethod.RecencyWeighted => RecencyWeightedMain(records),
                AnalysisMethod.Last30Frequency => RecentFrequencyMain(records, 30),
                AnalysisMethod.Last10Frequency => RecentFrequencyMain(records, 10),
                AnalysisMethod.Hybrid => HybridMain(records),
                AnalysisMethod.TimeSeries => TimeSeriesMain(records),
                _ => FrequencyMain(records)
            };
            var specialProbs = method switch
            {
                AnalysisMethod.Frequency => FrequencySpecial(records),
                AnalysisMethod.RecencyWeighted => RecencyWeightedSpecial(records),
                AnalysisMethod.Last30Frequency => RecentFrequencySpecial(records, 30),
                AnalysisMethod.Last10Frequency => RecentFrequencySpecial(records, 10),
                AnalysisMethod.Hybrid => HybridSpecial(records),
                AnalysisMethod.TimeSeries => TimeSeriesSpecial(records),
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
            // 根據計算出的機率，選出機率最高的 6 個主號
            var predictedMain = mainProbs.OrderByDescending(kv => kv.Value)
                                         .Take(6)
                                         .Select(kv => kv.Key)
                                         .ToArray();
            // 機率最高的特別號
            int predictedSpecial = specialProbs.OrderByDescending(kv => kv.Value)
                                               .First().Key;

            // 反過來找出機率最低的號碼，供參考
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

        /// <summary>
        /// 從指定的 Excel 檔案讀取開獎紀錄。
        /// </summary>
        /// <param name="path">Excel 檔路徑</param>
        /// <param name="count">最多讀取的期數</param>
        static List<LotteryRecord> LoadRecords(string path, int count)
        {
            // 儲存結果的集合
            var results = new List<LotteryRecord>();

            // 以 ClosedXML 開啟 Excel
            using var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheets.First();

            // 假設第一列是標題，資料從第二列開始
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
                // 第七欄若有資料則視為特別號
                if (ws.Cell(row, 7).TryGetValue<int>(out int sp))
                {
                    if (sp < MinSpecial || sp > MaxSpecial)
                        throw new InvalidDataException($"特別號 {sp} 在第 {row} 列超出範圍");
                    special = sp;
                }
                // 將讀到的紀錄加入清單
                results.Add(new LotteryRecord { MainNumbers = numbers, SpecialNumber = special });
                row++;
            }
            return results;
        }

        /// <summary>
        /// 以所有歷史紀錄計算各主號出現的頻率。
        /// </summary>
        static Dictionary<int, double> FrequencyMain(List<LotteryRecord> records)
        {
            // 初始化號碼計數器
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
            // 轉換成機率分布
            return counts.ToDictionary(kv => kv.Key, kv => kv.Value / (double)totalNumbers);
        }

        /// <summary>
        /// 以所有歷史紀錄計算特別號出現頻率。
        /// </summary>
        static Dictionary<int, double> FrequencySpecial(List<LotteryRecord> records)
        {
            // 初始化特別號計數器
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

        /// <summary>
        /// 以遞減權重計算主號機率，越近期的期數權重越高。
        /// </summary>
        static Dictionary<int, double> RecencyWeightedMain(List<LotteryRecord> records)
        {
            // 各號碼的累積權重
            var scores = Enumerable.Range(MinMain, MaxMain - MinMain + 1)
                                   .ToDictionary(n => n, _ => 0d);
            double weight = 1.0;
            double decay = 0.95; // 每往前一期，權重乘以 decay
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

        /// <summary>
        /// 以遞減權重計算特別號機率。
        /// </summary>
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

        /// <summary>
        /// 只統計最近指定期數的主號頻率。
        /// </summary>
        static Dictionary<int, double> RecentFrequencyMain(List<LotteryRecord> records, int recent)
        {
            var slice = records.TakeLast(recent).ToList();
            return FrequencyMain(slice);
        }

        /// <summary>
        /// 只統計最近指定期數的特別號頻率。
        /// </summary>
        static Dictionary<int, double> RecentFrequencySpecial(List<LotteryRecord> records, int recent)
        {
            var slice = records.TakeLast(recent).ToList();
            return FrequencySpecial(slice);
        }

        /// <summary>
        /// 將頻率與近期加權兩種方法平均後的主號機率。
        /// </summary>
        static Dictionary<int, double> HybridMain(List<LotteryRecord> records)
        {
            var freq = FrequencyMain(records);
            var recency = RecencyWeightedMain(records);
            return freq.ToDictionary(kv => kv.Key, kv => (kv.Value + recency[kv.Key]) / 2);
        }

        /// <summary>
        /// 將頻率與近期加權兩種方法平均後的特別號機率。
        /// </summary>
        static Dictionary<int, double> HybridSpecial(List<LotteryRecord> records)
        {
            var freq = FrequencySpecial(records);
            var recency = RecencyWeightedSpecial(records);
            return freq.ToDictionary(kv => kv.Key, kv => (kv.Value + recency[kv.Key]) / 2);
        }

        /// <summary>
        /// 利用簡易 AR(1) 模型預測主號再次出現的機率。
        /// </summary>
        static Dictionary<int, double> TimeSeriesMain(List<LotteryRecord> records)
        {
            var result = Enumerable.Range(MinMain, MaxMain - MinMain + 1)
                                    .ToDictionary(n => n, _ => 0d);
            foreach (int n in result.Keys.ToList())
            {
                // 將歷史紀錄轉換為 0/1 序列作為時間序列資料
                var series = records.Select(r => r.MainNumbers.Contains(n) ? 1.0 : 0.0).ToList();
                result[n] = Ar1Predict(series);
            }
            return result;
        }

        /// <summary>
        /// 利用簡易 AR(1) 模型預測特別號再次出現的機率。
        /// </summary>
        static Dictionary<int, double> TimeSeriesSpecial(List<LotteryRecord> records)
        {
            var result = Enumerable.Range(MinSpecial, MaxSpecial - MinSpecial + 1)
                                    .ToDictionary(n => n, _ => 0d);
            foreach (int n in result.Keys.ToList())
            {
                var series = records.Select(r => r.SpecialNumber.HasValue && r.SpecialNumber.Value == n ? 1.0 : 0.0).ToList();
                result[n] = Ar1Predict(series);
            }
            return result;
        }

        /// <summary>
        /// 簡單的 AR(1) 預測函式，輸入 0/1 序列後回傳下一期為 1 的機率。
        /// </summary>
        static double Ar1Predict(IList<double> series)
        {
            if (series.Count == 0) return 0;
            if (series.Count == 1) return series[0];

            double mean = series.Average();
            double num = 0, den = 0;
            for (int i = 1; i < series.Count; i++)
            {
                num += (series[i - 1] - mean) * (series[i] - mean);
                den += (series[i - 1] - mean) * (series[i - 1] - mean);
            }
            double phi = den == 0 ? 0 : num / den;
            double c = mean - phi * mean;
            double pred = c + phi * series[^1];
            if (pred < 0) pred = 0;
            if (pred > 1) pred = 1;
            return pred;
        }
    }
}

