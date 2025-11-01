using Duckov.Economy;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace MoreFormulasQX
{
    public static class FormulaExcelLoader
    {
        [Serializable]
        public struct CraftingFormulaInfo
        {
            public bool enabled;
            public string formulaID;
            public CraftingFormula.ItemEntry resultItem;
            public Cost cost;
            public string[] tags;
        }

        // 列名（首行为表头；第2、3行是辅助信息需跳过）
        private const string COL_ENABLED = "IsEnabled";
        private const string COL_RECIPE_ID = "ID";
        private const string COL_RESULT_ID = "ResultItemID";
        private const string COL_RESULT_AMOUNT = "ResultItemAmount";
        private const string COL_MONEY = "CostMoney";
        private const string COL_CONSUME_IDS = "CostItemID";
        private const string COL_CONSUME_AMOUNTS = "CostItemAmount";
        private const string COL_TAG = "Tag";
        // 注释列：ResultItemName、CostItemName（忽略，不参与读取）

        // path: Excel 路径；sheetName 或 sheetIndex 二选一（都不填则取第一个表）
        public static List<CraftingFormulaInfo> Load(string path, string sheetName = null, int sheetIndex = -1)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Excel 文件不存在：{path}");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            if (!MoveToSheet(reader, sheetName, sheetIndex))
                throw new Exception(TargetSheetNotFoundMessage(sheetName, sheetIndex));

            // 第1行：表头
            if (!reader.Read())
                throw new Exception("目标工作表没有任何行（缺少表头）。");

            var headerMap = BuildHeaderMap(reader);
            RequireColumns(headerMap, new[]
            {
            COL_ENABLED, COL_RECIPE_ID, COL_RESULT_ID, COL_RESULT_AMOUNT,
            COL_MONEY, COL_CONSUME_IDS, COL_CONSUME_AMOUNTS, COL_TAG
        });

            // 第2、3行：辅助信息，跳过
            int dataRowNumber = 1;            // 当前 reader 所在行号（1=表头）
            if (reader.Read()) dataRowNumber = 2;
            if (reader.Read()) dataRowNumber = 3;

            var list = new List<CraftingFormulaInfo>(16);

            // 从第4行开始读取数据
            while (reader.Read())
            {
                dataRowNumber++;

                if (RowIsCompletelyEmpty(reader)) continue;

                // 是否启用
                bool enabled = ParseBool(GetCell(reader, headerMap[COL_ENABLED]));

                // 必填字段
                string formulaID = GetCell(reader, headerMap[COL_RECIPE_ID]).Trim();
                if (string.IsNullOrEmpty(formulaID))
                    throw new Exception($"第{dataRowNumber}行：{COL_RECIPE_ID} 为空。");

                int resultId = ParseInt32Required(GetCell(reader, headerMap[COL_RESULT_ID]), $"{COL_RESULT_ID}（第{dataRowNumber}行）");
                int resultAmount = ParseInt32Required(GetCell(reader, headerMap[COL_RESULT_AMOUNT]), $"{COL_RESULT_AMOUNT}（第{dataRowNumber}行）", mustBePositive: true);

                long money = ParseInt64Required(GetCell(reader, headerMap[COL_MONEY]), $"{COL_MONEY}（第{dataRowNumber}行）", allowZero: true);

                // 消耗项（逗号分隔）
                var consumeIds = ParseIntList(GetCell(reader, headerMap[COL_CONSUME_IDS]));
                var consumeAmts = ParseLongList(GetCell(reader, headerMap[COL_CONSUME_AMOUNTS]), mustBePositive: true);

                if (consumeIds.Count != consumeAmts.Count)
                    throw new Exception($"第{dataRowNumber}行：{COL_CONSUME_IDS} 与 {COL_CONSUME_AMOUNTS} 数量不一致，IDs={consumeIds.Count}, Amounts={consumeAmts.Count}。");

                var consumeItems = new Cost.ItemEntry[consumeIds.Count];
                for (int i = 0; i < consumeIds.Count; i++)
                {
                    consumeItems[i] = new Cost.ItemEntry
                    {
                        id = consumeIds[i],
                        amount = consumeAmts[i]
                    };
                }

                // Tag（逗号分隔，空则给空数组）
                var tags = ParseStringList(GetCell(reader, headerMap[COL_TAG])).ToArray();

                list.Add(new CraftingFormulaInfo
                {
                    enabled = enabled,
                    formulaID = formulaID,
                    resultItem = new CraftingFormula.ItemEntry { id = resultId, amount = resultAmount },
                    cost = new Cost { money = money, items = consumeItems },
                    tags = tags
                });
            }

            return list;
        }

        // ========= 辅助 =========

        private static bool MoveToSheet(IExcelDataReader reader, string sheetName, int sheetIndex)
        {
            if (sheetName == null && sheetIndex < 0) return true;

            int idx = 0;
            do
            {
                bool nameMatch = sheetName != null && string.Equals(reader.Name, sheetName, StringComparison.OrdinalIgnoreCase);
                bool indexMatch = sheetIndex >= 0 && idx == sheetIndex;
                if ((sheetName != null && nameMatch) || (sheetIndex >= 0 && indexMatch))
                    return true;
                idx++;
            } while (reader.NextResult());

            return false;
        }

        private static string TargetSheetNotFoundMessage(string sheetName, int sheetIndex)
        {
            return sheetName != null
                ? $"找不到指定工作表：{sheetName}"
                : $"找不到索引为 {sheetIndex} 的工作表";
        }

        private static Dictionary<string, int> BuildHeaderMap(IExcelDataReader reader)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < reader.FieldCount; c++)
            {
                var name = reader.GetValue(c)?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(name) && !map.ContainsKey(name))
                    map[name] = c;
            }
            return map;
        }

        private static void RequireColumns(Dictionary<string, int> headerMap, IEnumerable<string> columns)
        {
            var missing = new List<string>();
            foreach (var col in columns)
                if (!headerMap.ContainsKey(col)) missing.Add(col);

            if (missing.Count > 0)
                throw new Exception("缺少必须的列：" + string.Join("，", missing));
        }

        private static bool RowIsCompletelyEmpty(IExcelDataReader reader)
        {
            for (int c = 0; c < reader.FieldCount; c++)
            {
                var v = reader.GetValue(c);
                if (v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                    return false;
            }
            return true;
        }

        private static string GetCell(IExcelDataReader reader, int colIndex)
        {
            if (colIndex < 0 || colIndex >= reader.FieldCount) return string.Empty;
            var v = reader.GetValue(colIndex);
            return v?.ToString() ?? string.Empty;
        }

        private static bool ParseBool(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            switch (s.ToLowerInvariant())
            {
                case "true":
                case "yes":
                case "y":
                case "1":
                case "是":
                case "开启":
                case "启用":
                    return true;
                case "false":
                case "no":
                case "n":
                case "0":
                case "否":
                case "关闭":
                case "禁用":
                    return false;
            }
            if (bool.TryParse(s, out var b)) return b;
            throw new Exception($"布尔值无法识别：'{s}'（支持 TRUE/FALSE、是/否、1/0）");
        }

        private static int ParseInt32Required(string s, string fieldName, bool mustBePositive = false, bool allowZero = false)
        {
            if (!TryParseInt32(s, out var v))
                throw new Exception($"{fieldName} 必须为整数，但值为：'{s}'");

            if (mustBePositive && v <= 0)
                throw new Exception($"{fieldName} 必须为正整数，但值为：{v}");

            if (!allowZero && !mustBePositive && v == 0)
                throw new Exception($"{fieldName} 不允许为 0。");

            return v;
        }

        private static long ParseInt64Required(string s, string fieldName, bool mustBePositive = false, bool allowZero = false)
        {
            if (!TryParseInt64(s, out var v))
                throw new Exception($"{fieldName} 必须为整数，但值为：'{s}'");

            if (mustBePositive && v <= 0)
                throw new Exception($"{fieldName} 必须为正整数，但值为：{v}");

            if (!allowZero && !mustBePositive && v == 0)
                throw new Exception($"{fieldName} 不允许为 0。");

            return v;
        }

        private static bool TryParseInt32(string s, out int value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return true;

            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
            {
                value = (int)d;
                return true;
            }
            return false;
        }

        private static bool TryParseInt64(string s, out long value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();

            if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return true;

            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
            {
                value = (long)d;
                return true;
            }
            return false;
        }

        private static List<int> ParseIntList(string raw)
        {
            var parts = SplitList(raw);
            var result = new List<int>(parts.Count);
            for (int i = 0; i < parts.Count; i++)
            {
                if (!TryParseInt32(parts[i], out var v))
                    throw new Exception($"消耗物品ID 列第 {i + 1} 项不是有效整数：'{parts[i]}'");
                result.Add(v);
            }
            return result;
        }

        private static List<long> ParseLongList(string raw, bool mustBePositive = false)
        {
            var parts = SplitList(raw);
            var result = new List<long>(parts.Count);
            for (int i = 0; i < parts.Count; i++)
            {
                if (!TryParseInt64(parts[i], out var v))
                    throw new Exception($"消耗物品数量 列第 {i + 1} 项不是有效整数：'{parts[i]}'");

                if (mustBePositive && v <= 0)
                    throw new Exception($"消耗物品数量 必须为正整数（第 {i + 1} 项），当前：{v}");

                result.Add(v);
            }
            return result;
        }

        private static List<string> ParseStringList(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new List<string>();
            var normalized = raw.Replace('，', ',');
            return normalized
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        private static List<string> SplitList(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new List<string>();
            var normalized = raw.Replace('，', ',');
            return normalized
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }
    }
}