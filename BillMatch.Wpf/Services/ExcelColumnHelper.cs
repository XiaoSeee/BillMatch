using System;
using System.Collections.Generic;
using System.Linq;

namespace BillMatch.Wpf.Services
{
    /// <summary>
    /// Excel列名辅助类 - 处理A,B,C...与索引的转换
    /// </summary>
    public static class ExcelColumnHelper
    {
        /// <summary>
        /// 将Excel列名(如A, B, AA, AB)转换为0-based索引
        /// </summary>
        public static int ColumnNameToIndex(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
                throw new ArgumentException("列名不能为空", nameof(columnName));

            columnName = columnName.Trim().ToUpperInvariant();
            int result = 0;

            foreach (char c in columnName)
            {
                if (c < 'A' || c > 'Z')
                    throw new ArgumentException($"无效的列名字符: {c}", nameof(columnName));

                result = result * 26 + (c - 'A' + 1);
            }

            return result - 1; // 转换为0-based索引
        }

        /// <summary>
        /// 将0-based索引转换为Excel列名(如A, B, AA, AB)
        /// </summary>
        public static string IndexToColumnName(int index)
        {
            if (index < 0)
                throw new ArgumentException("索引不能为负数", nameof(index));

            string result = "";
            int temp = index + 1;

            while (temp > 0)
            {
                temp--;
                result = (char)('A' + (temp % 26)) + result;
                temp /= 26;
            }

            return result;
        }

        /// <summary>
        /// 获取所有可用的Excel列名列表(从A到指定列)
        /// </summary>
        public static List<string> GetColumnNames(int count)
        {
            return Enumerable.Range(0, count)
                .Select(IndexToColumnName)
                .ToList();
        }

        /// <summary>
        /// 验证列名是否有效
        /// </summary>
        public static bool IsValidColumnName(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
                return false;

            columnName = columnName.Trim().ToUpperInvariant();
            return columnName.All(c => c >= 'A' && c <= 'Z');
        }

        /// <summary>
        /// 解析用户输入的列名，支持多种格式(A, 1, 第A列等)
        /// </summary>
        public static string? NormalizeColumnName(string? input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            input = input.Trim().ToUpperInvariant();

            // 移除常见前缀
            input = input.Replace("第", "").Replace("列", "").Trim();

            // 如果输入是纯数字，转换为列名
            if (int.TryParse(input, out int number) && number > 0)
            {
                return IndexToColumnName(number - 1);
            }

            // 验证是否为有效的列名
            if (IsValidColumnName(input))
            {
                return input;
            }

            return null;
        }
    }
}
