using OfficeOpenXml;
using BillMatch.Wpf.Models;
using System.Text.RegularExpressions;
using System.IO;

namespace BillMatch.Wpf.Services;

public class ExcelService : IExcelService
{
    public List<Transaction> LoadExcel(string filePath, ExcelMapping mapping)
    {
        var transactions = new List<Transaction>();

        try
        {
            var fileInfo = new FileInfo(filePath);
            if (!fileInfo.Exists)
            {
                throw new FileNotFoundException($"找不到文件: {filePath}");
            }

            if (string.Equals(fileInfo.Extension, ".xls", StringComparison.OrdinalIgnoreCase))
            {
                throw new NotSupportedException($"不支持 .xls 旧格式文件: {Path.GetFileName(filePath)}。请先在 Excel 中另存为 .xlsx 后再导入。");
            }

            using (var package = new ExcelPackage(fileInfo))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("Excel 文件中没有工作表。");
                }

                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet.Dimension == null)
                {
                    return transactions;
                }

                // 智能表头识别
                int headerRow = FindHeaderRow(worksheet);
                int startRow = headerRow + 1;

                // 将列名转换为索引
                int dateCol = ExcelColumnHelper.ColumnNameToIndex(mapping.DateColumn) + 1;
                int amountCol = ExcelColumnHelper.ColumnNameToIndex(mapping.AmountColumn) + 1;
                int cardCol = !string.IsNullOrEmpty(mapping.CardColumn) ? ExcelColumnHelper.ColumnNameToIndex(mapping.CardColumn) + 1 : 0;
                int descCol = !string.IsNullOrEmpty(mapping.DescriptionColumn) ? ExcelColumnHelper.ColumnNameToIndex(mapping.DescriptionColumn) + 1 : 0;
                int acc1Col = !string.IsNullOrEmpty(mapping.Account1Column) ? ExcelColumnHelper.ColumnNameToIndex(mapping.Account1Column) + 1 : 0;
                int acc2Col = !string.IsNullOrEmpty(mapping.Account2Column) ? ExcelColumnHelper.ColumnNameToIndex(mapping.Account2Column) + 1 : 0;

                for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    var transaction = new Transaction();

                    // 读取日期列
                    if (dateCol > 0 && dateCol <= worksheet.Dimension.End.Column)
                    {
                        var dateCell = worksheet.Cells[row, dateCol];
                        if (dateCell.Value != null)
                        {
                            transaction.Date = NormalizeDate(dateCell.Value);
                        }
                    }

                    // 读取金额列
                    if (amountCol > 0 && amountCol <= worksheet.Dimension.End.Column)
                    {
                        var amountCell = worksheet.Cells[row, amountCol];
                        if (amountCell.Value != null)
                        {
                            transaction.Amount = NormalizeAmount(amountCell.Value);
                        }
                    }

                    // 读取卡号列
                    if (cardCol > 0 && cardCol <= worksheet.Dimension.End.Column)
                    {
                        var cardCell = worksheet.Cells[row, cardCol];
                        if (cardCell.Value != null)
                        {
                            transaction.CardNumber = ExtractCardTail(cardCell.Value);
                        }
                    }

                    // 读取描述列
                    if (descCol > 0 && descCol <= worksheet.Dimension.End.Column)
                    {
                        var descCell = worksheet.Cells[row, descCol];
                        if (descCell.Value != null)
                        {
                            transaction.Description = descCell.Value?.ToString();
                        }
                    }

                    // 读取账户1列
                    if (acc1Col > 0 && acc1Col <= worksheet.Dimension.End.Column)
                    {
                        var acc1Cell = worksheet.Cells[row, acc1Col];
                        if (acc1Cell.Value != null)
                        {
                            transaction.Account1 = acc1Cell.Value?.ToString();
                        }
                    }

                    // 读取账户2列
                    if (acc2Col > 0 && acc2Col <= worksheet.Dimension.End.Column)
                    {
                        var acc2Cell = worksheet.Cells[row, acc2Col];
                        if (acc2Cell.Value != null)
                        {
                            transaction.Account2 = acc2Cell.Value?.ToString();
                        }
                    }

                    // 只添加有效的交易（有日期或金额）
                    if (transaction.Date != default || transaction.Amount != 0)
                    {
                        transactions.Add(transaction);
                    }
                }
            }
        }
        catch (NotSupportedException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new Exception($"解析 Excel 文件时出错 ({Path.GetFileName(filePath)}): {ex.Message}", ex);
        }

        return transactions;
    }

    /// <summary>
    /// 从银行账单文件读取日期范围
    /// </summary>
    public (DateTime? MinDate, DateTime? MaxDate) GetDateRange(string filePath, string dateColumn)
    {
        DateTime? minDate = null;
        DateTime? maxDate = null;

        try
        {
            var fileInfo = new FileInfo(filePath);
            if (!fileInfo.Exists)
            {
                return (null, null);
            }

            if (string.Equals(fileInfo.Extension, ".xls", StringComparison.OrdinalIgnoreCase))
            {
                return (null, null);
            }

            using (var package = new ExcelPackage(fileInfo))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    return (null, null);
                }

                var worksheet = package.Workbook.Worksheets[0];
                if (worksheet.Dimension == null)
                {
                    return (null, null);
                }

                // 将列名转换为索引
                int dateCol = ExcelColumnHelper.ColumnNameToIndex(dateColumn) + 1;
                if (dateCol <= 0 || dateCol > worksheet.Dimension.End.Column)
                {
                    return (null, null);
                }

                // 查找表头行
                int headerRow = FindHeaderRow(worksheet);
                int startRow = headerRow + 1;

                for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    var cell = worksheet.Cells[row, dateCol];
                    if (cell.Value != null)
                    {
                        var date = NormalizeDate(cell.Value);
                        if (date != default)
                        {
                            if (!minDate.HasValue || date < minDate.Value)
                            {
                                minDate = date;
                            }
                            if (!maxDate.HasValue || date > maxDate.Value)
                            {
                                maxDate = date;
                            }
                        }
                    }
                }
            }
        }
        catch
        {
            // 读取失败时返回null
        }

        return (minDate, maxDate);
    }

    /// <summary>
    /// 智能查找表头行（启发式搜索）
    /// </summary>
    internal int FindHeaderRow(ExcelWorksheet worksheet)
    {
        for (int row = 1; row <= Math.Min(5, worksheet.Dimension.End.Row); row++)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                var cellValue = worksheet.Cells[row, col].Value?.ToString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    if (cellValue.Contains("交易日期") || cellValue.Contains("日期") || 
                        cellValue.Contains("时间") || cellValue.Contains("金额"))
                    {
                        return row;
                    }
                }
            }
        }
        return 1;
    }

    /// <summary>
    /// 归一化日期
    /// </summary>
    internal DateTime NormalizeDate(object value)
    {
        if (value == null) return default;

        if (value is double oaDouble)
        {
            try { return DateTime.FromOADate(oaDouble).Date; }
            catch { return default; }
        }

        if (value is float oaFloat)
        {
            try { return DateTime.FromOADate(oaFloat).Date; }
            catch { return default; }
        }

        if (value is decimal oaDecimal)
        {
            try { return DateTime.FromOADate((double)oaDecimal).Date; }
            catch { return default; }
        }

        if (value is int oaInt)
        {
            try { return DateTime.FromOADate(oaInt).Date; }
            catch { return default; }
        }

        if (value is DateTime dt)
        {
            return dt.Date;
        }

        if (value is string str)
        {
            if (DateTime.TryParse(str, out var parsedDate))
            {
                return parsedDate.Date;
            }
        }
        else
        {
            if (DateTime.TryParse(value.ToString(), out var parsedDate))
            {
                return parsedDate.Date;
            }
        }

        return default;
    }

    /// <summary>
    /// 归一化金额
    /// </summary>
    internal decimal NormalizeAmount(object value)
    {
        if (value == null) return 0m;

        if (value is decimal dec) return dec;
        if (value is double dbl) return (decimal)dbl;
        if (value is int i) return i;
        if (value is long l) return l;

        string strVal = value.ToString() ?? "";
        strVal = strVal.Replace("￥", "").Replace(",", "").Replace("¥", "").Trim();
        
        if (decimal.TryParse(strVal, out var amount))
        {
            return amount;
        }

        return 0m;
    }

    /// <summary>
    /// 从字符串中提取最后4位数字
    /// </summary>
    internal string? ExtractCardTail(object value)
    {
        if (value == null) return null;

        string valueStr;

        if (value is double dbl && dbl == Math.Floor(dbl))
        {
            valueStr = ((long)dbl).ToString();
        }
        else if (value is decimal dec && dec == Math.Floor(dec))
        {
            valueStr = ((long)dec).ToString();
        }
        else if (value is int intVal)
        {
            valueStr = intVal.ToString();
        }
        else if (value is long longVal)
        {
            valueStr = longVal.ToString();
        }
        else
        {
            valueStr = value.ToString() ?? "";
        }

        string digits = Regex.Replace(valueStr, @"\D", "");

        if (digits.Length >= 4)
        {
            return digits.Substring(digits.Length - 4);
        }

        return null;
    }
}
