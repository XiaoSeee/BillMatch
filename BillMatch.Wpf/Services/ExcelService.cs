using System.IO;
using System.Text.RegularExpressions;
using BillMatch.Wpf.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;

namespace BillMatch.Wpf.Services;

public class ExcelService : IExcelService
{
    public List<Transaction> LoadExcel(string filePath, ExcelMapping mapping)
    {
        var fileInfo = new FileInfo(filePath);
        if (!fileInfo.Exists)
        {
            throw new FileNotFoundException($"找不到文件: {filePath}");
        }

        try
        {
            return IsXls(fileInfo.Extension)
                ? LoadXls(fileInfo, mapping)
                : LoadOpenXml(fileInfo, mapping);
        }
        catch (Exception ex)
        {
            throw new Exception($"解析 Excel 文件时出错({Path.GetFileName(filePath)}): {ex.Message}", ex);
        }
    }

    public (DateTime? MinDate, DateTime? MaxDate) GetDateRange(string filePath, string dateColumn)
    {
        var fileInfo = new FileInfo(filePath);
        if (!fileInfo.Exists)
        {
            return (null, null);
        }

        try
        {
            return IsXls(fileInfo.Extension)
                ? GetDateRangeFromXls(fileInfo, dateColumn)
                : GetDateRangeFromOpenXml(fileInfo, dateColumn);
        }
        catch
        {
            return (null, null);
        }
    }

    internal int FindHeaderRow(ExcelWorksheet worksheet)
    {
        for (int row = 1; row <= Math.Min(5, worksheet.Dimension.End.Row); row++)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                var cellValue = worksheet.Cells[row, col].Value?.ToString();
                if (IsHeaderKeyword(cellValue))
                {
                    return row;
                }
            }
        }

        return 1;
    }

    internal int FindHeaderRow(ISheet sheet)
    {
        var maxScanRow = Math.Min(4, sheet.LastRowNum);

        for (int rowIndex = 0; rowIndex <= maxScanRow; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                continue;
            }

            for (int colIndex = row.FirstCellNum; colIndex < row.LastCellNum; colIndex++)
            {
                if (colIndex < 0)
                {
                    continue;
                }

                var cellValue = GetCellValue(row.GetCell(colIndex))?.ToString();
                if (IsHeaderKeyword(cellValue))
                {
                    return rowIndex;
                }
            }
        }

        return 0;
    }

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

    internal decimal NormalizeAmount(object value)
    {
        if (value == null) return 0m;

        if (value is decimal dec) return dec;
        if (value is double dbl) return (decimal)dbl;
        if (value is float flt) return (decimal)flt;
        if (value is int i) return i;
        if (value is long l) return l;

        var strVal = value.ToString() ?? string.Empty;
        strVal = strVal.Replace("￥", "")
            .Replace("¥", "")
            .Replace(",", "")
            .Replace("元", "")
            .Trim();

        return decimal.TryParse(strVal, out var amount) ? amount : 0m;
    }

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
            valueStr = value.ToString() ?? string.Empty;
        }

        var digits = Regex.Replace(valueStr, @"\D", "");

        return digits.Length >= 4 ? digits.Substring(digits.Length - 4) : null;
    }

    private List<Transaction> LoadOpenXml(FileInfo fileInfo, ExcelMapping mapping)
    {
        var transactions = new List<Transaction>();

        using var package = new ExcelPackage(fileInfo);
        if (package.Workbook.Worksheets.Count == 0)
        {
            throw new Exception("Excel 文件中没有工作表。");
        }

        var worksheet = package.Workbook.Worksheets[0];
        if (worksheet.Dimension == null)
        {
            return transactions;
        }

        var headerRow = FindHeaderRow(worksheet);
        var startRow = headerRow + 1;

        var dateCol = ExcelColumnHelper.ColumnNameToIndex(mapping.DateColumn) + 1;
        var amountCol = ExcelColumnHelper.ColumnNameToIndex(mapping.AmountColumn) + 1;
        var cardCol = TryGetColumnIndex(mapping.CardColumn, oneBased: true);
        var descCol = TryGetColumnIndex(mapping.DescriptionColumn, oneBased: true);
        var acc1Col = TryGetColumnIndex(mapping.Account1Column, oneBased: true);
        var acc2Col = TryGetColumnIndex(mapping.Account2Column, oneBased: true);

        for (var row = startRow; row <= worksheet.Dimension.End.Row; row++)
        {
            var transaction = new Transaction();

            if (dateCol <= worksheet.Dimension.End.Column)
            {
                var value = worksheet.Cells[row, dateCol].Value;
                if (value != null)
                {
                    transaction.Date = NormalizeDate(value);
                }
            }

            if (amountCol <= worksheet.Dimension.End.Column)
            {
                var value = worksheet.Cells[row, amountCol].Value;
                if (value != null)
                {
                    transaction.Amount = NormalizeAmount(value);
                }
            }

            if (cardCol.HasValue && cardCol.Value <= worksheet.Dimension.End.Column)
            {
                var value = worksheet.Cells[row, cardCol.Value].Value;
                if (value != null)
                {
                    transaction.CardNumber = ExtractCardTail(value);
                }
            }

            if (descCol.HasValue && descCol.Value <= worksheet.Dimension.End.Column)
            {
                var value = worksheet.Cells[row, descCol.Value].Value;
                if (value != null)
                {
                    transaction.Description = value.ToString();
                }
            }

            if (acc1Col.HasValue && acc1Col.Value <= worksheet.Dimension.End.Column)
            {
                var value = worksheet.Cells[row, acc1Col.Value].Value;
                if (value != null)
                {
                    transaction.Account1 = value.ToString();
                }
            }

            if (acc2Col.HasValue && acc2Col.Value <= worksheet.Dimension.End.Column)
            {
                var value = worksheet.Cells[row, acc2Col.Value].Value;
                if (value != null)
                {
                    transaction.Account2 = value.ToString();
                }
            }

            if (transaction.Date != default || transaction.Amount != 0)
            {
                transactions.Add(transaction);
            }
        }

        return transactions;
    }

    private List<Transaction> LoadXls(FileInfo fileInfo, ExcelMapping mapping)
    {
        var transactions = new List<Transaction>();

        using var stream = fileInfo.OpenRead();
        IWorkbook workbook = new HSSFWorkbook(stream);
        if (workbook.NumberOfSheets == 0)
        {
            throw new Exception("Excel 文件中没有工作表。");
        }

        var sheet = workbook.GetSheetAt(0);
        if (sheet == null || sheet.LastRowNum < 0)
        {
            return transactions;
        }

        var headerRowIndex = FindHeaderRow(sheet);
        var startRowIndex = headerRowIndex + 1;

        var dateCol = ExcelColumnHelper.ColumnNameToIndex(mapping.DateColumn);
        var amountCol = ExcelColumnHelper.ColumnNameToIndex(mapping.AmountColumn);
        var cardCol = TryGetColumnIndex(mapping.CardColumn, oneBased: false);
        var descCol = TryGetColumnIndex(mapping.DescriptionColumn, oneBased: false);
        var acc1Col = TryGetColumnIndex(mapping.Account1Column, oneBased: false);
        var acc2Col = TryGetColumnIndex(mapping.Account2Column, oneBased: false);

        for (var rowIndex = startRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                continue;
            }

            var transaction = new Transaction();

            var dateValue = GetCellValue(row.GetCell(dateCol));
            if (dateValue != null)
            {
                transaction.Date = NormalizeDate(dateValue);
            }

            var amountValue = GetCellValue(row.GetCell(amountCol));
            if (amountValue != null)
            {
                transaction.Amount = NormalizeAmount(amountValue);
            }

            if (cardCol.HasValue)
            {
                var cardValue = GetCellValue(row.GetCell(cardCol.Value));
                if (cardValue != null)
                {
                    transaction.CardNumber = ExtractCardTail(cardValue);
                }
            }

            if (descCol.HasValue)
            {
                var descValue = GetCellValue(row.GetCell(descCol.Value));
                if (descValue != null)
                {
                    transaction.Description = descValue.ToString();
                }
            }

            if (acc1Col.HasValue)
            {
                var acc1Value = GetCellValue(row.GetCell(acc1Col.Value));
                if (acc1Value != null)
                {
                    transaction.Account1 = acc1Value.ToString();
                }
            }

            if (acc2Col.HasValue)
            {
                var acc2Value = GetCellValue(row.GetCell(acc2Col.Value));
                if (acc2Value != null)
                {
                    transaction.Account2 = acc2Value.ToString();
                }
            }

            if (transaction.Date != default || transaction.Amount != 0)
            {
                transactions.Add(transaction);
            }
        }

        return transactions;
    }

    private (DateTime? MinDate, DateTime? MaxDate) GetDateRangeFromOpenXml(FileInfo fileInfo, string dateColumn)
    {
        DateTime? minDate = null;
        DateTime? maxDate = null;

        using var package = new ExcelPackage(fileInfo);
        if (package.Workbook.Worksheets.Count == 0)
        {
            return (null, null);
        }

        var worksheet = package.Workbook.Worksheets[0];
        if (worksheet.Dimension == null)
        {
            return (null, null);
        }

        var dateCol = ExcelColumnHelper.ColumnNameToIndex(dateColumn) + 1;
        if (dateCol <= 0 || dateCol > worksheet.Dimension.End.Column)
        {
            return (null, null);
        }

        var headerRow = FindHeaderRow(worksheet);
        var startRow = headerRow + 1;

        for (var row = startRow; row <= worksheet.Dimension.End.Row; row++)
        {
            var cellValue = worksheet.Cells[row, dateCol].Value;
            if (cellValue == null)
            {
                continue;
            }

            var date = NormalizeDate(cellValue);
            if (date == default)
            {
                continue;
            }

            if (!minDate.HasValue || date < minDate.Value)
            {
                minDate = date;
            }

            if (!maxDate.HasValue || date > maxDate.Value)
            {
                maxDate = date;
            }
        }

        return (minDate, maxDate);
    }

    private (DateTime? MinDate, DateTime? MaxDate) GetDateRangeFromXls(FileInfo fileInfo, string dateColumn)
    {
        DateTime? minDate = null;
        DateTime? maxDate = null;

        using var stream = fileInfo.OpenRead();
        IWorkbook workbook = new HSSFWorkbook(stream);
        if (workbook.NumberOfSheets == 0)
        {
            return (null, null);
        }

        var sheet = workbook.GetSheetAt(0);
        if (sheet == null || sheet.LastRowNum < 0)
        {
            return (null, null);
        }

        var dateCol = ExcelColumnHelper.ColumnNameToIndex(dateColumn);
        var headerRowIndex = FindHeaderRow(sheet);
        var startRowIndex = headerRowIndex + 1;

        for (var rowIndex = startRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                continue;
            }

            var value = GetCellValue(row.GetCell(dateCol));
            if (value == null)
            {
                continue;
            }

            var date = NormalizeDate(value);
            if (date == default)
            {
                continue;
            }

            if (!minDate.HasValue || date < minDate.Value)
            {
                minDate = date;
            }

            if (!maxDate.HasValue || date > maxDate.Value)
            {
                maxDate = date;
            }
        }

        return (minDate, maxDate);
    }

    private static object? GetCellValue(ICell? cell)
    {
        if (cell == null)
        {
            return null;
        }

        var cellType = cell.CellType == CellType.Formula
            ? cell.CachedFormulaResultType
            : cell.CellType;

        return cellType switch
        {
            CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? cell.DateCellValue
                : cell.NumericCellValue,
            CellType.String => cell.StringCellValue,
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Blank => null,
            CellType.Error => null,
            _ => cell.ToString()
        };
    }

    private static bool IsXls(string extension) =>
        string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase);

    private static int? TryGetColumnIndex(string? columnName, bool oneBased)
    {
        if (string.IsNullOrWhiteSpace(columnName))
        {
            return null;
        }

        var index = ExcelColumnHelper.ColumnNameToIndex(columnName);
        return oneBased ? index + 1 : index;
    }

    private static bool IsHeaderKeyword(string? cellValue)
    {
        if (string.IsNullOrWhiteSpace(cellValue))
        {
            return false;
        }

        return cellValue.Contains("交易日期", StringComparison.OrdinalIgnoreCase)
               || cellValue.Contains("日期", StringComparison.OrdinalIgnoreCase)
               || cellValue.Contains("时间", StringComparison.OrdinalIgnoreCase)
               || cellValue.Contains("金额", StringComparison.OrdinalIgnoreCase);
    }
}
