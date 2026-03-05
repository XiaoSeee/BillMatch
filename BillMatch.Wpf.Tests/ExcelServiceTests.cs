using BillMatch.Wpf.Models;
using BillMatch.Wpf.Services;
using OfficeOpenXml;
using System.IO;
using Xunit;

namespace BillMatch.Wpf.Tests;

public class ExcelServiceTests
{
    private readonly ExcelService _service = new();

    [Theory]
    [InlineData("2023-10-01 12:34:56", 2023, 10, 1)]
    [InlineData("2023/10/02", 2023, 10, 2)]
    [InlineData("2023-10-03", 2023, 10, 3)]
    public void NormalizeDate_ShouldReturnDateOnly(string input, int year, int month, int day)
    {
        // Act
        var result = _service.NormalizeDate(input);

        // Assert
        Assert.Equal(new DateTime(year, month, day), result);
        Assert.Equal(TimeSpan.Zero, result.TimeOfDay);
    }

    [Fact]
    public void NormalizeDate_WithDateTime_ShouldReturnDateOnly()
    {
        // Arrange
        var input = new DateTime(2023, 10, 1, 15, 30, 0);

        // Act
        var result = _service.NormalizeDate(input);

        // Assert
        Assert.Equal(new DateTime(2023, 10, 1), result);
    }

    [Fact]
    public void NormalizeDate_WithExcelSerialNumber_ShouldReturnDateOnly()
    {
        // Arrange
        const double excelSerial = 46066.5208333333d; // 2026-02-13 12:30:00

        // Act
        var result = _service.NormalizeDate(excelSerial);

        // Assert
        Assert.Equal(new DateTime(2026, 2, 13), result);
    }

    [Theory]
    [InlineData("￥1,234.56", 1234.56)]
    [InlineData("1234.56", 1234.56)]
    [InlineData("-1234.56", -1234.56)]
    [InlineData("￥-1,234.56", -1234.56)]
    [InlineData(1234.56, 1234.56)]
    public void NormalizeAmount_ShouldHandleVariousFormats(object input, double expected)
    {
        // Act
        var result = _service.NormalizeAmount(input);

        // Assert
        Assert.Equal((decimal)expected, result);
    }

    [Theory]
    [InlineData("**** **** **** 8820", "8820")]
    [InlineData("尾号8820", "8820")]
    [InlineData("8820", "8820")]
    [InlineData(8820.0, "8820")]
    [InlineData("123", null)]
    public void ExtractCardTail_ShouldExtractLastFourDigits(object input, string? expected)
    {
        // Act
        var result = _service.ExtractCardTail(input);

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void LoadExcel_WithXlsExtension_ShouldThrowNotSupportedException()
    {
        // Arrange
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xls");
        File.WriteAllText(tempFilePath, "dummy");

        try
        {
            // Act
            var ex = Assert.Throws<NotSupportedException>(() =>
                _service.LoadExcel(tempFilePath, new ExcelMapping()));

            // Assert
            Assert.Contains(".xls", ex.Message);
            Assert.Contains(".xlsx", ex.Message);
        }
        finally
        {
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
    }

    [Fact]
    public void LoadExcel_WithQianJiLikeColumnsAndCorrectMapping_ShouldReadRows()
    {
        // Arrange
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "时间";
                worksheet.Cells[1, 6].Value = "金额";
                worksheet.Cells[1, 8].Value = "账户1";
                worksheet.Cells[1, 9].Value = "账户2";
                worksheet.Cells[1, 10].Value = "备注";

                worksheet.Cells[2, 1].Value = "qj-1";
                worksheet.Cells[2, 2].Value = new DateTime(2026, 2, 13, 12, 30, 0);
                worksheet.Cells[2, 2].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
                worksheet.Cells[2, 6].Value = 14.25m;
                worksheet.Cells[2, 8].Value = "中信 8820";
                worksheet.Cells[2, 9].Value = "";
                worksheet.Cells[2, 10].Value = "午餐";

                package.SaveAs(new FileInfo(tempFilePath));
            }

            var mapping = new ExcelMapping
            {
                DateColumn = 2,
                AmountColumn = 6,
                Account1Column = 8,
                Account2Column = 9,
                DescriptionColumn = 10
            };

            // Act
            var result = _service.LoadExcel(tempFilePath, mapping);

            // Assert
            Assert.Single(result);
            Assert.Equal(new DateTime(2026, 2, 13), result[0].Date);
            Assert.Equal(14.25m, result[0].Amount);
            Assert.Equal("中信 8820", result[0].Account1);
            Assert.Equal("午餐", result[0].Description);
        }
        finally
        {
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
    }

    [Fact]
    public void LoadExcel_WithQianJiLikeColumnsAndOldMapping_ShouldReturnZeroRows()
    {
        // Arrange
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "时间";
                worksheet.Cells[1, 6].Value = "金额";

                worksheet.Cells[2, 1].Value = "qj-1";
                worksheet.Cells[2, 2].Value = "2026-02-13 12:30:00";
                worksheet.Cells[2, 6].Value = 14.25m;

                package.SaveAs(new FileInfo(tempFilePath));
            }

            var oldMapping = new ExcelMapping
            {
                DateColumn = 1,
                AmountColumn = 2,
                Account1Column = 3,
                Account2Column = 4,
                DescriptionColumn = 5
            };

            // Act
            var result = _service.LoadExcel(tempFilePath, oldMapping);

            // Assert
            Assert.Empty(result);
        }
        finally
        {
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
    }
}
