using System.IO;
using System.Text;
using BillMatch.Wpf.Models;
using BillMatch.Wpf.Services;
using NPOI.HSSF.UserModel;
using OfficeOpenXml;
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
        var result = _service.NormalizeDate(input);

        Assert.Equal(new DateTime(year, month, day), result);
        Assert.Equal(TimeSpan.Zero, result.TimeOfDay);
    }

    [Fact]
    public void NormalizeDate_WithDateTime_ShouldReturnDateOnly()
    {
        var input = new DateTime(2023, 10, 1, 15, 30, 0);

        var result = _service.NormalizeDate(input);

        Assert.Equal(new DateTime(2023, 10, 1), result);
    }

    [Fact]
    public void NormalizeDate_WithExcelSerialNumber_ShouldReturnDateOnly()
    {
        const double excelSerial = 46066.5208333333d;

        var result = _service.NormalizeDate(excelSerial);

        Assert.Equal(new DateTime(2026, 2, 13), result);
    }

    [Theory]
    [InlineData("￥1,234.56", 1234.56)]
    [InlineData("1234.56", 1234.56)]
    [InlineData("-1234.56", -1234.56)]
    [InlineData("¥-1,234.56", -1234.56)]
    [InlineData(1234.56, 1234.56)]
    public void NormalizeAmount_ShouldHandleVariousFormats(object input, double expected)
    {
        var result = _service.NormalizeAmount(input);

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
        var result = _service.ExtractCardTail(input);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void LoadExcel_WithQianJiCsvAndDefaultMapping_ShouldReadRows()
    {
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.csv");

        try
        {
            var csvContent = string.Join(Environment.NewLine, new[]
            {
                "\"ID\",\"时间\",\"分类\",\"二级分类\",\"类型\",\"金额\",\"币种\",\"账户1\",\"账户2\",\"备注\"",
                "\"qj-1\",\"2026-02-12 12:21:10\",\"伙食\",\"三餐\",\"支出\",\"14.0\",\"CNY\",\"中信 8820\",,\"重庆小面\"",
                "\"qj-2\",\"2026-02-12 09:20:55\",\"日常\",\"付费会员\",\"支出\",\"8.06\",\"CNY\",\"中信 8820\",,\"火山引擎\""
            });

            File.WriteAllText(tempFilePath, csvContent, new UTF8Encoding(true));

            var mapping = new ExcelMapping
            {
                DateColumn = "B",
                AmountColumn = "F",
                Account1Column = "H",
                Account2Column = "I",
                DescriptionColumn = "J"
            };

            var result = _service.LoadExcel(tempFilePath, mapping);

            Assert.Equal(2, result.Count);
            Assert.Equal(new DateTime(2026, 2, 12), result[0].Date);
            Assert.Equal(14.0m, result[0].Amount);
            Assert.Equal("中信 8820", result[0].Account1);
            Assert.Equal("重庆小面", result[0].Description);
            Assert.Null(result[0].Account2);
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
    public void GetDateRange_WithCsvFile_ShouldReturnMinAndMaxDate()
    {
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.csv");

        try
        {
            var csvContent = string.Join(Environment.NewLine, new[]
            {
                "\"ID\",\"时间\",\"分类\",\"二级分类\",\"类型\",\"金额\",\"币种\",\"账户1\",\"账户2\",\"备注\"",
                "\"qj-1\",\"2026-02-11 19:44:00\",\"购物\",\"数码\",\"支出\",\"120.81\",\"CNY\",\"招商 0273\",,\"OPPO Find X8\"",
                "\"qj-2\",\"2026-02-13 09:20:55\",\"日常\",\"付费会员\",\"支出\",\"8.06\",\"CNY\",\"中信 8820\",,\"火山引擎\""
            });

            File.WriteAllText(tempFilePath, csvContent, new UTF8Encoding(true));

            var (minDate, maxDate) = _service.GetDateRange(tempFilePath, "B");

            Assert.Equal(new DateTime(2026, 2, 11), minDate);
            Assert.Equal(new DateTime(2026, 2, 13), maxDate);
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
    public void LoadExcel_WithXlsFile_ShouldReadRows()
    {
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xls");

        try
        {
            var workbook = new HSSFWorkbook();
            var worksheet = workbook.CreateSheet("Sheet1");

            var header = worksheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("ID");
            header.CreateCell(1).SetCellValue("时间");
            header.CreateCell(5).SetCellValue("金额");
            header.CreateCell(7).SetCellValue("账户1");
            header.CreateCell(8).SetCellValue("账户2");
            header.CreateCell(9).SetCellValue("备注");

            var row = worksheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("qj-1");
            row.CreateCell(1).SetCellValue("2026-02-13 12:30:00");
            row.CreateCell(5).SetCellValue(14.25d);
            row.CreateCell(7).SetCellValue("中信 8820");
            row.CreateCell(8).SetCellValue(string.Empty);
            row.CreateCell(9).SetCellValue("午餐");

            using (var stream = File.Create(tempFilePath))
            {
                workbook.Write(stream);
            }

            var mapping = new ExcelMapping
            {
                DateColumn = "B",
                AmountColumn = "F",
                Account1Column = "H",
                Account2Column = "I",
                DescriptionColumn = "J"
            };

            var result = _service.LoadExcel(tempFilePath, mapping);

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
    public void GetDateRange_WithXlsFile_ShouldReturnMinAndMaxDate()
    {
        var tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xls");

        try
        {
            var workbook = new HSSFWorkbook();
            var worksheet = workbook.CreateSheet("Sheet1");

            var header = worksheet.CreateRow(0);
            header.CreateCell(1).SetCellValue("交易日期");

            var row1 = worksheet.CreateRow(1);
            row1.CreateCell(1).SetCellValue("2026-02-11");

            var row2 = worksheet.CreateRow(2);
            row2.CreateCell(1).SetCellValue("2026-02-13");

            using (var stream = File.Create(tempFilePath))
            {
                workbook.Write(stream);
            }

            var (minDate, maxDate) = _service.GetDateRange(tempFilePath, "B");

            Assert.Equal(new DateTime(2026, 2, 11), minDate);
            Assert.Equal(new DateTime(2026, 2, 13), maxDate);
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
                worksheet.Cells[2, 9].Value = string.Empty;
                worksheet.Cells[2, 10].Value = "午餐";

                package.SaveAs(new FileInfo(tempFilePath));
            }

            var mapping = new ExcelMapping
            {
                DateColumn = "B",
                AmountColumn = "F",
                Account1Column = "H",
                Account2Column = "I",
                DescriptionColumn = "J"
            };

            var result = _service.LoadExcel(tempFilePath, mapping);

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
                DateColumn = "A",
                AmountColumn = "B",
                Account1Column = "C",
                Account2Column = "D",
                DescriptionColumn = "E"
            };

            var result = _service.LoadExcel(tempFilePath, oldMapping);

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
