using BillMatch.Wpf.Models;

namespace BillMatch.Wpf.Services;

public interface IExcelService
{
    /// <summary>
    /// 从 Excel 文件加载交易数据
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <param name="mapping">列映射配置</param>
    /// <returns>交易数据列表</returns>
    List<Transaction> LoadExcel(string filePath, ExcelMapping mapping);

    /// <summary>
    /// 从银行账单文件读取日期范围
    /// </summary>
    /// <param name="filePath">银行账单文件路径</param>
    /// <param name="dateColumn">日期列名(如"A", "B")</param>
    /// <returns>(最早日期, 最晚日期)元组</returns>
    (DateTime? MinDate, DateTime? MaxDate) GetDateRange(string filePath, string dateColumn);
}
