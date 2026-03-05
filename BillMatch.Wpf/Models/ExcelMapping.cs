namespace BillMatch.Wpf.Models;

/// <summary>
/// Excel列映射配置 - 使用Excel列名(A,B,C...)
/// </summary>
public class ExcelMapping
{
    /// <summary>
    /// 日期列 (如 "A", "B")
    /// </summary>
    public string DateColumn { get; set; } = "A";

    /// <summary>
    /// 金额列 (如 "C", "D")
    /// </summary>
    public string AmountColumn { get; set; } = "B";

    /// <summary>
    /// 卡号列 (如 "E", "F")
    /// </summary>
    public string CardColumn { get; set; } = "C";

    /// <summary>
    /// 描述列 (如 "G", "H")
    /// </summary>
    public string DescriptionColumn { get; set; } = "D";

    /// <summary>
    /// 账户1列 (钱迹专用)
    /// </summary>
    public string Account1Column { get; set; } = "E";

    /// <summary>
    /// 账户2列 (钱迹专用)
    /// </summary>
    public string Account2Column { get; set; } = "F";
}
