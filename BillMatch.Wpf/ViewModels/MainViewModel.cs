using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using BillMatch.Wpf.Models;
using BillMatch.Wpf.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;

namespace BillMatch.Wpf.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly IExcelService _excelService;

    // 文件名属性
    [ObservableProperty]
    private string? _qianjiFileName;

    [ObservableProperty]
    private string? _billFileName;

    // 实际文件路径
    private string? _qianjiFilePath;
    private List<string> _billFilePaths = new();

    // 对账设置属性
    [ObservableProperty]
    private DateTime _startDate = DateTime.Today.AddDays(-(DateTime.Today.Day - 1));

    [ObservableProperty]
    private DateTime _endDate = DateTime.Today;

    [ObservableProperty]
    private string? _targetCard = "8820";

    [ObservableProperty]
    private int _daysTolerance = 2;

    // 列映射配置 - 钱迹 (使用Excel列名如 "B", "F")
    [ObservableProperty]
    private string _qianjiDateColumn = "B";

    [ObservableProperty]
    private string _qianjiAmountColumn = "F";

    [ObservableProperty]
    private string _qianjiAccount1Column = "H";

    [ObservableProperty]
    private string _qianjiAccount2Column = "I";

    [ObservableProperty]
    private string _qianjiDescriptionColumn = "J";

    // 列映射配置 - 账单 (使用Excel列名如 "A", "G")
    [ObservableProperty]
    private string _billDateColumn = "A";

    [ObservableProperty]
    private string _billAmountColumn = "G";

    [ObservableProperty]
    private string _billCardColumn = "D";

    [ObservableProperty]
    private string _billDescriptionColumn = "C";

    // 结果集合属性
    [ObservableProperty]
    private ObservableCollection<Transaction> _unmatchedBills = new();

    [ObservableProperty]
    private ObservableCollection<MatchResult> _matchedPairs = new();

    [ObservableProperty]
    private ObservableCollection<Transaction> _unmatchedQianji = new();

    [ObservableProperty]
    private string _logText = string.Empty;

    // 存储加载的数据
    private List<Transaction> _qianjiTransactions = new();
    private List<Transaction> _billTransactions = new();

    public MainViewModel(IExcelService excelService)
    {
        _excelService = excelService;
    }

    public MainViewModel() : this(new ExcelService())
    {
    }

    [RelayCommand]
    private void SelectQianjiFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel 文件 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|所有文件 (*.*)|*.*",
            Title = "选择钱迹导出文件"
        };

        if (dialog.ShowDialog() == true)
        {
            _qianjiFilePath = dialog.FileName;
            QianjiFileName = Path.GetFileName(dialog.FileName);
            AppendLog($"已选择钱迹文件: {_qianjiFilePath}");

            if (string.Equals(Path.GetExtension(_qianjiFilePath), ".xls", StringComparison.OrdinalIgnoreCase))
            {
                AppendLog("警告：当前钱迹文件为 .xls 旧格式，EPPlus 无法解析。请先另存为 .xlsx 后再导入。");
            }
        }
    }

    [RelayCommand]
    private void SelectBillFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel 文件 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|所有文件 (*.*)|*.*",
            Title = "选择银行账单文件",
            Multiselect = true
        };

        if (dialog.ShowDialog() == true)
        {
            _billFilePaths = dialog.FileNames.ToList();
            BillFileName = dialog.FileNames.Length > 1
                ? $"{Path.GetFileName(dialog.FileNames[0])} 等 {dialog.FileNames.Length} 个文件"
                : Path.GetFileName(dialog.FileNames[0]);

            AppendLog($"已选择银行账单文件数量: {_billFilePaths.Count}");
            foreach (var path in _billFilePaths)
            {
                AppendLog($"  - {path}");

                if (string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase))
                {
                    AppendLog("  警告：检测到 .xls 旧格式文件，EPPlus 无法解析。请先另存为 .xlsx 后再导入。");
                }
            }

            // 自动读取日期范围
            AutoDetectDateRange();
        }
    }

    /// <summary>
    /// 自动检测银行账单的日期范围并设置
    /// </summary>
    private void AutoDetectDateRange()
    {
        if (_billFilePaths.Count == 0) return;

        try
        {
            AppendLog("正在自动读取银行账单的日期范围...");
            
            DateTime? overallMinDate = null;
            DateTime? overallMaxDate = null;

            foreach (var filePath in _billFilePaths)
            {
                var (minDate, maxDate) = _excelService.GetDateRange(filePath, BillDateColumn);
                
                if (minDate.HasValue && (!overallMinDate.HasValue || minDate.Value < overallMinDate.Value))
                {
                    overallMinDate = minDate;
                }
                
                if (maxDate.HasValue && (!overallMaxDate.HasValue || maxDate.Value > overallMaxDate.Value))
                {
                    overallMaxDate = maxDate;
                }
            }

            if (overallMinDate.HasValue && overallMaxDate.HasValue)
            {
                StartDate = overallMinDate.Value;
                EndDate = overallMaxDate.Value;
                AppendLog($"已自动设置日期范围: {StartDate:yyyy-MM-dd} 至 {EndDate:yyyy-MM-dd}");
            }
            else
            {
                AppendLog("未能从银行账单中读取到有效的日期范围，请手动设置。");
            }
        }
        catch (Exception ex)
        {
            AppendLog($"读取日期范围时出错: {ex.Message}");
        }
    }

    [RelayCommand]
    private void CopyLog()
    {
        if (string.IsNullOrWhiteSpace(LogText))
        {
            MessageBox.Show("当前没有可复制的日志。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            Clipboard.SetText(LogText);
            MessageBox.Show("日志已复制到剪贴板。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"复制日志失败: {ex}");
            MessageBox.Show($"复制日志失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ClearLog()
    {
        LogText = string.Empty;
    }

    [RelayCommand]
    private async Task StartMatchAsync()
    {
        AppendLog("开始执行对账。\n");

        // 验证文件已选择
        if (string.IsNullOrEmpty(_qianjiFilePath) || _billFilePaths.Count == 0)
        {
            AppendLog("未选择完整文件：请先选择钱迹文件和银行账单文件。\n");
            MessageBox.Show("请先选择钱迹文件和银行账单文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            // 1. 加载 Excel 文件（使用配置的列映射）
            var qianjiMapping = new ExcelMapping
            {
                DateColumn = QianjiDateColumn,
                AmountColumn = QianjiAmountColumn,
                Account1Column = QianjiAccount1Column,
                Account2Column = QianjiAccount2Column,
                DescriptionColumn = QianjiDescriptionColumn
            };

            var billMapping = new ExcelMapping
            {
                DateColumn = BillDateColumn,
                AmountColumn = BillAmountColumn,
                CardColumn = BillCardColumn,
                DescriptionColumn = BillDescriptionColumn
            };

            AppendLog($"当前钱迹列映射: 日期={QianjiDateColumn}, 金额={QianjiAmountColumn}, 账户1={QianjiAccount1Column}, 账户2={QianjiAccount2Column}, 描述={QianjiDescriptionColumn}");
            AppendLog($"当前账单列映射: 日期={BillDateColumn}, 金额={BillAmountColumn}, 卡号={BillCardColumn}, 描述={BillDescriptionColumn}");

            var loadingLog = new StringBuilder();

            // 异步加载文件
            await Task.Run(() =>
            {
                loadingLog.AppendLine($"开始读取钱迹文件: {_qianjiFilePath}");
                _qianjiTransactions = _excelService.LoadExcel(_qianjiFilePath!, qianjiMapping);
                loadingLog.AppendLine($"钱迹文件读取完成，记录数: {_qianjiTransactions.Count}");
                
                _billTransactions.Clear();
                foreach (var path in _billFilePaths)
                {
                    loadingLog.AppendLine($"开始读取银行文件: {path}");
                    var transactions = _excelService.LoadExcel(path, billMapping);
                    _billTransactions.AddRange(transactions);
                    loadingLog.AppendLine($"银行文件读取完成，新增记录数: {transactions.Count}");
                }

                loadingLog.AppendLine($"银行文件汇总记录数: {_billTransactions.Count}");
            });

            AppendLog(loadingLog.ToString().TrimEnd());

            if (_qianjiTransactions.Count == 0 && _billTransactions.Count == 0)
            {
                AppendLog("钱迹与账单均为 0 条，建议优先检查列映射是否匹配实际导出格式。钱迹常见映射：日期=B、金额=F、账户1=H、账户2=I、备注=J。");
                AppendLog("读取完成，但未从文件中提取到交易数据。\n");
                MessageBox.Show("未从所选文件中读取到任何交易数据。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (_qianjiTransactions.Count == 0)
            {
                AppendLog("钱迹读取结果为 0 条，可能是列映射不匹配，或日期/金额列格式无法识别。请核对钱迹列映射。\n");
            }

            // 2. 按日期范围过滤
            var filteredQianji = _qianjiTransactions
                .Where(t => t.Date >= StartDate.Date && t.Date <= EndDate.Date)
                .ToList();

            var filteredBills = _billTransactions
                .Where(t => t.Date >= StartDate.Date && t.Date <= EndDate.Date)
                .ToList();

            // 3. 按目标卡号过滤
            if (!string.IsNullOrEmpty(TargetCard))
            {
                var targetCard = ExtractCardTail(TargetCard);
                if (!string.IsNullOrEmpty(targetCard))
                {
                    filteredQianji = filteredQianji
                        .Where(t => ExtractCardTail(t.Account1) == targetCard ||
                                    ExtractCardTail(t.Account2) == targetCard)
                        .ToList();

                    filteredBills = filteredBills
                        .Where(t => ExtractCardTail(t.CardNumber) == targetCard)
                        .ToList();

                    AppendLog($"按卡号 {targetCard} 过滤后：钱迹 {filteredQianji.Count} 条，银行 {filteredBills.Count} 条。");
                }
            }

            // 4. 执行匹配算法
            MatchTransactions(filteredQianji, filteredBills);

            AppendLog($"对账完成：匹配成功 {MatchedPairs.Count} 笔，未匹配钱迹 {UnmatchedQianji.Count} 笔，未匹配账单 {UnmatchedBills.Count} 笔。\n");

            // 显示成功消息
            MessageBox.Show($"对账完成！\n匹配成功: {MatchedPairs.Count} 笔\n未匹配钱迹: {UnmatchedQianji.Count} 笔\n未匹配账单: {UnmatchedBills.Count} 笔", 
                "完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"对账过程中发生异常:\n{ex}\n");
            MessageBox.Show($"对账过程中出错: {ex.Message}\n\n详细错误已写入下方[运行日志]，可点击[复制日志]发送给我。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void AppendLog(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var lines = message.Replace("\r\n", "\n").Split('\n');
        var builder = new StringBuilder();

        foreach (var rawLine in lines)
        {
            var line = rawLine.TrimEnd();
            if (line.Length == 0)
            {
                continue;
            }

            builder.Append('[')
                .Append(now)
                .Append("] ")
                .AppendLine(line);
        }

        var formatted = builder.ToString().TrimEnd();
        if (formatted.Length == 0)
        {
            return;
        }

        LogText = string.IsNullOrEmpty(LogText)
            ? formatted
            : $"{LogText}{Environment.NewLine}{formatted}";
    }

    internal void MatchTransactions(List<Transaction> qianjiList, List<Transaction> billList)
    {
        var matchedPairs = new List<MatchResult>();
        
        var qianjiPool = qianjiList.Where(t => t.Amount != 0).ToList();

        foreach (var qianji in qianjiPool)
        {
            qianji.IsMatched = false;
        }

        foreach (var bill in billList)
        {
            if (bill.Amount == 0)
            {
                continue;
            }

            var billDate = bill.Date.Date;
            var billAmountAbs = Math.Abs(bill.Amount);
            var minDate = billDate.AddDays(-DaysTolerance);
            var maxDate = billDate.AddDays(DaysTolerance);

            var candidates = qianjiPool
                .Where(q => !q.IsMatched &&
                            Math.Abs(Math.Abs(q.Amount) - billAmountAbs) <= 0.01m &&
                            q.Date.Date >= minDate &&
                            q.Date.Date <= maxDate)
                .ToList();

            if (candidates.Any())
            {
                var bestMatch = candidates
                    .OrderBy(q => Math.Abs((q.Date.Date - billDate).Days))
                    .First();

                bestMatch.IsMatched = true;

                matchedPairs.Add(new MatchResult
                {
                    BillTransaction = bill,
                    QianjiTransaction = bestMatch,
                    Status = "已匹配"
                });
            }
        }

        var unmatchedBillsList = billList
            .Where(b => !matchedPairs.Any(mp => mp.BillTransaction == b))
            .OrderByDescending(b => b.Date)
            .ToList();
            
        UnmatchedBills = new ObservableCollection<Transaction>(unmatchedBillsList);

        MatchedPairs = new ObservableCollection<MatchResult>(
            matchedPairs.OrderByDescending(mp => mp.BillTransaction.Date)
        );

        var unmatchedQianjiList = qianjiPool
            .Where(q => !q.IsMatched)
            .OrderByDescending(q => q.Date)
            .ToList();
            
        UnmatchedQianji = new ObservableCollection<Transaction>(unmatchedQianjiList);
    }

    private string? ExtractCardTail(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        var digits = new string(value.Where(char.IsDigit).ToArray());

        if (digits.Length >= 4)
        {
            return digits.Substring(digits.Length - 4);
        }

        return null;
    }
}
