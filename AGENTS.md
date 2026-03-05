# BillMatch.Wpf - AGENTS.md

> AI 代理项目知识库
> 最后更新: 2026-03-03

---

## 1. 项目概述

BillMatch.Wpf 是一个 **Windows 桌面账单对账工具**，用于比较 **钱迹记账软件** 和 **银行账单** 的差异，帮助用户发现漏记或错误的账单。

### 核心功能
- 📊 Excel 文件读取（钱迹导出 + 银行账单）
- 🔄 智能匹配算法（金额 + 日期容差）
- 📅 自动日期范围检测
- 🏷️ 卡号过滤（末四位匹配）
- 📝 详细日志记录

---

## 2. 技术栈

| 类别 | 技术 | 版本 | 用途 |
|------|------|------|------|
| 框架 | .NET | 8.0 | 运行时 |
| UI | WPF | - | 桌面界面 |
| MVVM | CommunityToolkit.Mvvm | 8.2.2 | MVVM 模式 |
| Excel | EPPlus | 4.5.3.3 | Excel 处理 |
| 日志 | NLog | 5.3.2 | 日志记录 |
| DI | Microsoft.Extensions.DependencyInjection | 8.0.0 | 依赖注入 |

### 重要说明
- **EPPlus 使用 4.5.3.3 版本**（LGPL 许可证），不能使用 5.x+（商业许可证）
- **目标框架**: `net8.0-windows`
- **输出类型**: `WinExe`

---

## 3. 项目结构

```
BillMatch/
├── BillMatch.Wpf/                    # 主项目
│   ├── Models/                       # 数据模型
│   │   ├── Transaction.cs            # 交易记录
│   │   ├── ExcelMapping.cs           # Excel 列映射
│   │   ├── MatchResult.cs            # 匹配结果
│   │   └── MatchResults.cs           # 匹配结果集合
│   ├── ViewModels/                   # MVVM 视图模型
│   │   └── MainViewModel.cs          # 主视图模型
│   ├── Services/                     # 业务服务
│   │   ├── IExcelService.cs          # Excel 服务接口
│   │   ├── ExcelService.cs           # Excel 服务实现
│   │   ├── ILoggingService.cs        # 日志服务接口
│   │   ├── LoggingService.cs         # 日志服务实现
│   │   └── ExcelColumnHelper.cs      # Excel 列名辅助类
│   ├── MainWindow.xaml               # 主窗口 XAML
│   ├── MainWindow.xaml.cs            # 主窗口代码
│   ├── App.xaml                      # 应用资源
│   ├── App.xaml.cs                   # 应用代码
│   └── BillMatch.Wpf.csproj          # 项目文件
├── BillMatch.Wpf.Tests/              # 测试项目
│   └── ExcelServiceTests.cs          # Excel 服务测试
├── app.py                             # 原 Python 版本（参考）
├── README.md                          # 项目说明
└── publish.bat                        # 发布脚本
```

---

## 4. 核心类说明

### 4.1 数据模型 (Models)

#### Transaction
交易记录实体，用于表示钱迹和银行账单的交易。

```csharp
public class Transaction
{
    public DateTime Date { get; set; }           // 交易日期
    public decimal Amount { get; set; }          // 交易金额
    public string? CardNumber { get; set; }       // 卡号（末四位）
    public string? Description { get; set; }      // 交易描述
    public string? Account1 { get; set; }         // 账户1（钱迹）
    public string? Account2 { get; set; }         // 账户2（钱迹）
    public bool IsMatched { get; set; }          // 是否已匹配
}
```

#### ExcelMapping
Excel 列映射配置，使用 **Excel 列名（A, B, C...）** 而非数字索引。

```csharp
public class ExcelMapping
{
    public string DateColumn { get; set; } = "A";        // 日期列（如 "B"）
    public string AmountColumn { get; set; } = "B";      // 金额列（如 "F"）
    public string CardColumn { get; set; } = "C";        // 卡号列（如 "D"）
    public string DescriptionColumn { get; set; } = "D"; // 描述列（如 "J"）
    public string Account1Column { get; set; } = "E";    // 账户1列
    public string Account2Column { get; set; } = "F";    // 账户2列
}
```

### 4.2 服务 (Services)

#### ExcelColumnHelper
Excel 列名与索引转换辅助类。

```csharp
public static class ExcelColumnHelper
{
    // 将 Excel 列名（如 "A", "B", "AA"）转换为 0-based 索引
    public static int ColumnNameToIndex(string columnName)
    
    // 将 0-based 索引转换为 Excel 列名
    public static string IndexToColumnName(int index)
    
    // 验证列名是否有效
    public static bool IsValidColumnName(string columnName)
    
    // 规范化用户输入（支持 "A", "第A列", "1" 等格式）
    public static string? NormalizeColumnName(string? input)
}
```

#### ExcelService
Excel 文件读取服务，基于 EPPlus。

```csharp
public class ExcelService : IExcelService
{
    // 从 Excel 加载交易数据
    public List<Transaction> LoadExcel(string filePath, ExcelMapping mapping)
    
    // 读取银行账单的日期范围（用于自动设置）
    public (DateTime? MinDate, DateTime? MaxDate) GetDateRange(string filePath, string dateColumn)
    
    // 内部方法：查找表头行
    internal int FindHeaderRow(ExcelWorksheet worksheet)
    
    // 内部方法：日期归一化（处理多种格式）
    internal DateTime NormalizeDate(object value)
    
    // 内部方法：金额归一化（去除货币符号）
    internal decimal NormalizeAmount(object value)
    
    // 内部方法：提取卡号末四位
    internal string? ExtractCardTail(object value)
}
```

#### LoggingService（未实际使用）
基于 NLog 的日志服务，已实现但当前使用 UI 内置日志。

### 4.3 视图模型 (ViewModels)

#### MainViewModel
主视图模型，处理所有业务逻辑。

**核心属性：**
- `QianjiFileName`, `BillFileName` - 选择的文件名
- `StartDate`, `EndDate` - 对账日期范围
- `TargetCard` - 目标卡号（末四位）
- `DaysTolerance` - 日期容差（天）
- 列映射配置（使用 Excel 列名如 "B", "F"）
- `UnmatchedBills`, `MatchedPairs`, `UnmatchedQianji` - 对账结果
- `LogText` - 运行日志

**核心命令：**
- `SelectQianjiFileCommand` - 选择钱迹文件
- `SelectBillFileCommand` - 选择银行账单（支持多选，自动检测日期范围）
- `StartMatchCommand` - 执行对账
- `CopyLogCommand`, `ClearLogCommand` - 日志操作

**核心方法：**
- `AutoDetectDateRange()` - 自动检测银行账单的日期范围
- `MatchTransactions()` - 执行匹配算法
- `ExtractCardTail()` - 提取卡号末四位
- `AppendLog()` - 追加日志（带时间戳）

---

## 5. 关键算法

### 5.1 匹配算法

匹配逻辑基于以下规则：

1. **金额匹配**：使用绝对值比较（解决正负号差异）
   ```csharp
   Math.Abs(Math.Abs(q.Amount) - billAmountAbs) <= 0.01m
   ```

2. **日期容差**：允许配置的日期偏差（默认2天）
   ```csharp
   q.Date.Date >= minDate && q.Date.Date <= maxDate
   // 其中 minDate = billDate.AddDays(-DaysTolerance)
   ```

3. **唯一性**：每笔账单和钱迹记录只能匹配一次
   ```csharp
   qianjiPool.Where(q => !q.IsMatched)
   ```

4. **最优匹配**：选择日期最接近的候选
   ```csharp
   .OrderBy(q => Math.Abs((q.Date.Date - billDate).Days))
   .First()
   ```

### 5.2 日期范围检测

从银行账单自动检测日期范围：

```csharp
// 遍历所有银行文件，找出最早和最晚的日期
foreach (var filePath in billFilePaths)
{
    var (minDate, maxDate) = excelService.GetDateRange(filePath, dateColumn);
    // 更新总体最小和最大日期
}
```

### 5.3 卡号提取

从各种格式中提取卡号末四位：

```csharp
// 支持格式："8820", "8820.0", "中信 8820", "****8820" 等
string digits = Regex.Replace(valueStr, @"\D", "");
if (digits.Length >= 4)
{
    return digits.Substring(digits.Length - 4);
}
```

---

## 6. 开发规范

### 6.1 代码风格

- **命名规范**：
  - 类/接口：PascalCase（如 `ExcelService`）
  - 方法：PascalCase（如 `LoadExcel`）
  - 属性：PascalCase（如 `DateColumn`）
  - 私有字段：camelCase 加下划线前缀（如 `_excelService`）

- **可空性**：
  - 启用可空性（`<Nullable>enable</Nullable>`）
  - 可空引用类型使用 `?` 后缀（如 `string?`）

### 6.2 架构原则

- **MVVM 模式**：
  - Model：数据模型（`Transaction`, `ExcelMapping`）
  - View：XAML 界面（`MainWindow.xaml`）
  - ViewModel：业务逻辑（`MainViewModel`）

- **依赖注入**：
  - 使用构造函数注入
  - 服务注册在 `App.xaml.cs`

- **服务层**：
  - 接口与实现分离（`IExcelService` / `ExcelService`）
  - 单一职责原则

### 6.3 Excel 处理规范

- **列映射**：使用 Excel 列名（A, B, C...）而非数字索引
- **日期处理**：支持多种格式（OA Date, DateTime, string）
- **金额处理**：去除货币符号和千分位
- **表头检测**：自动检测表头行（前5行内查找关键词）

### 6.4 日志规范

- **日志格式**：`[yyyy-MM-dd HH:mm:ss] 消息内容`
- **日志级别**：Debug, Info, Warn, Error, Fatal
- **日志位置**：`%LocalAppData%\BillMatch\logs\`
- **日志文件**：`BillMatch_{yyyyMMdd}.log`

---

## 7. 常见任务

### 7.1 添加新的 Excel 列映射

1. 在 `ExcelMapping.cs` 中添加属性：
   ```csharp
   public string NewColumn { get; set; } = "G";
   ```

2. 在 `MainViewModel.cs` 中添加对应的 ViewModel 属性

3. 在 `MainWindow.xaml` 中添加 UI 控件

4. 在 `ExcelService.cs` 的 `LoadExcel` 方法中添加读取逻辑

### 7.2 修改匹配算法

匹配算法位于 `MainViewModel.MatchTransactions()` 方法中。

关键逻辑：
```csharp
var candidates = qianjiPool
    .Where(q => !q.IsMatched &&
                Math.Abs(Math.Abs(q.Amount) - billAmountAbs) <= 0.01m &&
                q.Date.Date >= minDate &&
                q.Date.Date <= maxDate)
    .ToList();
```

### 7.3 添加新的日志输出

在 `MainViewModel` 中使用 `AppendLog` 方法：

```csharp
AppendLog("您的日志消息");
```

日志会自动添加时间戳并显示在"运行日志"标签页中。

---

## 8. 注意事项

### 8.1 EPPlus 许可证

**重要**：项目使用 EPPlus **4.5.3.3** 版本（LGPL 许可证）。

- ✅ 可以使用 4.x 版本（LGPL，开源项目可用）
- ❌ 不要使用 5.x+ 版本（商业许可证，需要购买）

### 8.2 Excel 格式支持

- ✅ 支持 `.xlsx` 格式
- ❌ 不支持 `.xls` 旧格式（会提示用户另存为 .xlsx）

### 8.3 日期格式

支持的日期格式：
- OA Date（Excel 内部格式）
- `DateTime` 对象
- 字符串格式：`yyyy-MM-dd`, `yyyy/MM/dd` 等

### 8.4 金额格式

自动处理的格式：
- 数字（14, 8.06）
- 带货币符号（"￥14.00", "¥8.06"）
- 带千分位（"1,234.56"）

---

## 9. 参考资源

### 9.1 相关文档
- [EPPlus 4.x 文档](https://epplus.codeplex.com/)
- [WPF 官方文档](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
- [CommunityToolkit.Mvvm 文档](https://learn.microsoft.com/en-us/dotnet/communitytoolkit/mvvm/)

### 9.2 示例数据格式

#### 钱迹导出文件列结构
| 列 | Excel列 | 说明 |
|----|---------|------|
| 时间 | B | 交易日期时间 |
| 金额 | F | 交易金额 |
| 账户1 | H | 主账户/支出账户 |
| 账户2 | I | 转入账户/对方账户 |
| 备注 | J | 交易备注/描述 |

#### 银行账单文件列结构
| 列 | Excel列 | 说明 |
|----|---------|------|
| 交易日期 | A | 交易日期 |
| 交易金额 | G | 交易金额 |
| 卡末四位 | D | 卡号末四位 |
| 交易描述 | C | 交易描述/商户名 |

---

## 10. 更新日志

### 2026-03-03
- 初始版本 AGENTS.md 文档
- 完成 WPF 项目重构
- 实现 Excel 列名映射（A,B,C...）
- 实现自动日期范围检测
- 实现完善日志系统

---

**文档维护者**: AI 代理
**审核状态**: ✅ 已验证
**适用范围**: BillMatch.Wpf 项目所有 AI 代理
