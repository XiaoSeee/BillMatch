# BillMatch.Wpf

BillMatch.Wpf 是一个 Windows 桌面对账工具，用于比对**钱迹导出账单**与**银行账单**，快速找出漏记、多记或不一致的交易。

---

## 1. 核心能力

- 读取账单文件（支持 `.csv/.xls/.xlsx/.xlsm`，支持多账单文件合并）
- 可配置列映射（使用 Excel 列名：`A`、`B`、`AA`）
- 自动检测银行账单日期范围并回填到对账区间
- 按卡号末四位过滤
- 按金额与日期容差进行智能匹配
- 输出三类结果：漏记账单、已匹配、冗余记账
- 运行日志支持复制与清空

---

## 2. 技术栈

- .NET 8（`net8.0-windows`）
- WPF
- CommunityToolkit.Mvvm
- EPPlus `4.5.3.3`（LGPL）
- NLog（项目中已实现日志服务，当前主流程主要使用 UI 内置日志）

---

## 3. 目录结构

```text
BillMatch/
├── BillMatch.Wpf/               # 主程序（WPF）
│   ├── Models/                  # 数据模型
│   ├── Services/                # Excel读取/列名转换/日志服务
│   ├── ViewModels/              # 主业务逻辑（MainViewModel）
│   ├── MainWindow.xaml          # 主界面
│   └── BillMatch.Wpf.csproj
├── BillMatch.Wpf.Tests/         # xUnit 测试项目
├── app.py                       # 旧版 Python 参考实现
└── publish.bat                  # 发布脚本
```

---

## 4. 运行环境

- Windows 10/11
- .NET 8 SDK（开发/运行源码时需要）

---

## 5. 本地运行

在项目根目录执行：

```bash
dotnet run --project BillMatch.Wpf/BillMatch.Wpf.csproj
```

---

## 6. 使用说明（UI）

1. 选择钱迹导出文件（支持 `.csv/.xls/.xlsx/.xlsm`）
2. 选择一个或多个银行账单文件
3. 检查自动识别的对账日期范围
4. 设置目标卡号末四位（可留空）
5. 根据导出格式确认列映射
6. 点击“开始对账”
7. 在“漏记账单 / 已匹配 / 冗余记账 / 运行日志”标签页查看结果

---

## 7. 默认列映射（当前代码默认值）

### 钱迹文件

- 日期：`B`
- 金额：`F`
- 账户1：`H`
- 账户2：`I`
- 描述：`J`

### 银行账单文件

- 日期：`A`
- 金额：`G`
- 卡号：`D`
- 描述：`C`

> 列映射使用 Excel 列名字母，不是数字索引。
> 对 `csv` 文件同样适用，按列位置读取，不按表头名称自动匹配。

---

## 8. 匹配规则

- 金额：按绝对值比较（忽略正负方向差异）
- 日期：允许 `DaysTolerance` 天范围内匹配（默认 2 天）
- 唯一性：一条钱迹记录最多匹配一条账单
- 最优选择：候选中优先选择与账单日期最接近的一条

---

## 9. 打包发布

在项目根目录执行：

```bat
publish.bat
```

脚本会：

1. 执行 `dotnet publish`（`Release` + `win-x64` + 单文件）
2. 从发布目录复制产物
3. 在项目根目录生成 `BillMatch.exe`

发布目录原始产物路径：

`BillMatch.Wpf\bin\Release\net8.0-windows\win-x64\publish\BillMatch.Wpf.exe`

---

## 10. 测试

运行测试：

```bash
dotnet test BillMatch.Wpf.Tests/BillMatch.Wpf.Tests.csproj
```

---

## 11. 已知限制

- `.csv` 当前按钱迹默认导出格式兼容（UTF-8 BOM、逗号分隔、双引号包裹）
- `.xls` 为兼容读取模式（仅读取，不涉及写回与格式转换）
- 当前默认读取每个文件的第一个工作表
- 表头识别为启发式逻辑（前 5 行内按关键词检测）
