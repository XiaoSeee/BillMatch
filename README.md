# BillMatch

BillMatch 是一个从 Python 移植到 .NET 8 WPF 的账单匹配工具。它旨在提供更快的处理速度和更好的用户体验。

## 项目特点

- **高性能**: 使用 .NET 8 构建，处理大型 Excel 文件更加高效。
- **独立运行**: 支持 Self-contained 发布，无需安装 .NET 运行时。
- **体积优化**: 启用剪裁 (Trimming) 和单文件发布，减小分发体积。

## 运行方式

1. 确保已安装 [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)。
2. 在项目根目录下运行：
   ```bash
   dotnet run --project BillMatch.Wpf
   ```

## 打包说明

项目提供了 `publish.bat` 脚本用于快速打包：

1. 双击运行 `publish.bat`。
2. 发布后的文件将位于 `BillMatch.Wpf\bin\Release\net8.0-windows\win-x64\publish\` 目录下。
3. `BillMatch.Wpf.exe` 是一个独立的单文件可执行程序。

## 技术栈

- .NET 8 (WPF)
- CommunityToolkit.Mvvm (MVVM 框架)
- EPPlus (Excel 处理)
- Microsoft.Extensions.DependencyInjection (依赖注入)
