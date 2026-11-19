## Overview

这是一个在 Windows 上通过 TIBCO STATISTICA COM 自动处理 sm3ABP 课程数据的 CLI，正从 PowerShell 迁移至 C#。C# 入口已支持 `probe`、`inspect` 和 `export-csv`。PowerShell 入口仍覆盖数据表、CSV、SVB 宏、图表和工作簿。

## Tech Stack

- 语言：C#、PowerShell、STATISTICA Visual Basic（SVB）
- 运行环境：Windows x64、.NET 11（`global.json` 固定 SDK 11.0.100）、PowerShell 7.5 或更新版本
- 依赖：本机已授权的 TIBCO STATISTICA（`Application` 和 `Library` 两种 COM 后端）。发布 `win-x64` Native AOT 产物需要 MSVC 和 Windows SDK。
- 包管理器：NuGet，C# 项目用 `System.CommandLine` 解析命令行

## Structural Map

```text
statistica-cli/
|- global.json               # 固定 .NET SDK 版本
|- TODO.md                   # 记录迁移进度和剩余工作，链接重构方案
|- SKILL.md                  # PowerShell CLI 用法速览
|- src/
|  \- statistica-cli/        # C# CLI 和 Native AOT 工具包项目
|     |- Program.cs          # 命令定义和内部工作进程入口
|     |- WorkerRunner.cs     # 父进程通过命名管道调度请求和处理超时
|     |- StaWorker.cs        # 在 STA 线程上分发 COM 操作
|     \- Automation/        # IDispatch、VARIANT 和 SAFEARRAY 互操作
|- scripts/                  # PowerShell CLI 和 STATISTICA 自动化实现
|  |- statistica.ps1         # PowerShell 入口，定义命令和参数
|  |- Worker.ps1             # 在独立 STA 进程中调度并执行请求
|  |- lib/                   # 请求校验、进程与 COM 会话管理、文档操作
|  \- svb/                   # 随 CLI 提供的图表样式宏
|- references/
|  \- cli-guide.md           # PowerShell CLI 命令、参数和用法示例
\- tests/
   |- Baseline/              # C# COM 基线和两种入口的数据表结果对照
   |- fixtures/              # 随仓库分发的课程测试夹具
   \- smoke.ps1              # PowerShell CLI 集成冒烟测试
```

## Commands

```powershell
# 在项目根目录构建 C# CLI，或发布用于集成测试的 Native AOT 可执行文件。
dotnet build ./src/statistica-cli/statistica-cli.csproj
dotnet publish ./src/statistica-cli/statistica-cli.csproj -c Release -r win-x64

# 使用本机 STATISTICA 验证 COM 后端和工作进程，并对照数据表检查和 CSV 导出结果。
$exe = './src/statistica-cli/bin/Release/net11.0/win-x64/publish/statistica-cli.exe'
& $exe probe --detailed
dotnet run --project ./tests/Baseline -- $exe
./tests/Baseline/spreadsheets.ps1 -Executable $exe

# 检查 PowerShell CLI 的 COM 后端和工作进程。
./scripts/statistica.ps1 Probe -Detailed

# 在项目根目录使用随附夹具运行冒烟测试。
# 在已解锁的桌面会话中可加 -IncludeDesktop，覆盖活动数据表宏、截图和超时清理。
./tests/smoke.ps1
```

## Workflows

[No content yet.]

## Notes

[No content yet.]
