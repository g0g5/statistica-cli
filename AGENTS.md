## Overview

这是一个 Windows x64 上通过 TIBCO STATISTICA COM 自动操作数据表、SVB 宏、图表和工作簿的 C# Native AOT CLI。

## Tech Stack

- 运行环境：Windows x64、本机已授权的 TIBCO STATISTICA（Application 和 Library COM 后端）。
- 构建工具：`global.json` 固定 .NET SDK 11.0.100，发布 `win-x64` Native AOT 产物需要 MSVC 和 Windows SDK。
- 包管理器：NuGet，CLI 使用 `System.CommandLine` 解析参数。
- 测试：PowerShell 7.5 或更新版本，桌面测试需要未锁定的交互式会话。

## Structural Map

```text
statistica-cli/
|- README.md                 # 用户依赖、构建、CLI 和 Skill 安装
|- global.json               # 固定 .NET SDK 版本
|- docs/
|  |- REFACTOR-PLAN.md        # 原始重构方案
|  \- TODO.md                 # 重构进度、验收记录和待确认事项
|- src/
|  \- statistica-cli/
|     |- Program.cs          # 命令行和内部 STA 工作进程入口
|     |- JobCommands.cs      # 命令与批处理清单的共享校验
|     |- JobRunner.cs        # 提交已验证产物和生成批处理结果
|     |- WorkerRunner.cs     # 命名管道、作业超时和进程清理
|     |- StaWorker*.cs       # STA 线程上的会话管理和 COM 操作调度
|     \- Automation/        # IDispatch、VARIANT 和 SAFEARRAY 互操作
|- skills/
|  \- statistica-cli/        # 嵌入 EXE 的技能和 CLI 指南
\- tests/
   |- Baseline/              # COM 互操作与数据表回归测试
   |- fixtures/              # 随仓库分发的集成测试夹具
   \- smoke.ps1              # 发布版 CLI 的集成冒烟测试
```

## Commands

```powershell
# 在项目根目录构建 CLI，或发布用于集成测试的 Native AOT EXE。
dotnet build ./src/statistica-cli/statistica-cli.csproj
dotnet publish ./src/statistica-cli/statistica-cli.csproj -c Release -r win-x64

# 使用本机 STATISTICA 验证 COM、工作进程和数据表。
$exe = './src/statistica-cli/bin/Release/net11.0/win-x64/publish/statistica-cli.exe'
& $exe probe --detailed
dotnet run --project ./tests/Baseline -- $exe
./tests/Baseline/spreadsheets.ps1 -Executable $exe

# 使用随附夹具运行冒烟测试。在未锁定的桌面可加 -IncludeDesktop，覆盖活动数据表宏、截图和真实宏超时清理。
./tests/smoke.ps1 -Executable $exe
```

## Workflows

[No content yet.]

## Notes

[No content yet.]
