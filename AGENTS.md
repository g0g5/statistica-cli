## Overview

这是一个在 Windows 上通过 PowerShell 和 TIBCO STATISTICA COM 自动处理 sm3ABP 课程数据的 CLI，支持原生数据表、CSV、SVB 宏、图表和工作簿。

## Tech Stack

- 语言：PowerShell、STATISTICA Visual Basic（SVB）
- 运行环境：Windows、PowerShell 7.5 或更新版本
- 依赖：本机已授权的 TIBCO STATISTICA，提供 `Application` 和 `Library` 两种 COM 后端
- 包管理器：无

## Structural Map

```text
statistical-cli/
|- SKILL.md                  # CLI 用法速览和常见操作入口
|- scripts/                  # CLI 和 STATISTICA 自动化实现
|  |- statistica.ps1         # CLI 入口，定义命令和参数
|  |- Worker.ps1             # 在独立 STA 进程中调度并执行请求
|  |- lib/                   # 请求校验、进程与 COM 会话管理、文档操作
|  \- svb/                   # 随 CLI 提供的图表样式宏
|- references/
|  \- cli-guide.md           # 命令、参数和用法示例
\- tests/
   \- smoke.ps1              # 使用 sm3ABP 课程样本验证 CLI
```

## Commands

```powershell
# 在项目根目录检查 STATISTICA COM 后端和工作进程。
./scripts/statistica.ps1 Probe -Detailed

# 将示例路径替换为 sm3ABP 课程仓库路径，在项目根目录运行冒烟测试。
# 在已解锁的桌面会话中可加 -IncludeDesktop，覆盖活动数据表宏、截图和超时清理。
./tests/smoke.ps1 -CourseRoot 'C:\path\to\sm3ABP'
```

## Workflows

[No content yet.]

## Notes

[No content yet.]
