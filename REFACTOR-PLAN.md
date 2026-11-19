# STATISTICA CLI 的 C# 重构方案

> 本文件保留重构方案原文。实施进度、验收结果和待办事项统一记录在 [TODO.md](TODO.md)，后续更新请写入该文件。

## 目标与边界

用 C# 重写 Windows CLI，以本机 .NET SDK 11.0.100 发布 `win-x64` Native AOT 产物。项目和命令统一使用 **statistica-cli**，发布文件名为 **`statistica-cli.exe`**，便于与 STATISTICA 产品及其 `statist.exe` 区分。运行端依赖本机已安装并授权的 STATISTICA 和 Windows 系统组件。首版在当前 x64 环境验收。

命令采用小写连字符子命令和双横线长选项。保留现有 14 项功能、受支持的文档类型、Batch 清单、数据表 JSON 规格和结果 JSON 格式。Batch 中已有的 `Action` 值与选项字段拼写继续有效。例如，命令行使用 `export-csv` 和 `--input-path`，清单继续接受 `ExportCsv` 和 `InputPath`。

| 功能 | 子命令 | 首版范围 |
| --- | --- | --- |
| 探测 | `probe` | 保留两种 COM 后端、Auto 选择和详细诊断 |
| 数据表与 CSV | `inspect`、`export-csv`、`edit-sta`、`create-sta`、`import-csv` | 保留 STA、UTF-8 CSV、行列选择、显示文本、缺失值、公式和保存后验证 |
| 宏 | `run-macro` | 保留活动数据表与 `GetScriptArgument` 两种 SVB 模式，以及预期单元格校验 |
| 图表与截图 | `create-graph`、`export-graph`、`screenshot` | 保留折线图、散点图、散点连线及 STG 到 PNG/SVG 的导出。截图仍支持 STA、STG、STW 和文档裁切 |
| 工作簿 | `list-workbook`、`extract-workbook-item`、`build-workbook` | 保留条目顺序、索引和路径选择、重名检测，以及 STA/STG 的提取与组装 |
| 批处理 | `batch` | 保留清单相对路径、前序输出依赖、逐作业状态、继续执行策略和部分成功语义 |
| 技能安装 | `install-skill` | 从 EXE 内嵌资源安装技能到当前目录或当前用户主目录 |

## 本机调研与技术选型

本机已确认安装 .NET SDK 11.0.100、Visual Studio 2026 C++ Build Tools 和 Windows SDK。临时探针使用无 `PackageReference` 的项目完成了 `win-x64` Native AOT 发布，并在显式 STA 初始化后，从 `Statistica.Application` 和 `Statistica.Library` 均读到了版本 14.0。注册信息与实测 PID 均表明 `Application` 是进程外服务器，`Library` 加载在调用进程内。探针在 `Library` 中用自建数据表读到了下界为 `(1,1)` 的二维 SAFEARRAY，并直接读写了 `Graph.Plot(1).DisplayLine`。二维数组写入、完整宏流程及截图仍待集成验证。

- **STATISTICA COM 无需额外包。** 用 `LibraryImport` 调用系统 COM/OLE Automation API，用窄封装执行 `IDispatch::GetIDsOfNames` 和 `Invoke`。互操作层集中处理参数逆序、属性赋值、可选参数、HRESULT 和 EXCEPINFO，并负责 VARIANT、BSTR、SAFEARRAY 和接口引用的所有权。DISPID 按接口身份缓存。业务层只接触强类型文档、会话和数据块。内置 COM 运行时互操作与 Native AOT 不兼容，SDK 的 COM 源生成器也无法直接投影这里的 Automation 接口。 [Microsoft COM 源生成文档](https://learn.microsoft.com/dotnet/standard/native-interop/comwrappers-source-generation) · [IDispatch 调用文档](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/accessing-members-through-idispatch)
- **Win32 截图无需额外包。** 通过 `LibraryImport` 调用 User32 和 GDI，用 `PrintWindow` 渲染至可复用的位图缓冲区，再用系统自带的 WIC 生成 PNG。按需声明所用 API。保留活动文档、窗口归属、未锁定桌面和连续两帧稳定且非空的检查。 [Win32 互操作指南](https://learn.microsoft.com/en-us/windows/apps/develop/interop/call-win32-apis) · [PrintWindow](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-printwindow) · [WIC 编码](https://learn.microsoft.com/en-us/windows/win32/wic/-wic-creating-encoder)
- **命令解析与序列化。** 建议用 `System.CommandLine` 自带的解析和帮助功能，显式绑定命令参数。这是命令行层的 NuGet 依赖。JSON 使用 SDK 自带的 `System.Text.Json` 源生成。诊断兼容测试要覆盖宿主从 PowerShell 切换到 Native AOT 的差异，原 `host.pwsh` 字段保留并设为 `null`。 [System.CommandLine 官方文档](https://learn.microsoft.com/en-us/dotnet/standard/commandline/) · [System.Text.Json 源生成文档](https://learn.microsoft.com/dotnet/standard/serialization/system-text-json/source-generation)

## 运行架构

**同一 EXE 承担 CLI 和内部工作进程两个角色。** 父进程负责解析、校验、调度、超时和结果提交。子进程在显式初始化的 STA 线程上执行 COM 调用和截图，通过管道发送作业进度、诊断及已完成结果。STA 线程负责处理 Windows 消息；COM 对象在所属线程上释放。同步 COM 调用和 `PrintWindow` 可能阻塞，因此保留可由父进程终止的进程边界。 [Microsoft STA 文档](https://learn.microsoft.com/en-us/windows/win32/com/single-threaded-apartments)

- **会话。** 保留 Auto 的后端优先顺序和激活阶段的回退条件。兼容的隐藏作业复用会话，截图和活动数据表宏各用独立的 `Application` 实例。按作业计时，默认 120 秒。超时可终止工作进程，仅在 PID、创建时间、会话和启动身份核验通过后才清理进程外服务器。激活阶段身份未确认时记录诊断，已开始执行的宏和写入作业避免自动重放。
- **结果提交。** 输出先写入目标目录内的唯一临时路径，关闭文件并执行命令对应的回读或格式验证后，再由父进程移动到目标路径。保留 `--force` 覆盖规则。Batch 按清单顺序处理前序输出依赖，完成并验证的产物可在后续作业超时后提交。保留 `succeeded`、`failed`、`skipped`、`timed_out` 及失败时的非零退出状态。
- **数据路径。** 用强类型请求和结果替换跨脚本传递的哈希表，CLI 与 Batch 共用业务校验。CSV 导入分块解析和写入，将规范化后的期望数据存入临时文件以流式回读比对，避免同时持有全部行和全部数据块。CSV 导出直接写 UTF-8，缓存列元数据并合并连续数值列的 COM 调用。旧式 ANSI 数据表的文本和显示文本保留逐单元格读取路径。用非方形数组、非零下界和混合类型样本校验 COM 数组的行列映射。
- **图表、工作簿与截图。** 图表连线使用已验证的 `DisplayLine` 属性，替换内部样式宏。按索引提取工作簿条目时遍历到目标即可，按路径提取时遍历以识别重名。截图直接采样位图缓冲区，省去逐点 `GetPixel` 与字符串拼接。

## 去除课程依赖

用项目自带的合成 CSV、JSON 和通用 SVB 替换 `CourseRoot`、`lab4/6/7`、课程宏和固定实验结果。测试中生成 STA、STG 和 STW；凡当前创建路径无法生成的兼容样本，例如旧式文本编码的数据表，使用来源明确的最小通用样本。同步更新 `SKILL.md`、CLI 指南和项目说明中的命令、路径及示例。

迁移时逐项核对数字常量和假设。图表分析 ID、COM 枚举、1 基行列索引和缺失值读取规则来自 STATISTICA 接口，集中定义并使用具名常量或类型。图表样式、文本编码和等值校验由通用样本覆盖，避免把课程数据的列号、行数或预期计算结果写入实现。

## 源码分发与全局安装

源码托管在 GitHub。用户克隆后，在 Windows x64 上使用 .NET SDK 11.0.100、MSVC 和 Windows SDK 构建 Native AOT 工具包，再从本地包目录安装为当前用户的 .NET 全局工具。

C# 项目设置 `PackAsTool=true`、`PublishAot=true`、`RuntimeIdentifiers=win-x64`、`ToolPackageRuntimeIdentifiers=win-x64`，并将 `PackageId`、`ToolCommandName` 和 `AssemblyName` 均设为 `statistica-cli`。入口包和 `win-x64` 包使用同一版本并放在同一个本地包目录。本机临时探针使用此打包配置完成了全局安装，并在仓库外通过 PATH 启动了实际的 Native AOT EXE。 [Microsoft Native AOT 工具打包文档](https://learn.microsoft.com/dotnet/core/tools/rid-specific-tools)

将下面 URL 中的 `OWNER` 替换为 GitHub 仓库所有者。从克隆后的仓库根目录打包并安装。

```powershell
git clone https://github.com/OWNER/statistica-cli.git
Set-Location statistica-cli
dotnet pack ./src/statistica-cli/statistica-cli.csproj -c Release -o ./artifacts/tool-packages
dotnet pack ./src/statistica-cli/statistica-cli.csproj -c Release -r win-x64 -o ./artifacts/tool-packages
dotnet tool install --global statistica-cli --source ./artifacts/tool-packages
Get-Command statistica-cli
statistica-cli probe --detailed
```

全局工具入口默认位于 `%USERPROFILE%\.dotnet\tools`。若当前终端无法解析命令名，打开新终端并确认该目录已加入用户 PATH。STATISTICA 业务命令仍需本机已安装并授权的 STATISTICA。 [Microsoft .NET 全局工具文档](https://learn.microsoft.com/dotnet/core/tools/global-tools)

## 技能分发

将现有 `SKILL.md` 移至 `skills/statistica-cli/SKILL.md`，将现有 CLI 指南移至 `skills/statistica-cli/references/cli-guide.md`，保留技能文件指向指南的相对链接。C# 项目将两份文件声明为 `EmbeddedResource`。独立发布的 EXE 和全局工具包中的 EXE 均可离线安装随构建版本提供的技能。

| 命令 | 安装目录 |
| --- | --- |
| `statistica-cli install-skill` | 当前工作目录下的 `.agents/skills/statistica-cli/` |
| `statistica-cli install-skill --global` | 当前用户主目录下的 `~/.agents/skills/statistica-cli/` |

命令直接创建目标目录并写入 `SKILL.md` 和 `references/cli-guide.md`，无需启动 COM 工作进程。重复安装相同内容时成功退出。写入前检查两个目标文件，任一文件已有不同内容时报告冲突并以非零状态退出。显式传入 `--force` 可覆盖这两份文件。

## 实施顺序与验收

1. **建立 AOT 基线。** 固定本机 SDK 11.0.100，把探针扩展为 `probe --detailed`。验证两个后端、SAFEARRAY 双向传输、带参数的属性、错误信息及引用释放。确认进程隔离和超时清理后进入功能迁移。
2. **迁移命令。** 依次实现数据表与 CSV、SVB 与 Batch、图表与工作簿、桌面截图。用同一套通用输入对照旧 CLI 与新 EXE 的退出状态、JSON、CSV 和原生文件回读结果。
3. **验证发布产物。** 执行 `dotnet publish -c Release -r win-x64`，以 `AssemblyName` 固定产物名 `statistica-cli.exe`。门槛为 AOT 和裁剪分析无警告，并直接运行发布的 EXE 完成集成测试。覆盖缺失值、旧 STA 文本、公式顺序、工作簿重名、Batch 依赖与部分成功、并发覆盖、宏超时清理及原始输入完整性。截图与活动数据表宏在未锁定的桌面会话验证。
4. **切换入口并验收安装。** 通过上述验收后，以 `statistica-cli.exe` 替换 PowerShell 入口，移除旧运行脚本和内部样式宏。从干净的 GitHub 克隆目录构建两份包，使用本地源全局安装，并在仓库外通过 PATH 运行 `statistica-cli probe --detailed`。分别使用全局工具入口和独立 EXE 验证技能安装，覆盖离线执行、当前工作目录与用户主目录、指南相对链接、重复安装、冲突和强制覆盖。

构建端使用 .NET SDK、MSVC 与 Windows SDK。运行端直接启动发布的可执行文件。

```powershell
.\statistica-cli.exe probe --detailed
.\statistica-cli.exe export-csv --input-path sample.sta --output-path sample.csv
.\statistica-cli.exe batch --manifest-path jobs.json --continue-on-error
```

[Microsoft Native AOT 发布文档](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/)
