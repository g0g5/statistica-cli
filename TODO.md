# C# 重构实施进度

实施依据见 [REFACTOR-PLAN.md](REFACTOR-PLAN.md)。

## 已完成

- [x] 固定 .NET SDK 11.0.100，发布 `win-x64` Native AOT 产物 `statistica-cli.exe`。AOT 和裁剪分析无警告。
- [x] 使用同一 EXE 运行父进程与独立 STA 工作进程，经命名管道传输请求、诊断和完成记录。超时终止工作进程，并核验 Application 服务器身份后清理。
- [x] 迁移 `probe`、`inspect` 和 `export-csv`。两个后端均读到 STATISTICA 14.0。验证了注册诊断、`host.pwsh`、进程归属、正常释放、SAFEARRAY 双向传输、非零下界、可选及索引参数、HRESULT 和 EXCEPINFO。
- [x] 迁移 `create-sta`、`edit-sta` 和 `import-csv`。验证了缺失值、列和单元格修改、公式执行顺序、保存后回读。CSV 按块写入并流式核对，测试包含 16,512 条记录及跨分块的空行。
- [x] 迁移 `run-macro` 的活动数据表和 `GetScriptArgument` 模式。验证了真实宏、预期单元格和超时后的服务器清理。全局工具通过 JSON 文件传入预期单元格，独立 EXE 也支持内联 JSON。
- [x] 迁移 `batch`。保留清单字段、相对路径、前序输出依赖、逐作业状态、继续执行和部分成功语义。验证了隐藏会话复用，以及后续宏超时后提交前序产物。
- [x] 迁移 `create-graph`、`export-graph`、`list-workbook`、`extract-workbook-item` 和 `build-workbook`。验证了图表原生文件回读、散点连线、PNG/SVG、工作簿条目顺序与重名路径。
- [x] 迁移 `screenshot`，使用 Win32 采集和 WIC 编码。通过未锁定桌面会话验证 STA、STG、STW 和文档裁切。
- [x] 对照旧入口检查合成数据及旧式 ANSI STA 的 CSV 字节。发布版通过基线和集成冒烟测试，覆盖并发覆盖、失败回滚、缺失值、旧式 ANSI 显示文本和原始输入完整性。
- [x] 将通用化的 [技能](skills/statistica-cli/SKILL.md) 与 [CLI 指南](skills/statistica-cli/references/cli-guide.md) 嵌入 EXE。验证独立版与全局工具的离线安装、工作目录与用户目录、相对链接、重复安装、冲突和强制覆盖。
- [x] 切换到 C# 入口，移除旧 PowerShell 运行脚本和内部样式宏，更新项目说明。在新的 GitHub 克隆中应用本地变更后构建入口包和 `win-x64` 包，从本地源安装当前用户的全局工具，并在仓库外通过 PATH 验证 `statistica-cli probe --detailed`。

## 待确认

- 本次变更进入 GitHub 后，可直接从远端版本复验源码克隆与打包。
- Library 曾间歇返回 `0x8000FFFF`。后续独立探测、两种入口的数据表对照和本次完整回归均通过，历史故障原因待查。
