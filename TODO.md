# C# 重构实施进度

实施依据见 [REFACTOR-PLAN.md](REFACTOR-PLAN.md)。

## 已完成

- [x] 固定 .NET SDK 11.0.100，建立 C# 项目和 `win-x64` Native AOT 构建。发布产物为 `statistica-cli.exe`，AOT 和裁剪分析均无警告。
- [x] 建立同一 EXE 的父进程、管道和独立 STA 工作进程。实现超时终止及核验身份后的 Application 进程清理。
- [x] 实现 `probe` 和 `probe --detailed`，保留 Auto 后端选择、COM 注册诊断和 `host.pwsh` 字段。两个后端均读到 STATISTICA 14.0，并验证了进程归属和正常释放。
- [x] 用发布的 EXE 在两个后端验证非方形混合类型 SAFEARRAY 双向传输、非零下界、可选参数、索引属性、`Plot.DisplayLine`、HRESULT 和 EXCEPINFO。用内部测试作业验证超时清理。
- [x] 迁移 `inspect` 和 `export-csv`。用同一合成 STA 对照旧入口，验证 JSON 字段和值、常规与显示文本模式的 CSV 字节、列选择顺序与去重、行范围、缺失值、覆盖规则、失败回滚和源文件完整性。
- [x] 生成入口包和 `win-x64` 包，通过隔离目录安装，并在仓库外通过 `PATH` 启动。

## 待办

- [ ] 迁移 `create-sta`、`edit-sta` 和 `import-csv`，保留数据表 JSON 规格、缺失值、公式执行顺序和保存后回读。实现 CSV 分块导入和流式比对。
- [ ] 迁移 `run-macro` 的活动数据表和 `GetScriptArgument` 模式，验证预期单元格和真实宏超时清理。
- [ ] 迁移 `batch`，保留清单字段、相对路径、前序输出依赖、逐作业状态、继续执行和部分成功语义。验证隐藏会话复用以及后续作业超时后的前序产物提交。
- [ ] 迁移 `create-graph`、`export-graph`、`list-workbook`、`extract-workbook-item` 和 `build-workbook`，覆盖原生文件回读、图表样式、工作簿条目顺序和重名检测。
- [ ] 迁移 `screenshot`，实现 Win32 窗口采集与 WIC PNG 编码，在未锁定的桌面验证 STA、STG、STW 和文档裁切。
- [ ] 完成通用测试样本和发布产物验收，覆盖旧式 ANSI STA 文本、并发覆盖、工作簿重名、Batch 依赖与部分成功，以及桌面宏和截图。
- [ ] 移动并更新技能与 CLI 指南，将两份文件嵌入 EXE。实现 `install-skill` 并验证离线安装、工作目录和用户目录、相对链接、重复安装、冲突与强制覆盖。
- [ ] 验收通过后切换入口，移除旧 PowerShell 运行脚本和内部样式宏，更新项目说明。从干净的 GitHub 克隆构建并执行当前用户全局工具安装验收。

## 待确认

- Library 在集成复验中曾间歇返回 `0x8000FFFF`。随后 12 次独立探测和一次完整的数据表对照均通过，原因待查。
