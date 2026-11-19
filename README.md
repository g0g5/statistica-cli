# statistica-cli

在 Windows 上通过命令行操作 TIBCO STATISTICA 数据表、SVB 宏、图表和工作簿。

## 依赖

- 运行需要 Windows x64 和本机已安装、授权的 TIBCO STATISTICA。截图和依赖活动数据表的宏需要未锁定的交互式桌面。
- 从源码构建需要 .NET SDK 11.0.100（由 `global.json` 固定），以及 Visual Studio Build Tools 中的 MSVC C++ 工具链和 Windows SDK。使用 Git 克隆源码。

## 构建与安装

在 PowerShell 中克隆仓库并进入项目根目录。

```powershell
git clone https://github.com/g0g5/statistica-cli.git
Set-Location statistica-cli
```

打包会在同一目录生成入口包和 `win-x64` 包。然后安装为当前用户的 .NET 全局工具。

```powershell
dotnet pack ./src/statistica-cli/statistica-cli.csproj -c Release -o ./artifacts/tool-packages
dotnet tool install --global statistica-cli --source ./artifacts/tool-packages
statistica-cli probe --detailed
```

如果终端找不到 `statistica-cli`，请确认 `%USERPROFILE%\.dotnet\tools` 已加入用户 `PATH`，然后打开新终端。

也可以发布并直接运行独立的 Native AOT 可执行文件。运行该文件无需安装 .NET SDK 或运行时。

```powershell
dotnet publish ./src/statistica-cli/statistica-cli.csproj -c Release -r win-x64
& './src/statistica-cli/bin/Release/net11.0/win-x64/publish/statistica-cli.exe' probe --detailed
```

## 安装 Skill

CLI 内置 Skill 和命令指南，可离线安装。全局工具和独立 EXE 都支持以下命令。使用独立 EXE 时，将 `statistica-cli` 换成该文件的调用路径。

```powershell
# 在需要使用 Skill 的项目根目录运行。
statistica-cli install-skill

# 或安装到当前用户的主目录。
statistica-cli install-skill --global
```

两种方式分别写入项目或用户主目录下的 `.agents/skills/statistica-cli/`。如需覆盖冲突文件，添加 `--force`。

更多用法见 [CLI 命令指南](skills/statistica-cli/references/cli-guide.md)。
