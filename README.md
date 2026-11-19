# statistica-cli

[简体中文](README_CN.md) · [Русский](README_RU.md)

Work with TIBCO STATISTICA spreadsheets, SVB macros, graphs, and workbooks from the command line on Windows.

## Requirements

- Running the CLI requires Windows x64 and a locally installed, licensed copy of TIBCO STATISTICA. Taking screenshots and running macros that depend on an active spreadsheet require an unlocked interactive desktop session.
- Building from source requires .NET SDK 11.0.100 (pinned in `global.json`), along with the MSVC C++ toolchain and Windows SDK from Visual Studio Build Tools. Use Git to clone the source code.

## Build and install

In PowerShell, clone the repository and change to the project root.

```powershell
git clone https://github.com/g0g5/statistica-cli.git
Set-Location statistica-cli
```

The packaging command creates the entry-point package and the `win-x64` package in the same directory. Then install the CLI as a .NET global tool for the current user.

```powershell
dotnet pack ./src/statistica-cli/statistica-cli.csproj -c Release -o ./artifacts/tool-packages
dotnet tool install --global statistica-cli --source ./artifacts/tool-packages
statistica-cli probe --detailed
```

If the terminal cannot find `statistica-cli`, make sure `%USERPROFILE%\.dotnet\tools` is on the user's `PATH`, then open a new terminal.

You can also publish a standalone Native AOT executable and run it directly. Running the executable requires neither the .NET SDK nor the .NET runtime.

```powershell
dotnet publish ./src/statistica-cli/statistica-cli.csproj -c Release -r win-x64
& './src/statistica-cli/bin/Release/net11.0/win-x64/publish/statistica-cli.exe' probe --detailed
```

## Install the Skill

The CLI includes the Skill and command guide for offline installation. Both the global tool and the standalone executable support the following commands. When using the standalone executable, replace `statistica-cli` with the path used to invoke it.

```powershell
# Run from the root of the project where you want to use the Skill.
statistica-cli install-skill

# Or install it in the current user's home directory.
statistica-cli install-skill --global
```

These options install the files under `.agents/skills/statistica-cli/` in the project root and the user's home directory, respectively. To overwrite conflicting files, add `--force`.

For more usage information, see the [CLI command guide](skills/statistica-cli/references/cli-guide.md).
