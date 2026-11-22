$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true

Push-Location -LiteralPath $PSScriptRoot
try {
    dotnet publish ./src/statistica-cli/statistica-cli.csproj -c Release -r win-x64 -o ./dist
}
finally {
    Pop-Location
}
