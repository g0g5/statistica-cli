#Requires -Version 7.5
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateSet('Probe', 'Inspect', 'ExportCsv', 'RunMacro', 'Batch',
        'EditSta', 'CreateSta', 'ImportCsv', 'CreateGraph', 'ExportGraph', 'Screenshot',
        'ListWorkbook', 'ExtractWorkbookItem', 'BuildWorkbook')]
    [string]$Action,
    [string]$InputPath,
    [string]$MacroPath,
    [string]$OutputPath,
    [string]$Columns,
    [ValidateRange(1, 2147483647)] [int]$FirstRow = 1,
    [ValidateRange(0, 2147483647)] [int]$LastRow = 0,
    [switch]$DisplayText,
    [switch]$PassInputAsScriptArgument,
    [ValidateSet('Auto', 'Application', 'Library')] [string]$Backend = 'Auto',
    [ValidateRange(1, 86400)] [int]$TimeoutSeconds = 120,
    [string]$DiagnosticPath,
    [switch]$Detailed,
    [switch]$Force,
    [string]$ManifestPath,
    [switch]$ContinueOnError,
    [string]$SpecPath,
    [string]$MissingToken,
    [hashtable]$VerifyCells,
    [ValidateRange(0, [double]::MaxValue)] [double]$Tolerance = 1e-8,
    [ValidateSet('Line', 'Scatter')] [string]$GraphType,
    [int]$XColumn,
    [int[]]$YColumns,
    [switch]$ConnectPoints,
    [switch]$CropToDocument,
    [ValidateRange(1, 120)] [int]$CaptureTimeoutSeconds = 8,
    [int]$ItemIndex,
    [string]$ItemPath,
    [string[]]$DocumentPaths
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'statistica-cli requires pwsh on Windows.' }

. (Join-Path $PSScriptRoot 'lib/Request.ps1')
. (Join-Path $PSScriptRoot 'lib/WorkerProcess.ps1')

$settings = @{
    Action = $Action
    Backend = $Backend
    TimeoutSeconds = $TimeoutSeconds
    DiagnosticPath = $DiagnosticPath
    ContinueOnError = [bool]$ContinueOnError
    Verbose = $VerbosePreference -eq 'Continue'
}
$request = New-StaRequest $settings (@{} + $PSBoundParameters)
Invoke-StaRequest $request (Join-Path $PSScriptRoot 'Worker.ps1')
