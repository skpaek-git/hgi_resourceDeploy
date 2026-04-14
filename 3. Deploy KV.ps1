[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExcelPath = '.\서버정보\20260320_샘플_리소스배포_정리.xlsx',

    [Parameter()]
    [switch]$ConnectAccount,

    [Parameter()]
    [switch]$DryRun
)

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '99. Deploy Resources.ps1'
& $scriptPath -ExcelPath $ExcelPath -DeployType @('KV') -ConnectAccount:$ConnectAccount -DryRun:$DryRun
