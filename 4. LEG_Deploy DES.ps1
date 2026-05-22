[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExcelPath = '.\서버정보\20260422_리소스배포_종합.xlsx',

    [Parameter()]
    [switch]$ConnectAccount,

    [Parameter()]
    [switch]$DryRun
)

throw "현재 운영 정책에서 DES 배포는 제외되었습니다. '99. Deploy Resources.ps1'의 DeployType에서 DES는 지원하지 않습니다."
