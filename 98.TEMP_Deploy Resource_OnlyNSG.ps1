[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExcelPath = '.\서버정보\20260422_리소스배포_개발.xlsx',

    [Parameter()]
    [string]$WorksheetName = 'NSG',

    [Parameter()]
    [Alias('Role')]
    [string[]]$Option = @(),

    [Parameter()]
    [switch]$ConnectAccount,

    [Parameter()]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Message)
    Write-Warning $Message
}

function Get-CellValue {
    param(
        [psobject]$Row,
        [string]$Field
    )
    if ($null -eq $Row -or [string]::IsNullOrWhiteSpace($Field)) { return $null }
    $prop = $Row.PSObject.Properties[$Field]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $null }
    $value = $prop.Value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    return $value
}

function Get-CellValueAny {
    param(
        [psobject]$Row,
        [string[]]$Fields
    )
    foreach ($field in $Fields) {
        $value = Get-CellValue -Row $Row -Field $field
        if ($value) { return $value }
    }
    return $null
}

function Ensure-Modules {
    Import-Module ImportExcel -ErrorAction Stop
    Import-Module Az.Accounts -ErrorAction Stop
    Import-Module Az.Network -ErrorAction Stop
}

function Ensure-AzSession {
    param(
        [switch]$ConnectAccount,
        [switch]$DryRun
    )

    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx -and $ConnectAccount) {
        Connect-AzAccount -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext -ErrorAction Stop
    }

    if (-not $ctx -and -not $DryRun) {
        throw 'Azure 세션이 없습니다. -ConnectAccount 옵션을 사용하거나 사전에 로그인해 주세요.'
    }

    return $ctx
}

function Get-FilteredRows {
    param(
        [psobject[]]$Rows,
        [string[]]$OptionFilters
    )

    if ($OptionFilters.Count -eq 0) { return @($Rows) }

    $filters = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $OptionFilters) {
        if (-not [string]::IsNullOrWhiteSpace($f)) {
            [void]$filters.Add($f.Trim())
        }
    }
    if ($filters.Count -eq 0) { return @($Rows) }

    return @(
        $Rows | Where-Object {
            $role = Get-CellValue -Row $_ -Field 'Role'
            $role -and $filters.Contains($role)
        }
    )
}

function Invoke-DeployNsgOnly {
    param(
        [string]$ExcelPath,
        [string]$WorksheetName,
        [string[]]$OptionFilters,
        [switch]$DryRun
    )

    if (-not (Test-Path -LiteralPath $ExcelPath)) {
        throw "Excel 파일을 찾을 수 없습니다: $ExcelPath"
    }

    Write-Info "Excel 시트 로드: $WorksheetName"
    $rows = @(Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName)
    if ($rows.Count -eq 0) {
        Write-Warn "시트 '$WorksheetName'에 데이터가 없습니다."
        return
    }

    $rows = Get-FilteredRows -Rows $rows -OptionFilters $OptionFilters
    if ($rows.Count -eq 0) {
        Write-Warn "필터 조건에 맞는 행이 없습니다. Option=$($OptionFilters -join ',')"
        return
    }

    $groups = $rows | Where-Object { Get-CellValue -Row $_ -Field 'NSGName' } | Group-Object -Property NSGName
    foreach ($group in $groups) {
        $base = $group.Group | Select-Object -First 1
        $nsgName = Get-CellValue -Row $base -Field 'NSGName'
        $rgName = Get-CellValueAny -Row $base -Fields @('RG', 'RGname')
        $location = Get-CellValue -Row $base -Field 'Location'
        if (-not $location) { $location = 'koreacentral' }

        if (-not $nsgName) { continue }
        if (-not $rgName) {
            Write-Warn "RG 값이 없어 건너뜁니다. NSG=$nsgName"
            continue
        }

        if ($DryRun) {
            Write-Info "[DryRun] NSG 생성/유지 대상: $nsgName (RG=$rgName, Location=$location)"
            Write-Info "[DryRun] NIC/서브넷 연결은 수행하지 않습니다."
            continue
        }

        $existing = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $rgName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Info "NSG 유지(이미 존재): $nsgName (RG=$rgName)"
            continue
        }

        [void](New-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $rgName -Location $location -Force -ErrorAction Stop)
        Write-Info "NSG 생성 완료: $nsgName (RG=$rgName)"
    }
}

try {
    Ensure-Modules
    [void](Ensure-AzSession -ConnectAccount:$ConnectAccount -DryRun:$DryRun)
    Invoke-DeployNsgOnly -ExcelPath $ExcelPath -WorksheetName $WorksheetName -OptionFilters $Option -DryRun:$DryRun
    Write-Info "작업 완료: NSG 생성(연결 없음) 플로우 종료"
} catch {
    Write-Error "치명적 오류: $($_.Exception.Message)"
    throw
}
