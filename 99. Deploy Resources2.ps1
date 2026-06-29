[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExcelPath = '.\서버정보\20260422_리소스배포_종합.xlsx',

    [Parameter()]
    [ValidateSet('RG','VNET','UDR','LB','LB_PROBE','LB_RULE','NSG','VM','DATADISK')]
    [string[]]$DeployType = @('RG','VNET','UDR','LB','NSG','VM','DATADISK'),

    [Parameter()]
    [Alias('VmRole','Role')]
    [string[]]$Option = @(),

    [Parameter()]
    [switch]$ConnectAccount,

    [Parameter()]
    [switch]$DryRun
    ,
    [Parameter()]
    [switch]$LbResourceOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:IsExcelPathExplicit = $PSBoundParameters.ContainsKey('ExcelPath')

class DeploymentContext {
    [string]$ExcelPath
    [string[]]$DeployType
    [string[]]$VmRoleFilter
    [bool]$DryRun
    [string]$SubscriptionId
    [bool]$LbResourceOnly

    DeploymentContext([string]$excelPath, [string[]]$deployType, [string[]]$vmRoleFilter, [bool]$dryRun, [string]$subscriptionId, [bool]$lbResourceOnly) {
        $this.ExcelPath = $excelPath
        $this.DeployType = $deployType
        $this.VmRoleFilter = $vmRoleFilter
        $this.DryRun = $dryRun
        $this.SubscriptionId = $subscriptionId
        $this.LbResourceOnly = $lbResourceOnly
    }
}

class ValidationIssue {
    [string]$Type
    [int]$Row
    [string]$ResourceName
    [string]$Field
    [string]$Message

    ValidationIssue([string]$type, [int]$row, [string]$resourceName, [string]$field, [string]$message) {
        $this.Type = $type
        $this.Row = $row
        $this.ResourceName = $resourceName
        $this.Field = $field
        $this.Message = $message
    }
}

function Write-Info {
    param([string]$Message)
    if (Get-Command Out-Log -ErrorAction SilentlyContinue) { Out-Log $Message } else { Write-Host "[INFO] $Message" }
}

function Write-WarnLog {
    param([string]$Message)
    if (Get-Command Out-WarnLog -ErrorAction SilentlyContinue) { Out-WarnLog $Message } else { Write-Warning $Message }
}

function Write-ErrorLog {
    param([string]$Message)
    if (Get-Command Out-ErrLog -ErrorAction SilentlyContinue) { Out-ErrLog $Message } else { Write-Error $Message }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$Operation = 'Operation',
        [int]$MaxAttempts = 3,
        [int]$InitialDelaySeconds = 2
    )

    $attempt = 1
    while ($true) {
        try {
            return & $ScriptBlock
        } catch {
            $message = $_.Exception.Message
            $isTransient = (
                $message -match 'while sending the request' -or
                $message -match 'The request was aborted' -or
                $message -match 'A task was canceled' -or
                $message -match 'temporarily unavailable' -or
                $message -match 'timed out' -or
                $message -match 'timeout' -or
                $message -match 'TooManyRequests' -or
                $message -match '\(429\)'
            )

            if (-not $isTransient -or $attempt -ge $MaxAttempts) {
                throw
            }

            Write-WarnLog "$Operation 일시 오류(시도 $attempt/$MaxAttempts): $message"
            $delay = [int]([Math]::Pow(2, ($attempt - 1)) * $InitialDelaySeconds)
            Start-Sleep -Seconds $delay
            $attempt++
        }
    }
}

function Start-Step {
    param([string]$Name)
    Write-Info "[START] $Name"
}

function End-Step {
    param([string]$Name)
    Write-Info "[END] $Name"
}

function Get-CellValue {
    param(
        [psobject]$Row,
        [string]$Field
    )
    if ($null -eq $Row -or [string]::IsNullOrWhiteSpace($Field)) { return $null }
    $prop = $Row.PSObject.Properties[$Field]
    if ($null -eq $prop) { return $null }
    if ($null -eq $prop.Value) { return $null }
    $text = $prop.Value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return $text
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

function Convert-ToBoolean {
    param(
        [string]$Value,
        [bool]$Default = $false
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    $raw = $Value.Trim().ToLowerInvariant()
    if ($raw -in @('1','true','t','yes','y','o')) { return $true }
    if ($raw -in @('0','false','f','no','n','x')) { return $false }
    return $Default
}

function Test-IsEnabledRow {
    param(
        [psobject]$Row,
        [bool]$Default = $true
    )
    $value = Get-CellValueAny -Row $Row -Fields @('Enable','Enabled')
    return (Convert-ToBoolean -Value $value -Default $Default)
}

function Convert-ToNullableDateTime {
    param(
        [string]$Value,
        [string]$FieldName = 'DateTime'
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try {
        return [datetime]::Parse($Value)
    } catch {
        throw "$FieldName 값의 날짜 형식이 올바르지 않습니다: '$Value'"
    }
}

function Convert-OsDiskStorageType {
    param([string]$InputValue)
    switch ($InputValue) {
        'PremiumSSD' { return 'Premium_LRS' }
        'StandardSSD' { return 'StandardSSD_LRS' }
        'StandardHDD' { return 'Standard_LRS' }
        'PremiumSSD_ZRS' { return 'Premium_ZRS' }
        'StandardSSD_ZRS' { return 'StandardSSD_ZRS' }
        default {
            if ($InputValue -in @('Premium_LRS','StandardSSD_LRS','Standard_LRS','Premium_ZRS','StandardSSD_ZRS')) {
                return $InputValue
            }
            return 'Standard_LRS'
        }
    }
}

function Test-WindowsComputerName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $trimmed = $Name.Trim()
    if ($trimmed.Length -gt 15) { return $false }
    if ($trimmed -match '^\d+$') { return $false }
    if ($trimmed -notmatch '^[A-Za-z0-9-]+$') { return $false }
    return $true
}

function Get-WindowsComputerName {
    param([string]$VmName)

    if ([string]::IsNullOrWhiteSpace($VmName)) { return $VmName }
    $name = $VmName.Trim()
    if ($name.Length -gt 15) {
        return $name.Substring(0, 15)
    }
    return $name
}

function Get-VmSourceType {
    param([psobject]$Row)
    $imageResourceId = Get-CellValue -Row $Row -Field 'ImageResourceId'
    if ($imageResourceId) { return 'CustomImage' }
    return 'Marketplace'
}

function Get-WorksheetName {
    param(
        [string]$Path,
        [string[]]$Candidates
    )
    $sheetInfos = Get-ExcelSheetInfo -Path $Path
    foreach ($candidate in $Candidates) {
        $exact = $sheetInfos | Where-Object { $_.Name -eq $candidate } | Select-Object -First 1
        if ($exact) { return $exact.Name }
    }
    foreach ($candidate in $Candidates) {
        $wild = $sheetInfos | Where-Object { $_.Name -like "$candidate*" } | Select-Object -First 1
        if ($wild) { return $wild.Name }
    }
    return $null
}

function Get-SheetRows {
    param(
        [DeploymentContext]$Context,
        [string[]]$SheetCandidates,
        [switch]$Optional
    )

    $sheetName = Get-WorksheetName -Path $Context.ExcelPath -Candidates $SheetCandidates
    if (-not $sheetName) {
        if ($Optional) {
            Write-WarnLog "시트를 찾지 못했습니다. 후보: $($SheetCandidates -join ', ')"
            return @()
        }
        throw "필수 시트를 찾지 못했습니다. 후보: $($SheetCandidates -join ', ')"
    }

    Write-Info "Excel 시트 로드: $sheetName"
    $rows = Import-Excel -Path $Context.ExcelPath -WorksheetName $sheetName
    if ($null -eq $rows) { return @() }
    return @($rows)
}

function Add-Issue {
    param(
        [System.Collections.Generic.List[ValidationIssue]]$Issues,
        [string]$Type,
        [int]$Row,
        [string]$ResourceName,
        [string]$Field,
        [string]$Message
    )
    $Issues.Add([ValidationIssue]::new($Type, $Row, $ResourceName, $Field, $Message))
}

function Get-FilteredVmRows {
    param(
        [DeploymentContext]$Context,
        [psobject[]]$Rows
    )

    $allRows = @($Rows)
    if ($Context.VmRoleFilter.Count -eq 0) { return $allRows }

    $filters = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $Context.VmRoleFilter) {
        if (-not [string]::IsNullOrWhiteSpace($f)) {
            [void]$filters.Add($f.Trim())
        }
    }
    if ($filters.Count -eq 0) { return $allRows }

    $roleFieldExists = $false
    foreach ($r in $allRows) {
        if ($r.PSObject.Properties['Role']) {
            $roleFieldExists = $true
            break
        }
    }
    if (-not $roleFieldExists) {
        throw "VM 시트에 'Role' 컬럼이 없습니다. -Option 필터를 사용하려면 Role 컬럼이 필요합니다."
    }

    $filtered = @(
        $allRows | Where-Object {
            $role = Get-CellValue -Row $_ -Field 'Role'
            $role -and $filters.Contains($role)
        }
    )
    Write-Info "VM Role 필터 적용: $($filters -join ', ') / 대상 행 수: $($filtered.Count)"
    return $filtered
}

function Get-FilteredRowsByOption {
    param(
        [psobject[]]$Rows,
        [string[]]$OptionFilters,
        [string[]]$Columns = @('Option')
    )

    $allRows = @($Rows)
    if ($OptionFilters.Count -eq 0) { return $allRows }

    $filters = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $OptionFilters) {
        if (-not [string]::IsNullOrWhiteSpace($f)) {
            [void]$filters.Add($f.Trim())
        }
    }
    if ($filters.Count -eq 0) { return $allRows }

    $columnExists = $false
    foreach ($r in $allRows) {
        foreach ($col in $Columns) {
            if ($r.PSObject.Properties[$col]) {
                $columnExists = $true
                break
            }
        }
        if ($columnExists) { break }
    }
    if (-not $columnExists) {
        throw "시트에 필터 컬럼이 없습니다. -Option 필터를 사용하려면 다음 중 하나가 필요합니다: $($Columns -join ', ')"
    }

    $filtered = @(
        $allRows | Where-Object {
            $value = $null
            foreach ($col in $Columns) {
                $value = Get-CellValue -Row $_ -Field $col
                if ($value) { break }
            }
            $value -and $filters.Contains($value)
        }
    )

    Write-Info "Option 필터 적용: $($filters -join ', ') / 대상 행 수: $($filtered.Count)"
    return $filtered
}

function Convert-ToRouteNextHopType {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw 'NextHopType 값이 필요합니다.'
    }

    $raw = $Value.Trim().ToUpperInvariant()
    switch ($raw) {
        'VIRTUALAPPLIANCE' { return 'VirtualAppliance' }
        'INTERNET' { return 'Internet' }
        'VIRTUALNETWORKGATEWAY' { return 'VirtualNetworkGateway' }
        'VNETLOCAL' { return 'VnetLocal' }
        'NONE' { return 'None' }
        default { throw "지원하지 않는 NextHopType 입니다: $Value" }
    }
}

function Test-RequiredField {
    param(
        [System.Collections.Generic.List[ValidationIssue]]$Issues,
        [string]$Type,
        [int]$Row,
        [psobject]$Data,
        [string]$Field,
        [string]$ResourceName
    )
    $value = Get-CellValue -Row $Data -Field $Field
    if (-not $value) {
        Add-Issue -Issues $Issues -Type $Type -Row $Row -ResourceName $ResourceName -Field $Field -Message '필수 값이 비어 있습니다.'
    }
}

function Validate-Inputs {
    param([DeploymentContext]$Context)

    Start-Step 'Validate-Inputs'
    $issues = [System.Collections.Generic.List[ValidationIssue]]::new()

    if ($Context.DeployType -contains 'RG') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('RG')
        $i = 1
        foreach ($r in $rows) {
            $rgName = Get-CellValue -Row $r -Field 'RGname'
            if (-not $rgName) { $i++; continue }
            Test-RequiredField -Issues $issues -Type 'RG' -Row $i -Data $r -Field 'Location' -ResourceName $rgName
            $i++
        }
    }

    if ($Context.DeployType -contains 'VNET') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('VNET')
        $i = 1
        foreach ($r in $rows) {
            $vnetName = Get-VnetName -Row $r
            if (-not $vnetName) { $i++; continue }
            Test-RequiredField -Issues $issues -Type 'VNET' -Row $i -Data $r -Field 'RGname' -ResourceName $vnetName
            Test-RequiredField -Issues $issues -Type 'VNET' -Row $i -Data $r -Field 'Location' -ResourceName $vnetName
            $vnetAddresses = @(Get-VnetAddresses -Row $r)
            if ($vnetAddresses.Count -eq 0) {
                Add-Issue -Issues $issues -Type 'VNET' -Row $i -ResourceName $vnetName -Field 'VNet* Address' -Message '필수 값이 비어 있습니다.'
            }
            $i++
        }
    }

    if ($Context.DeployType -contains 'UDR') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('UDR','RouteTable','UDR_Route')
        $rows = Get-FilteredRowsByOption -Rows $rows -OptionFilters $Context.VmRoleFilter -Columns @('Option','UDRType','Direction','Role')
        $i = 1
        foreach ($r in $rows) {
            if (-not (Test-IsEnabledRow -Row $r -Default $true)) { $i++; continue }

            $udrName = Get-CellValueAny -Row $r -Fields @('UDRName','RouteTableName')
            $routeName = Get-CellValue -Row $r -Field 'RouteName'

            if (-not $udrName -and -not $routeName) { $i++; continue }

            if (-not $udrName) {
                Add-Issue -Issues $issues -Type 'UDR' -Row $i -ResourceName '' -Field 'UDRName' -Message '필수 값이 비어 있습니다.'
            }
            if (-not (Get-CellValueAny -Row $r -Fields @('RGname','RGName'))) {
                Add-Issue -Issues $issues -Type 'UDR' -Row $i -ResourceName $udrName -Field 'RGname' -Message '필수 값이 비어 있습니다.'
            }
            Test-RequiredField -Issues $issues -Type 'UDR' -Row $i -Data $r -Field 'Location' -ResourceName $udrName
            Test-RequiredField -Issues $issues -Type 'UDR' -Row $i -Data $r -Field 'RouteName' -ResourceName $udrName
            Test-RequiredField -Issues $issues -Type 'UDR' -Row $i -Data $r -Field 'AddressPrefix' -ResourceName $udrName
            Test-RequiredField -Issues $issues -Type 'UDR' -Row $i -Data $r -Field 'NextHopType' -ResourceName $udrName

            $nextHopType = Get-CellValue -Row $r -Field 'NextHopType'
            if ($nextHopType -and $nextHopType.Trim().ToUpperInvariant() -eq 'VIRTUALAPPLIANCE') {
                if (-not (Get-CellValue -Row $r -Field 'NextHopIpAddress')) {
                    Add-Issue -Issues $issues -Type 'UDR' -Row $i -ResourceName $udrName -Field 'NextHopIpAddress' -Message 'NextHopType=VirtualAppliance 인 경우 값이 필요합니다.'
                }
            }

            $associateSubnet = Convert-ToBoolean -Value (Get-CellValue -Row $r -Field 'AssociateToSubnet') -Default $false
            if ($associateSubnet) {
                if (-not (Get-CellValueAny -Row $r -Fields @('VnetName','VirtualNetworkName','VNetName'))) {
                    Add-Issue -Issues $issues -Type 'UDR' -Row $i -ResourceName $udrName -Field 'VnetName' -Message 'AssociateToSubnet=TRUE 인 경우 값이 필요합니다.'
                }
                if (-not (Get-CellValueAny -Row $r -Fields @('SubnetName','Subnet'))) {
                    Add-Issue -Issues $issues -Type 'UDR' -Row $i -ResourceName $udrName -Field 'SubnetName' -Message 'AssociateToSubnet=TRUE 인 경우 값이 필요합니다.'
                }
            }
            $i++
        }
    }

    if ($Context.DeployType -contains 'STORAGE') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('Storage')
        $i = 1
        foreach ($r in $rows) {
            $name = Get-CellValueAny -Row $r -Fields @('Storage name','StorageName','StorgeName')
            if (-not $name) { $i++; continue }
            Test-RequiredField -Issues $issues -Type 'STORAGE' -Row $i -Data $r -Field 'RGname' -ResourceName $name
            Test-RequiredField -Issues $issues -Type 'STORAGE' -Row $i -Data $r -Field 'Location' -ResourceName $name
            $i++
        }
    }

    if ($Context.DeployType -contains 'KV') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('KV','KeyVault')
        $i = 1
        foreach ($r in $rows) {
            if (-not (Test-IsEnabledRow -Row $r -Default $true)) { $i++; continue }
            $name = Get-CellValueAny -Row $r -Fields @('KVName','KeyVaultName')
            if (-not $name) { $i++; continue }
            Test-RequiredField -Issues $issues -Type 'KV' -Row $i -Data $r -Field 'RGname' -ResourceName $name
            Test-RequiredField -Issues $issues -Type 'KV' -Row $i -Data $r -Field 'Location' -ResourceName $name

            $keyName = Get-CellValueAny -Row $r -Fields @('KeyName','CMKKeyName')
            $secretName = Get-CellValue -Row $r -Field 'SecretName'
            $secretValue = Get-CellValue -Row $r -Field 'SecretValue'

            if (-not $keyName -and -not $secretName) {
                Add-Issue -Issues $issues -Type 'KV' -Row $i -ResourceName $name -Field 'KeyName/SecretName' -Message 'KeyName 또는 SecretName 중 하나는 필요합니다.'
            }

            if ($secretName -and -not $secretValue) {
                Add-Issue -Issues $issues -Type 'KV' -Row $i -ResourceName $name -Field 'SecretValue' -Message 'SecretName이 있으면 SecretValue가 필요합니다.'
            }
            if ($secretValue -and -not $secretName) {
                Add-Issue -Issues $issues -Type 'KV' -Row $i -ResourceName $name -Field 'SecretName' -Message 'SecretValue가 있으면 SecretName이 필요합니다.'
            }
            $i++
        }
    }

    if ($Context.DeployType -contains 'DES') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('DES','DES_PRD')
        $i = 1
        foreach ($r in $rows) {
            if (-not (Test-IsEnabledRow -Row $r -Default $true)) { $i++; continue }
            $name = Get-CellValueAny -Row $r -Fields @('DESName','DiskEncryptionSetName')
            if (-not $name) { $i++; continue }
            Test-RequiredField -Issues $issues -Type 'DES' -Row $i -Data $r -Field 'RGname' -ResourceName $name
            Test-RequiredField -Issues $issues -Type 'DES' -Row $i -Data $r -Field 'Location' -ResourceName $name
            if (-not (Get-CellValueAny -Row $r -Fields @('KeyVaultName','KVName'))) {
                Add-Issue -Issues $issues -Type 'DES' -Row $i -ResourceName $name -Field 'KeyVaultName' -Message '필수 값이 비어 있습니다.'
            }
            if (-not (Get-CellValueAny -Row $r -Fields @('KeyName','CMKKeyName'))) {
                Add-Issue -Issues $issues -Type 'DES' -Row $i -ResourceName $name -Field 'KeyName' -Message '필수 값이 비어 있습니다.'
            }
            $i++
        }
    }

    if ($Context.DeployType -contains 'NSG') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('NSG')
        $ruleRowsForNsg = Get-SheetRows -Context $Context -SheetCandidates @('NSG_Detail','NSG_Rule','NSG_PRD_Rule') -Optional
        $rows = Get-FilteredRowsByOption -Rows $rows -OptionFilters $Context.VmRoleFilter -Columns @('Role','Option')

        $selectedNsgNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($nr in $rows) {
            $nn = Get-CellValue -Row $nr -Field 'NSGName'
            if ($nn) { [void]$selectedNsgNames.Add($nn) }
        }
        if ($selectedNsgNames.Count -gt 0) {
            $ruleRowsForNsg = @($ruleRowsForNsg | Where-Object {
                $target = Get-CellValue -Row $_ -Field 'NSGName'
                $target -and $selectedNsgNames.Contains($target)
            })
        }
        $nsgNameSet = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        $i = 1
        foreach ($r in $rows) {
            $name = Get-CellValue -Row $r -Field 'NSGName'
            if (-not $name) { $i++; continue }
            [void]$nsgNameSet.Add($name)
            Test-RequiredField -Issues $issues -Type 'NSG' -Row $i -Data $r -Field 'RG' -ResourceName $name

            $useNic = Get-CellValue -Row $r -Field 'UseNIC'
            $nicName = Get-CellValueAny -Row $r -Fields @('NIC','NicName')
            $vnetName = Get-CellValueAny -Row $r -Fields @('VirtualNetwork','VNetName','VnetName')
            $subnetName = Get-CellValueAny -Row $r -Fields @('Subnet','SubnetName')
            $vnetRg = Get-CellValueAny -Row $r -Fields @('VnetRG','VNetRG','VirtualNetworkRG')

            if ($useNic -and $useNic.ToUpperInvariant() -eq 'O') {
                if (-not $nicName) {
                    Add-Issue -Issues $issues -Type 'NSG' -Row $i -ResourceName $name -Field 'NIC' -Message 'UseNIC=O 인 경우 NIC 값이 필요합니다.'
                }
            } else {
                if (-not $vnetName) {
                    Add-Issue -Issues $issues -Type 'NSG' -Row $i -ResourceName $name -Field 'VNetName/VirtualNetwork' -Message '서브넷 연결용 VNet 이름이 필요합니다.'
                }
                if (-not $subnetName) {
                    Add-Issue -Issues $issues -Type 'NSG' -Row $i -ResourceName $name -Field 'SubnetName/Subnet' -Message '서브넷 연결용 Subnet 이름이 필요합니다.'
                }
                if (-not $vnetRg) {
                    Add-Issue -Issues $issues -Type 'NSG' -Row $i -ResourceName $name -Field 'VNetRG' -Message '서브넷 연결용 VNetRG 값이 필요합니다.'
                }
            }

            $i++
        }

        $j = 1
        foreach ($rr in $ruleRowsForNsg) {
            $detailNsgName = Get-CellValue -Row $rr -Field 'NSGName'
            if (-not $detailNsgName) { $j++; continue }
            if (-not $nsgNameSet.Contains($detailNsgName)) {
                Add-Issue -Issues $issues -Type 'NSG_Detail' -Row $j -ResourceName $detailNsgName -Field 'NSGName' -Message 'NSG 시트에 존재하지 않는 NSGName 입니다.'
            }
            $j++
        }
    }

    if (@($Context.DeployType | Where-Object { $_ -in @('LB','LB_PROBE','LB_RULE') }).Count -gt 0) {
        $lbRows = Get-SheetRows -Context $Context -SheetCandidates @('LB','LB_PRD','Load Balancer')
        $probeRows = Get-SheetRows -Context $Context -SheetCandidates @('LB_Probe','LB_PRD_Probe','Load Balancer_Probe') -Optional
        $ruleRows = Get-SheetRows -Context $Context -SheetCandidates @('LB_Rule','LB_PRD_Rule','Load Balancer_Rule') -Optional
        $lbRows = Get-FilteredRowsByOption -Rows $lbRows -OptionFilters $Context.VmRoleFilter -Columns @('Role','Option')
        $probeRows = Get-FilteredRowsByOption -Rows $probeRows -OptionFilters $Context.VmRoleFilter -Columns @('Role','Option')
        $ruleRows = Get-FilteredRowsByOption -Rows $ruleRows -OptionFilters $Context.VmRoleFilter -Columns @('Role','Option')

        $selectedLbNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($lr in $lbRows) {
            $ln = Get-CellValueAny -Row $lr -Fields @('LBName','LoadBalancerName')
            if ($ln) { [void]$selectedLbNames.Add($ln) }
        }
        if ($selectedLbNames.Count -gt 0) {
            $probeRows = @($probeRows | Where-Object {
                $target = Get-CellValue -Row $_ -Field 'LBName'
                $target -and $selectedLbNames.Contains($target)
            })
            $ruleRows = @($ruleRows | Where-Object {
                $target = Get-CellValue -Row $_ -Field 'LBName'
                $target -and $selectedLbNames.Contains($target)
            })
        }

        $lbNameSet = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        $i = 1
        foreach ($r in $lbRows) {
            $lbName = Get-CellValueAny -Row $r -Fields @('LBName','LoadBalancerName')
            if (-not $lbName) { $i++; continue }

            [void]$lbNameSet.Add($lbName)
            if (-not (Get-CellValueAny -Row $r -Fields @('RGName','RGname'))) {
                Add-Issue -Issues $issues -Type 'LB' -Row $i -ResourceName $lbName -Field 'RGName' -Message '필수 값이 비어 있습니다.'
            }
            Test-RequiredField -Issues $issues -Type 'LB' -Row $i -Data $r -Field 'Location' -ResourceName $lbName
            if (-not (Get-CellValueAny -Row $r -Fields @('FEName','FrontendName'))) {
                Add-Issue -Issues $issues -Type 'LB' -Row $i -ResourceName $lbName -Field 'FEName' -Message '필수 값이 비어 있습니다.'
            }
            if (-not (Get-CellValueAny -Row $r -Fields @('BEPoolName','BackendPoolName'))) {
                Add-Issue -Issues $issues -Type 'LB' -Row $i -ResourceName $lbName -Field 'BEPoolName' -Message '필수 값이 비어 있습니다.'
            }

            $feType = Get-CellValueAny -Row $r -Fields @('FEType','FrontendType')
            if (-not $feType) { $feType = 'Internal' }
            if ($feType.ToUpperInvariant() -eq 'INTERNAL') {
                if (-not (Get-CellValueAny -Row $r -Fields @('FEVNetName','FrontendVnetName','VnetName'))) {
                    Add-Issue -Issues $issues -Type 'LB' -Row $i -ResourceName $lbName -Field 'FEVNetName' -Message 'Internal Frontend 사용 시 VNet 이름이 필요합니다.'
                }
                if (-not (Get-CellValueAny -Row $r -Fields @('FESubnetName','FrontendSubnetName','SubnetName'))) {
                    Add-Issue -Issues $issues -Type 'LB' -Row $i -ResourceName $lbName -Field 'FESubnetName' -Message 'Internal Frontend 사용 시 Subnet 이름이 필요합니다.'
                }
            }

            $allocation = Get-CellValue -Row $r -Field 'PrivateIPAllocation'
            if ($allocation -and $allocation.ToUpperInvariant() -eq 'STATIC' -and -not (Get-CellValue -Row $r -Field 'PrivateIPAddress')) {
                Add-Issue -Issues $issues -Type 'LB' -Row $i -ResourceName $lbName -Field 'PrivateIPAddress' -Message 'PrivateIPAllocation=Static 인 경우 PrivateIPAddress가 필요합니다.'
            }
            $i++
        }

        $validateProbes = (($Context.DeployType -contains 'LB') -and -not $Context.LbResourceOnly) -or ($Context.DeployType -contains 'LB_PROBE')
        $validateRules = (($Context.DeployType -contains 'LB') -and -not $Context.LbResourceOnly) -or ($Context.DeployType -contains 'LB_RULE')

        if ($validateProbes) {
            $j = 1
            foreach ($pr in $probeRows) {
                $targetLb = Get-CellValue -Row $pr -Field 'LBName'
                if (-not $targetLb) { $j++; continue }
                $probeName = Get-CellValue -Row $pr -Field 'ProbeName'
                $probePort = Get-CellValue -Row $pr -Field 'Port'

                # ProbeName/Port가 비어 있으면 해당 행은 배포 대상에서 제외(검증 오류 아님)
                if (-not $probeName -or -not $probePort) { $j++; continue }

                if (-not $lbNameSet.Contains($targetLb)) {
                    Add-Issue -Issues $issues -Type 'LB_Probe' -Row $j -ResourceName $targetLb -Field 'LBName' -Message 'LB 시트에 존재하지 않는 LBName 입니다.'
                }
                $j++
            }

        }

        if ($validateRules) {
            $j = 1
            foreach ($rr in $ruleRows) {
                $targetLb = Get-CellValue -Row $rr -Field 'LBName'
                if (-not $targetLb) { $j++; continue }
                $ruleName = Get-CellValue -Row $rr -Field 'RuleName'
                $enableHaPorts = Convert-ToBoolean -Value (Get-CellValue -Row $rr -Field 'EnableHAports') -Default $false
                $fePortText = Get-CellValueAny -Row $rr -Fields @('FEPort','FrontendPort')
                $bePortText = Get-CellValueAny -Row $rr -Fields @('BEPort','BackendPort')

                # RuleName/Port가 비어 있으면 해당 행은 배포 대상에서 제외(검증 오류 아님)
                if (-not $ruleName) { $j++; continue }
                if (-not $enableHaPorts -and (-not $fePortText -or -not $bePortText)) { $j++; continue }

                if (-not $lbNameSet.Contains($targetLb)) {
                    Add-Issue -Issues $issues -Type 'LB_Rule' -Row $j -ResourceName $targetLb -Field 'LBName' -Message 'LB 시트에 존재하지 않는 LBName 입니다.'
                }
                foreach ($f in @('FEName','BEPoolName')) {
                    if (-not (Get-CellValue -Row $rr -Field $f)) {
                        Add-Issue -Issues $issues -Type 'LB_Rule' -Row $j -ResourceName $targetLb -Field $f -Message '필수 값이 비어 있습니다.'
                    }
                }
                $j++
            }
        }
    }

    if ($Context.DeployType -contains 'VM') {
        $rows = Get-SheetRows -Context $Context -SheetCandidates @('VM','VM_PRD')
        $rows = Get-FilteredVmRows -Context $Context -Rows $rows
        $i = 1
        foreach ($r in $rows) {
            $vmName = Get-CellValue -Row $r -Field 'Name'
            if (-not $vmName) { $i++; continue }

            foreach ($field in @('RGname','Location','NicName','SubnetName','VnetName','PrivateIP','VmSize','OsDiskStorageType','AdminUsername','OsType')) {
                Test-RequiredField -Issues $issues -Type 'VM' -Row $i -Data $r -Field $field -ResourceName $vmName
            }

            $osType = Get-CellValue -Row $r -Field 'OsType'
            if ($osType -and $osType.Trim().ToUpperInvariant() -eq 'WINDOWS') {
                $computerName = Get-WindowsComputerName -VmName $vmName
                if (-not (Test-WindowsComputerName -Name $computerName)) {
                    Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'Name' -Message "Windows ComputerName 생성값이 유효하지 않습니다. 생성값=$computerName, 규칙: 1~15자, 영문/숫자/하이픈(-)만 허용, 숫자만으로 구성 불가"
                }
            }

            $useKvPassword = Convert-ToBoolean -Value (Get-CellValue -Row $r -Field 'UseKeyVaultPassword') -Default $false
            if ($useKvPassword) {
                $secretUri = Get-CellValue -Row $r -Field 'AdminPasswordSecretUri'
                $kvName = Get-CellValue -Row $r -Field 'AdminPasswordKVName'
                $secretName = Get-CellValue -Row $r -Field 'AdminPasswordSecretName'

                if ($secretUri) {
                    if ($secretUri -notmatch '/secrets/' -and -not $secretName) {
                        Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'AdminPasswordSecretName' -Message 'AdminPasswordSecretUri가 Vault URL인 경우 SecretName이 필요합니다.'
                    }
                    if ($secretUri -notmatch '/secrets/' -and -not $kvName -and $secretUri -notmatch 'https://[^./]+\.vault\.azure\.net/?$') {
                        Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'AdminPasswordKVName' -Message 'AdminPasswordSecretUri 또는 AdminPasswordKVName 값이 올바르지 않습니다.'
                    }
                } else {
                    if (-not $kvName) {
                        Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'AdminPasswordKVName' -Message 'UseKeyVaultPassword=Y 인 경우 값이 필요합니다.'
                    }
                    if (-not $secretName) {
                        Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'AdminPasswordSecretName' -Message 'UseKeyVaultPassword=Y 인 경우 값이 필요합니다.'
                    }
                }
            } else {
                Test-RequiredField -Issues $issues -Type 'VM' -Row $i -Data $r -Field 'AdminPassword' -ResourceName $vmName
            }

            if (-not (Get-CellValue -Row $r -Field 'OsDiskName')) {
                Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'OsDiskName' -Message '권장 필드가 비어 있어 기본값(<VM>-OsDisk)을 적용합니다.'
            }
            if (-not (Get-CellValue -Row $r -Field 'VnetRG')) {
                Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'VnetRG' -Message '권장 필드가 비어 있어 RGname 값을 대체 사용합니다.'
            }
            $desName = Get-CellValueAny -Row $r -Fields @('DESName','DiskEncryptionSetName')
            $desId = Get-CellValueAny -Row $r -Fields @('DiskEncryptionSetId','DesResourceId')
            if ($desName -and $desId) {
                Add-Issue -Issues $issues -Type 'VM' -Row $i -ResourceName $vmName -Field 'DESName/DiskEncryptionSetId' -Message 'DESName과 DiskEncryptionSetId가 모두 있으면 DiskEncryptionSetId를 우선 사용합니다.'
            }

            $sourceType = Get-VmSourceType -Row $r
            if ($sourceType -eq 'Marketplace') {
                foreach ($field in @('Publisher','Offer','Sku','Version')) {
                    Test-RequiredField -Issues $issues -Type 'VM' -Row $i -Data $r -Field $field -ResourceName $vmName
                }
            }
            else {
                Test-RequiredField -Issues $issues -Type 'VM' -Row $i -Data $r -Field 'ImageResourceId' -ResourceName $vmName
            }

            $i++
        }
    }

    # Cross-sheet consistency checks for CMK chain (KV -> DES -> VM)
    $kvRows = @()
    if ($Context.DeployType -contains 'DES') {
        try { $kvRows = @(Get-SheetRows -Context $Context -SheetCandidates @('KV','KeyVault') -Optional) } catch {}
    }
    $desRows = @()
    if ($Context.DeployType -contains 'DES') {
        try { $desRows = @(Get-SheetRows -Context $Context -SheetCandidates @('DES','DES_PRD') -Optional) } catch {}
    }
    $vmRows = @()
    if ($Context.DeployType -contains 'VM') {
        try {
            $vmRows = @(Get-SheetRows -Context $Context -SheetCandidates @('VM','VM_PRD') -Optional)
            $vmRows = @(Get-FilteredVmRows -Context $Context -Rows $vmRows)
        } catch {}
    }

    $kvNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $kvRows) {
        $n = Get-CellValueAny -Row $r -Fields @('KVName','KeyVaultName')
        if ($n) { [void]$kvNames.Add($n) }
    }

    if ($Context.DeployType -contains 'DES') {
        $desNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        $rowNo = 1
        foreach ($r in $desRows) {
            $desName = Get-CellValueAny -Row $r -Fields @('DESName','DiskEncryptionSetName')
            if ($desName) { [void]$desNames.Add($desName) }
            $refKvName = Get-CellValueAny -Row $r -Fields @('KVName','KeyVaultName')
            if ($refKvName -and $kvNames.Count -gt 0 -and -not $kvNames.Contains($refKvName)) {
                Add-Issue -Issues $issues -Type 'DES' -Row $rowNo -ResourceName $desName -Field 'KVName' -Message "KV 시트에 존재하지 않는 Key Vault 참조입니다: $refKvName"
            }
            $rowNo++
        }

        if ($Context.DeployType -contains 'VM') {
            $rowNo = 1
            foreach ($r in $vmRows) {
                $vmName = Get-CellValue -Row $r -Field 'Name'
                if (-not $vmName) { $rowNo++; continue }
                $vmDesName = Get-CellValueAny -Row $r -Fields @('DESName','DiskEncryptionSetName')
                $vmDesId = Get-CellValueAny -Row $r -Fields @('DiskEncryptionSetId','DesResourceId')
                if ($vmDesName -and $desNames.Count -gt 0 -and -not $desNames.Contains($vmDesName) -and -not $vmDesId) {
                    Add-Issue -Issues $issues -Type 'VM' -Row $rowNo -ResourceName $vmName -Field 'DESName' -Message "DES 시트에 존재하지 않는 DES 참조입니다: $vmDesName"
                }
                $rowNo++
            }
        }
    }

    End-Step 'Validate-Inputs'
    return $issues
}

function Ensure-ResourceGroup {
    param(
        [string]$Name,
        [string]$Location,
        [bool]$DryRun
    )

    if ($DryRun) {
        Write-Info "[DryRun] 리소스 그룹 생성 예정: $Name ($Location)"
        return
    }

    $rg = Invoke-WithRetry -Operation "Get-AzResourceGroup:$Name" -ScriptBlock {
        Get-AzResourceGroup -Name $Name -ErrorAction SilentlyContinue
    }
    if ($rg) {
        Write-Info "리소스 그룹 유지: $Name"
        return
    }

    Invoke-WithRetry -Operation "New-AzResourceGroup:$Name" -ScriptBlock {
        New-AzResourceGroup -Name $Name -Location $Location -Force | Out-Null
    }
    Write-Info "리소스 그룹 생성 완료: $Name"
}

function Deploy-ResourceGroups {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-ResourceGroups'
    $rows = Get-SheetRows -Context $Context -SheetCandidates @('RG')
    foreach ($row in $rows) {
        $name = Get-CellValue -Row $row -Field 'RGname'
        if (-not $name) { continue }
        $location = Get-CellValue -Row $row -Field 'Location'
        Ensure-ResourceGroup -Name $name -Location $location -DryRun:$Context.DryRun
    }
    End-Step 'Deploy-ResourceGroups'
}

function Get-VnetAddresses {
    param([psobject]$Row)
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($prop in $Row.PSObject.Properties) {
        if ($prop.Name -match '^VNet.*Address$' -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            $result.Add($prop.Value.ToString().Trim())
        }
    }
    $specialAddress = Get-CellValueAny -Row $Row -Fields @('VNet& Address')
    if ($specialAddress) { $result.Add($specialAddress) }
    return @($result)
}

function Get-VnetName {
    param([psobject]$Row)
    $name = Get-CellValueAny -Row $Row -Fields @('VNet name','VNet1 Name','VNet& Name')
    if ($name) { return $name }
    foreach ($prop in $Row.PSObject.Properties) {
        if ($prop.Name -match '^VNet.*Name$' -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            return $prop.Value.ToString().Trim()
        }
    }
    return $null
}

function Get-VnetSubnets {
    param([psobject]$Row)
    $pairs = @{}

    $indices = @(
        $Row.PSObject.Properties |
            Where-Object { $_.Name -match '^Subnet(\d+)\s+Name$' } |
            ForEach-Object { [int]$matches[1] } |
            Sort-Object -Unique
    )

    foreach ($i in $indices) {
        $nameKey = "Subnet$i Name"
        $addrKey = "Subnet$i Address"
        $subnetName = Get-CellValue -Row $Row -Field $nameKey
        $subnetAddress = Get-CellValue -Row $Row -Field $addrKey
        if ($subnetName -and $subnetAddress) {
            $pairs[$subnetName] = $subnetAddress
        }
    }
    return $pairs
}

function New-ManagedSubnetConfig {
    param(
        [string]$Name,
        [string]$AddressPrefix
    )

    return New-AzVirtualNetworkSubnetConfig `
        -Name $Name `
        -AddressPrefix $AddressPrefix `
        -DefaultOutboundAccess $false `
        -PrivateEndpointNetworkPoliciesFlag 'Enabled'
}

function Deploy-Vnets {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-Vnets'
    $rows = Get-SheetRows -Context $Context -SheetCandidates @('VNET')
    foreach ($row in $rows) {
        $rgName = Get-CellValue -Row $row -Field 'RGname'
        $location = Get-CellValue -Row $row -Field 'Location'
        $vnetName = Get-VnetName -Row $row
        if (-not $vnetName) { continue }

        Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun

        $addressPrefixes = Get-VnetAddresses -Row $row
        $subnets = Get-VnetSubnets -Row $row

        if ($Context.DryRun) {
            Write-Info "[DryRun] VNet 배포 예정: $vnetName"
            continue
        }

        $existing = Get-AzVirtualNetwork -ResourceGroupName $rgName -Name $vnetName -ErrorAction SilentlyContinue
        if (-not $existing) {
            if ($Context.DryRun) {
                Write-Info "[DryRun] VNet 생성 예정: $vnetName"
                continue
            }

            $firstSubnetName = $subnets.Keys | Select-Object -First 1
            $firstSubnetAddress = $subnets[$firstSubnetName]
            $vnet = New-AzVirtualNetwork -Name $vnetName -ResourceGroupName $rgName -Location $location -AddressPrefix $addressPrefixes -Subnet @(New-ManagedSubnetConfig -Name $firstSubnetName -AddressPrefix $firstSubnetAddress)
            foreach ($entry in $subnets.GetEnumerator() | Where-Object { $_.Key -ne $firstSubnetName }) {
                Add-AzVirtualNetworkSubnetConfig -Name $entry.Key -AddressPrefix $entry.Value -DefaultOutboundAccess $false -PrivateEndpointNetworkPoliciesFlag 'Enabled' -VirtualNetwork $vnet | Out-Null
            }
            $vnet | Set-AzVirtualNetwork | Out-Null
            Write-Info "VNet 생성 완료: $vnetName"
            continue
        }

        $changed = $false
        foreach ($prefix in $addressPrefixes) {
            if ($existing.AddressSpace.AddressPrefixes -notcontains $prefix) {
                $existing.AddressSpace.AddressPrefixes.Add($prefix)
                $changed = $true
            }
        }

        foreach ($entry in $subnets.GetEnumerator()) {
            $existingSubnet = $existing.Subnets | Where-Object Name -eq $entry.Key | Select-Object -First 1
            if (-not $existingSubnet) {
                Add-AzVirtualNetworkSubnetConfig -Name $entry.Key -AddressPrefix $entry.Value -DefaultOutboundAccess $false -PrivateEndpointNetworkPoliciesFlag 'Enabled' -VirtualNetwork $existing | Out-Null
                $changed = $true
                continue
            }

            if ($existingSubnet.PrivateEndpointNetworkPolicies -ne 'Enabled') {
                $existingSubnet.PrivateEndpointNetworkPolicies = 'Enabled'
                $changed = $true
            }

            if (($null -eq $existingSubnet.DefaultOutboundAccess) -or ([bool]$existingSubnet.DefaultOutboundAccess -ne $false)) {
                $existingSubnet.DefaultOutboundAccess = $false
                $changed = $true
            }
        }

        if ($changed) {
            if ($Context.DryRun) {
                Write-Info "[DryRun] VNet 업데이트 예정: $vnetName"
            } else {
                $existing | Set-AzVirtualNetwork | Out-Null
                Write-Info "VNet 업데이트 완료: $vnetName"
            }
        } else {
            Write-Info "VNet 변경 없음: $vnetName"
        }
    }
    End-Step 'Deploy-Vnets'
}

function Deploy-Storages {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-Storages'
    $rows = Get-SheetRows -Context $Context -SheetCandidates @('Storage')
    foreach ($row in $rows) {
        $name = Get-CellValueAny -Row $row -Fields @('Storage name','StorageName','StorgeName')
        if (-not $name) { continue }

        $name = $name.ToLowerInvariant()
        $rgName = Get-CellValue -Row $row -Field 'RGname'
        $location = Get-CellValue -Row $row -Field 'Location'
        $sku = Get-CellValue -Row $row -Field 'SkuName'
        if (-not $sku) { $sku = 'Standard_LRS' }
        $kind = Get-CellValue -Row $row -Field 'Kind'
        if (-not $kind) { $kind = 'StorageV2' }

        Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun

        if ($Context.DryRun) {
            Write-Info "[DryRun] 스토리지 생성 예정: $name"
            continue
        }

        $existing = Get-AzStorageAccount -ResourceGroupName $rgName -Name $name -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Info "스토리지 유지: $name"
            continue
        }

        New-AzStorageAccount -ResourceGroupName $rgName -Name $name -Location $location -SkuName $sku -Kind $kind -EnableHttpsTrafficOnly $true | Out-Null
        Write-Info "스토리지 생성 완료: $name"
    }
    End-Step 'Deploy-Storages'
}

function Deploy-KeyVaults {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-KeyVaults'
    $rows = Get-SheetRows -Context $Context -SheetCandidates @('KV','KeyVault')
    $runnerObjectId = $null
    if (-not $Context.DryRun) {
        $runnerObjectId = Resolve-CurrentPrincipalObjectId -Context $Context
    }

    foreach ($row in $rows) {
        if (-not (Test-IsEnabledRow -Row $row -Default $true)) { continue }
        $kvName = Get-CellValueAny -Row $row -Fields @('KVName','KeyVaultName')
        if (-not $kvName) { continue }

        $rgName = Get-CellValue -Row $row -Field 'RGname'
        $location = Get-CellValue -Row $row -Field 'Location'
        $keyName = Get-CellValueAny -Row $row -Fields @('KeyName','CMKKeyName')
        $secretName = Get-CellValue -Row $row -Field 'SecretName'
        $secretValue = Get-CellValue -Row $row -Field 'SecretValue'
        $secretContentType = Get-CellValueAny -Row $row -Fields @('SecretContentType','SecretContenctType')
        $secretExpiresText = Get-CellValue -Row $row -Field 'SecretExpiresOn'
        $secretNotBeforeText = Get-CellValue -Row $row -Field 'SecretNotBefore'
        $secretExpires = Convert-ToNullableDateTime -Value $secretExpiresText -FieldName 'SecretExpiresOn'
        $secretNotBefore = Convert-ToNullableDateTime -Value $secretNotBeforeText -FieldName 'SecretNotBefore'
        $skuName = Get-CellValue -Row $row -Field 'SkuName'
        if (-not $skuName) { $skuName = 'Standard' }
        $keyType = Get-CellValue -Row $row -Field 'KeyType'
        if (-not $keyType) { $keyType = 'RSA' }
        $keySizeText = Get-CellValue -Row $row -Field 'KeySize'
        $keySize = if ($keySizeText) { [int]$keySizeText } else { 2048 }
        $enableRbac = Convert-ToBoolean -Value (Get-CellValue -Row $row -Field 'EnableRbacAuthorization') -Default $true

        Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun

        if ($Context.DryRun) {
            $plan = New-Object System.Collections.Generic.List[string]
            if ($keyName) { $plan.Add("Key=$keyName") }
            if ($secretName) { $plan.Add("Secret=$secretName") }
            if ($plan.Count -eq 0) { $plan.Add('No key/secret') }
            Write-Info "[DryRun] Key Vault 배포 예정: $kvName, $($plan -join ', ')"
            continue
        }

        $kv = $null
        $kvByName = Get-AzKeyVault -VaultName $kvName -ErrorAction SilentlyContinue
        if ($kvByName) {
            if ($kvByName.ResourceGroupName -ieq $rgName) {
                $kv = $kvByName
                Write-Info "Key Vault 유지: $kvName"
            } else {
                throw "Key Vault 이름 '$kvName' 이(가) 이미 다른 리소스 그룹($($kvByName.ResourceGroupName))에서 사용 중입니다. KVName을 변경해 주세요."
            }
        }

        if (-not $kv) {
            $deletedKv = Get-AzKeyVault -InRemovedState -ErrorAction SilentlyContinue | Where-Object { $_.VaultName -ieq $kvName } | Select-Object -First 1
            if ($deletedKv) {
                throw "Key Vault 이름 '$kvName' 이(가) 삭제 보류(soft-delete) 상태입니다. Purge 후 재사용하거나 KVName을 변경해 주세요."
            }

            $createParams = @{
                Name                  = $kvName
                ResourceGroupName     = $rgName
                Location              = $location
                Sku                   = $skuName
                EnablePurgeProtection = $true
            }
            if (-not $enableRbac) {
                $createParams.DisableRbacAuthorization = $true
            }
            try {
                $kv = New-AzKeyVault @createParams
                Write-Info "Key Vault 생성 완료: $kvName"
            } catch {
                if ($_.Exception.Message -like '*already in use*' -or $_.Exception.Message -like '*VaultAlreadyExists*') {
                    throw "Key Vault 이름 '$kvName' 이(가) 전역에서 이미 사용 중이거나 soft-delete 상태입니다. KVName을 유니크하게 변경해 주세요."
                }
                throw
            }
        }

        if ($keyName) {
            Ensure-KvOperatorRoleAssignment -PrincipalId $runnerObjectId -Scope $kv.ResourceId
            $existingKey = Get-AzKeyVaultKey -VaultName $kvName -Name $keyName -ErrorAction SilentlyContinue
            if (-not $existingKey) {
                Add-KeyWithRetry -VaultName $kvName -KeyName $keyName -KeyType $keyType -KeySize $keySize
                Write-Info "Key Vault Key 생성 완료: $kvName/$keyName"
            } else {
                Write-Info "Key Vault Key 유지: $kvName/$keyName"
            }
        }

        if ($secretName) {
            Ensure-KvSecretsRoleAssignment -PrincipalId $runnerObjectId -Scope $kv.ResourceId
            $existingSecret = Get-AzKeyVaultSecret -VaultName $kvName -Name $secretName -ErrorAction SilentlyContinue
            if (-not $existingSecret) {
                Add-SecretWithRetry -VaultName $kvName -SecretName $secretName -SecretValue $secretValue -ContentType $secretContentType -Expires $secretExpires -NotBefore $secretNotBefore
                Write-Info "Key Vault Secret 생성 완료: $kvName/$secretName"
            } else {
                Write-Info "Key Vault Secret 유지: $kvName/$secretName"
            }
        }
    }
    End-Step 'Deploy-KeyVaults'
}

function Get-KeyVaultKeyUrl {
    param(
        [string]$VaultName,
        [string]$KeyName,
        [string]$KeyVersion
    )
    if ($KeyVersion) {
        $k = Get-AzKeyVaultKey -VaultName $VaultName -Name $KeyName -Version $KeyVersion -ErrorAction Stop
        return $k.Key.Kid
    }
    $latest = Get-AzKeyVaultKey -VaultName $VaultName -Name $KeyName -ErrorAction Stop
    return $latest.Key.Kid
}

function Ensure-DesRoleAssignment {
    param(
        [string]$PrincipalId,
        [string]$Scope
    )
    $roleName = 'Key Vault Crypto Service Encryption User'
    $exists = Get-AzRoleAssignment -ObjectId $PrincipalId -Scope $Scope -RoleDefinitionName $roleName -ErrorAction SilentlyContinue
    if (-not $exists) {
        New-AzRoleAssignment -ObjectId $PrincipalId -Scope $Scope -RoleDefinitionName $roleName -ErrorAction Stop | Out-Null
        Write-Info "DES 권한 할당 완료: $roleName"
    } else {
        Write-Info "DES 권한 유지: $roleName"
    }
}

function Deploy-DiskEncryptionSets {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-DiskEncryptionSets'
    $rows = Get-SheetRows -Context $Context -SheetCandidates @('DES','DES_PRD')
    foreach ($row in $rows) {
        if (-not (Test-IsEnabledRow -Row $row -Default $true)) { continue }
        $desName = Get-CellValueAny -Row $row -Fields @('DESName','DiskEncryptionSetName')
        if (-not $desName) { continue }

        $rgName = Get-CellValue -Row $row -Field 'RGname'
        $location = Get-CellValue -Row $row -Field 'Location'
        $kvName = Get-CellValueAny -Row $row -Fields @('KeyVaultName','KVName')
        $kvRg = Get-CellValueAny -Row $row -Fields @('KeyVaultRG','KVRG')
        if (-not $kvRg) { $kvRg = $rgName }
        $keyName = Get-CellValueAny -Row $row -Fields @('KeyName','CMKKeyName')
        $keyVersion = Get-CellValue -Row $row -Field 'KeyVersion'

        Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun

        if ($Context.DryRun) {
            Write-Info "[DryRun] DES 배포 예정: $desName (KV=$kvName, Key=$keyName)"
            continue
        }

        $kv = Get-AzKeyVault -VaultName $kvName -ResourceGroupName $kvRg -ErrorAction Stop
        $keyUrl = Get-KeyVaultKeyUrl -VaultName $kvName -KeyName $keyName -KeyVersion $keyVersion

        $des = Get-AzDiskEncryptionSet -ResourceGroupName $rgName -Name $desName -ErrorAction SilentlyContinue
        if (-not $des) {
            $desConfig = New-AzDiskEncryptionSetConfig -Location $location -SourceVaultId $kv.ResourceId -KeyUrl $keyUrl -IdentityType 'SystemAssigned'
            $des = New-AzDiskEncryptionSet -ResourceGroupName $rgName -Name $desName -InputObject $desConfig -ErrorAction Stop
            Write-Info "DES 생성 완료: $desName"
        } else {
            Write-Info "DES 유지: $desName"
        }

        if (-not $des.Identity -or -not $des.Identity.PrincipalId) {
            throw "DES 시스템 할당 ID를 확인할 수 없습니다: $desName"
        }

        Ensure-DesRoleAssignment -PrincipalId $des.Identity.PrincipalId -Scope $kv.ResourceId
    }
    End-Step 'Deploy-DiskEncryptionSets'
}

function Resolve-CurrentPrincipalObjectId {
    param([DeploymentContext]$Context)

    $ctx = Get-AzContext -ErrorAction Stop
    if (-not $ctx -or -not $ctx.Account) {
        throw '현재 Azure 계정 정보를 확인할 수 없습니다.'
    }

    if ($ctx.Account.Type -eq 'User') {
        $user = Get-AzADUser -UserPrincipalName $ctx.Account.Id -ErrorAction SilentlyContinue
        if ($user) { return $user.Id }
        $users = Get-AzADUser -StartsWith ($ctx.Account.Id.Split('@')[0]) -ErrorAction SilentlyContinue
        if ($users) {
            $matched = $users | Where-Object { $_.UserPrincipalName -ieq $ctx.Account.Id } | Select-Object -First 1
            if ($matched) { return $matched.Id }
        }
    } elseif ($ctx.Account.Type -eq 'ServicePrincipal') {
        $sp = Get-AzADServicePrincipal -ApplicationId $ctx.Account.Id -ErrorAction SilentlyContinue
        if ($sp) { return $sp.Id }
    }

    # Graph 조회가 제한된 환경에서는 액세스 토큰의 oid 클레임을 사용한다.
    try {
        $access = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -ErrorAction Stop
        $rawToken = $access.Token
        if ($rawToken -is [securestring]) {
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($rawToken)
            try { $rawToken = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        }

        $parts = $rawToken.Split('.')
        if ($parts.Length -ge 2) {
            $payload = $parts[1].Replace('-', '+').Replace('_', '/')
            switch ($payload.Length % 4) {
                2 { $payload += '==' }
                3 { $payload += '=' }
            }
            $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
            $obj = $json | ConvertFrom-Json
            if ($obj.oid) { return [string]$obj.oid }
        }
    } catch {
        Write-WarnLog "실행자 ObjectId 토큰 파싱 fallback 실패: $($_.Exception.Message)"
    }

    throw "실행 계정의 ObjectId를 조회하지 못했습니다. AccountType=$($ctx.Account.Type), AccountId=$($ctx.Account.Id)"
}

function Ensure-KvOperatorRoleAssignment {
    param(
        [string]$PrincipalId,
        [string]$Scope
    )

    $roleName = 'Key Vault Crypto Officer'
    $exists = Get-AzRoleAssignment -ObjectId $PrincipalId -Scope $Scope -RoleDefinitionName $roleName -ErrorAction SilentlyContinue
    if (-not $exists) {
        New-AzRoleAssignment -ObjectId $PrincipalId -Scope $Scope -RoleDefinitionName $roleName -ErrorAction Stop | Out-Null
        Write-Info "실행 계정 권한 할당 완료: $roleName"
    } else {
        Write-Info "실행 계정 권한 유지: $roleName"
    }
}

function Ensure-KvSecretsRoleAssignment {
    param(
        [string]$PrincipalId,
        [string]$Scope
    )

    $roleName = 'Key Vault Secrets Officer'
    $exists = Get-AzRoleAssignment -ObjectId $PrincipalId -Scope $Scope -RoleDefinitionName $roleName -ErrorAction SilentlyContinue
    if (-not $exists) {
        New-AzRoleAssignment -ObjectId $PrincipalId -Scope $Scope -RoleDefinitionName $roleName -ErrorAction Stop | Out-Null
        Write-Info "실행 계정 권한 할당 완료: $roleName"
    } else {
        Write-Info "실행 계정 권한 유지: $roleName"
    }
}

function Add-KeyWithRetry {
    param(
        [string]$VaultName,
        [string]$KeyName,
        [string]$KeyType,
        [int]$KeySize
    )

    $maxTry = 8
    for ($try = 1; $try -le $maxTry; $try++) {
        try {
            Add-AzKeyVaultKey -VaultName $VaultName -Name $KeyName -Destination Software -Size $KeySize -KeyType $KeyType -ErrorAction Stop | Out-Null
            return
        } catch {
            $msg = $_.Exception.Message
            if ($msg -like '*ForbiddenByRbac*' -or $msg -like '*not authorized*') {
                if ($try -lt $maxTry) {
                    Write-WarnLog "Key Vault RBAC 전파 대기 중... ($try/$maxTry)"
                    Start-Sleep -Seconds 15
                    continue
                }
            }
            throw
        }
    }
}

function Add-SecretWithRetry {
    param(
        [string]$VaultName,
        [string]$SecretName,
        [string]$SecretValue,
        [string]$ContentType,
        [Nullable[datetime]]$Expires,
        [Nullable[datetime]]$NotBefore
    )

    $secureSecretValue = (ConvertTo-SecureString -String $SecretValue -AsPlainText -Force)
    $setParams = @{
        VaultName = $VaultName
        Name = $SecretName
        SecretValue = $secureSecretValue
    }
    if ($ContentType) { $setParams['ContentType'] = $ContentType }
    if ($Expires) { $setParams['Expires'] = $Expires.Value }
    if ($NotBefore) { $setParams['NotBefore'] = $NotBefore.Value }

    $maxTry = 8
    for ($try = 1; $try -le $maxTry; $try++) {
        try {
            Set-AzKeyVaultSecret @setParams -ErrorAction Stop | Out-Null
            return
        } catch {
            $msg = $_.Exception.Message
            if ($msg -like '*ForbiddenByRbac*' -or $msg -like '*not authorized*') {
                if ($try -lt $maxTry) {
                    Write-WarnLog "Key Vault RBAC 전파 대기 중... ($try/$maxTry)"
                    Start-Sleep -Seconds 15
                    continue
                }
            }
            throw
        }
    }
}

function Convert-ToRuleValue {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '*' }
    $raw = $Value.Trim()
    if ($raw -eq '*' -or $raw.ToUpperInvariant() -eq 'ANY') { return '*' }
    if ($raw.Contains(',')) { return $raw.Replace(' ','').Split(',') }
    return $raw
}

function Convert-LbProtocol {
    param(
        [string]$Value,
        [string]$Default = 'Tcp'
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    $raw = $Value.Trim().ToUpperInvariant()
    switch ($raw) {
        'TCP' { return 'Tcp' }
        'UDP' { return 'Udp' }
        'HTTP' { return 'Http' }
        'HTTPS' { return 'Https' }
        'ALL' { return 'All' }
        default { return $Default }
    }
}

function Convert-LbLoadDistribution {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'Default' }
    $raw = $Value.Trim().ToUpperInvariant()
    switch ($raw) {
        'NONE' { return 'Default' }
        'DEFAULT' { return 'Default' }
        'SOURCEIP' { return 'SourceIP' }
        'CLIENTIP' { return 'SourceIP' }
        'SOURCEIPPROTOCOL' { return 'SourceIPProtocol' }
        'CLIENTIPANDPROTOCOL' { return 'SourceIPProtocol' }
        default { return 'Default' }
    }
}

function Get-LbZoneArray {
    param([psobject]$Row)
    $zoneMode = Get-CellValueAny -Row $Row -Fields @('FEZoneMode','FrontendZoneMode')
    if (-not $zoneMode) { return @() }

    if ($zoneMode.Trim().ToUpperInvariant() -eq 'ZONAL') {
        $zone = Get-CellValueAny -Row $Row -Fields @('FEZone','FrontendZone')
        if ($zone) { return @($zone) }
    }
    return @()
}

function Deploy-LoadBalancers {
param([DeploymentContext]$Context)
    Start-Step 'Deploy-LoadBalancers'
    
    # 1. 엑셀 시트에서 데이터 로드
    $lbRows = Get-SheetRows -Context $Context -SheetCandidates @('LB','LB_PRD','Load Balancer')
    $probeRows = Get-SheetRows -Context $Context -SheetCandidates @('LB_Probe','LB_PRD_Probe','Load Balancer_Probe') -Optional
    $ruleRows = Get-SheetRows -Context $Context -SheetCandidates @('LB_Rule','LB_PRD_Rule','Load Balancer_Rule') -Optional

    # 2. -Option 필터(예: 'DTG')가 전달된 경우 엑셀 데이터 매핑 및 필터링
    if ($Context.VmRoleFilter.Count -gt 0) {
        $lbRows = @($lbRows | Where-Object { $target = Get-CellValue -Row $_ -Field 'Role'; $target -and $Context.VmRoleFilter -contains $target })
        
        $selectedLbNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $lbRows) {
            $lbName = Get-CellValueAny -Row $row -Fields @('LBName','LoadBalancerName')
            if ($lbName) { [void]$selectedLbNames.Add($lbName) }
        }
        
        $probeRows = @($probeRows | Where-Object { $target = Get-CellValue -Row $_ -Field 'LBName'; $target -and $selectedLbNames.Contains($target) })
        $ruleRows = @($ruleRows | Where-Object { $target = Get-CellValue -Row $_ -Field 'LBName'; $target -and $selectedLbNames.Contains($target) })
    }

    # 3. 로드밸런서별로 그룹핑하여 순차 처리 시작
    $groups = $lbRows | Where-Object { Get-CellValueAny -Row $_ -Fields @('LBName','LoadBalancerName') } | Group-Object -Property LBName
    foreach ($group in $groups) {
        $base = $group.Group | Select-Object -First 1
        $lbName = $group.Name
        $rgName = Get-CellValue -Row $base -Field 'RGName'
        $location = Get-CellValue -Row $base -Field 'Location'
        $skuName = Convert-LbSku -Value (Get-CellValue -Row $base -Field 'SKU') -Default 'Standard'
        
        Write-Info "로드밸런서 상태 확인 및 정리 중: $lbName (RG: $rgName)"
        
        # 4. 기존 로드밸런서 존재 여부 체크
        $lb = Get-AzLoadBalancer -Name $lbName -ResourceGroupName $rgName -ErrorAction SilentlyContinue
        
        # -------------------------------------------------------------
        # [신규 추가] 기존 자식 리소스 강제 삭제(Clean-up) 로직
        # -------------------------------------------------------------
        if ($lb) {
            $isChanged = $false
            
            # (1) 부하 분산 규칙 초기화 (의존성 때문에 규칙을 가장 먼저 제거해야 합니다)
            if ($Context.DeployType -contains 'LB_RULE' -or $Context.DeployType -contains 'LB') {
                if ($lb.LoadBalancingRules.Count -gt 0) {
                    Write-Info "기존 부하 분산 규칙(LB Rules)을 초기화합니다: $lbName"
                    $lb.LoadBalancingRules.Clear()
                    $isChanged = $true
                }
            }
            
            # (2) 상태 프로브 초기화
            if ($Context.DeployType -contains 'LB_PROBE' -or $Context.DeployType -contains 'LB') {
                if ($lb.Probes.Count -gt 0) {
                    Write-Info "기존 상태 프로브(Probes)를 초기화합니다: $lbName"
                    $lb.Probes.Clear()
                    $isChanged = $true
                }
            }
            
            # (3) 프론트엔드 IP 구성 초기화 (Zone 구성을 변경하려면 기존 것이 완전히 밀려야 함)
            if ($Context.DeployType -contains 'LB') {
                if ($lb.FrontendIpConfigurations.Count -gt 0) {
                    Write-Info "기존 프론트엔드 IP 구성(Frontend IP)을 초기화합니다: $lbName"
                    $lb.FrontendIpConfigurations.Clear()
                    $isChanged = $true
                }
            }
            
            # 변경 사항이 있을 경우 Azure 실 서버에 세팅 반영하여 빈 껍데기로 만듦
            if ($isChanged) {
                if (-not $Context.DryRun) {
                    Write-Info "Azure 기존 내부 설정 동기화 중 (Set-AzLoadBalancer)..."
                    $lb | Set-AzLoadBalancer -ErrorAction Stop | Out-Null
                    # 깨끗해진 로드밸런서 객체를 다시 불러옴
                    $lb = Get-AzLoadBalancer -Name $lbName -ResourceGroupName $rgName -ErrorAction Stop
                } else {
                    Write-Info "[DryRun] 기존 자식 리소스 규칙/프로브/프론트엔드가 초기화될 예정입니다."
                }
            }
        }
        # -------------------------------------------------------------

        # 5. 프론트엔드 IP 정보 생성 및 주입
        $feConfigs = @()
        foreach ($row in $group.Group) {
            if (-not (Test-IsEnabledRow -Row $row)) { continue }
            
            $feName = Get-CellValue -Row $row -Field 'FEName'
            $feType = Get-CellValue -Row $row -Field 'FEType'
            $feVnetRg = Get-CellValue -Row $row -Field 'FEVNetRG'
            $feVnetName = Get-CellValue -Row $row -Field 'FEVNetName'
            $feSubnetName = Get-CellValue -Row $row -Field 'FESubnetName'
            $privateIpAllocation = Get-CellValue -Row $row -Field 'PrivateIPAllocation'
            if (-not $privateIpAllocation) { $privateIpAllocation = 'Dynamic' }
            $privateIpAddress = Get-CellValue -Row $row -Field 'PrivateIPAddress'
            
            # [보완] FEZoneMode가 ZoneRedundant이거나 FEZone에 영역이 적혀있을 때 '1','2','3' 명시적 추출
            $feZoneMode = Get-CellValue -Row $row -Field 'FEZoneMode'
            $zoneRaw = Get-CellValue -Row $row -Field 'FEZone'
            $zones = @()
            if ($feZoneMode -eq 'ZoneRedundant') {
                $zones = @('1', '2', '3')
            } elseif ($zoneRaw) {
                $zones = $zoneRaw.Split(',').Trim()
            }

            if (-not $feName) { continue }

            if ($feType -and $feType.ToUpperInvariant() -eq 'INTERNAL') {
                if (-not $feVnetRg) { $feVnetRg = $rgName }
                $subnet = Get-AzVirtualNetwork -Name $feVnetName -ResourceGroupName $feVnetRg -ErrorAction SilentlyContinue | Get-AzVirtualNetworkSubnetConfig -Name $feSubnetName -ErrorAction SilentlyContinue
                if (-not $subnet) {
                    throw "LB Frontend Subnet을 찾을 수 없습니다. VNet=$feVnetName, Subnet=$feSubnetName, RG=$feVnetRg"
                }
                
                $feParams = @{
                    Name = $feName
                    SubnetId = $subnet.Id
                }
                if ($privateIpAllocation.ToUpperInvariant() -eq 'STATIC' -and $privateIpAddress) {
                    $feParams['PrivateIpAddress'] = $privateIpAddress
                }
                if ($zones.Count -gt 0) {
                    $feParams['Zone'] = $zones
                }

                if ($Context.DryRun) {
                    Write-Info "[DryRun] Frontend IP 구성 예정: $feName (Zones=$($zones -join ','))"
                } else {
                    $feConfigs += New-AzLoadBalancerFrontendIpConfig @feParams
                }
            }
        }

        # 6. 백엔드 주소 풀(Backend Pool) 구성 설정
        $beConfigs = @()
        $bePoolNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $group.Group) {
            $bePoolName = Get-CellValue -Row $row -Field 'BEPoolName'
            if ($bePoolName -and -not $bePoolNames.Contains($bePoolName)) {
                [void]$bePoolNames.Add($bePoolName)
                if ($Context.DryRun) {
                    Write-Info "[DryRun] Backend Address Pool 구성 예정: $bePoolName"
                } else {
                    $beConfigs += New-AzLoadBalancerBackendAddressPoolConfig -Name $bePoolName
                }
            }
        }

        # 7. 큰 껍데기 LB 리소스 자체의 유무에 따른 동기화/생성 처리
        Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun
        if (-not $lb) {
            if ($Context.DryRun) {
                Write-Info "[DryRun] Load Balancer 신규 생성 예정: $lbName"
            } else {
                $lbParams = @{
                    ResourceGroupName = $rgName
                    Name = $lbName
                    Location = $location
                    Sku = $skuName
                }
                if ($feConfigs.Count -gt 0) { $lbParams['FrontendIpConfiguration'] = $feConfigs }
                if ($beConfigs.Count -gt 0) { $lbParams['BackendAddressPool'] = $beConfigs }
                
                $lb = New-AzLoadBalancer @lbParams -ErrorAction Stop
                Write-Info "Load Balancer 기본 뼈대 생성 완료: $lbName"
            }
        } else {
            if (-not $Context.DryRun) {
                if ($feConfigs.Count -gt 0) { $lb.FrontendIpConfigurations = $feConfigs }
                if ($beConfigs.Count -gt 0) { $lb.BackendAddressPools = $beConfigs }
                $lb | Set-AzLoadBalancer -ErrorAction Stop | Out-Null
                $lb = Get-AzLoadBalancer -Name $lbName -ResourceGroupName $rgName -ErrorAction Stop
                Write-Info "Load Balancer 기본 뼈대 업데이트 완료: $lbName"
            }
        }

        # 8. 상태 프로브(Probes) 생성 및 적용 루프
        $currentLbProbes = @($probeRows | Where-Object { (Get-CellValue -Row $_ -Field 'LBName') -eq $lbName })
        foreach ($pr in $currentLbProbes) {
            if (-not (Test-IsEnabledRow -Row $pr)) { continue }
            
            $probeName = Get-CellValue -Row $pr -Field 'ProbeName'
            $probePortText = Get-CellValue -Row $pr -Field 'Port'
            if (-not $probeName -or -not $probePortText) {
                Write-WarnLog "[SKIP] LB Probe 필수값 누락(Port): $lbName/$probeName"
                continue
            }
            
            $probeIntervalText = Get-CellValue -Row $pr -Field 'IntervalInSeconds'
            $probeCountText = Get-CellValue -Row $pr -Field 'ProbeCount'
            $probeProtocol = Convert-LbProtocol -Value (Get-CellValue -Row $pr -Field 'Protocol') -Default 'Tcp'
            if ($probeProtocol -eq 'All') { $probeProtocol = 'Tcp' }

            $probeParams = @{
                LoadBalancer = $lb
                Name = $probeName
                Port = [int]$probePortText
                Protocol = $probeProtocol
            }
            if ($probeIntervalText) { $probeParams['IntervalInSeconds'] = [int]$probeIntervalText }
            if ($probeCountText) { $probeParams['ProbeCount'] = [int]$probeCountText }
            
            $requestPath = Get-CellValue -Row $pr -Field 'RequestPath'
            if ($probeProtocol -eq 'Http' -or $probeProtocol -eq 'Https') {
                $probeParams['RequestPath'] = if ($requestPath) { $requestPath } else { '/' }
            }

            if ($Context.DryRun) {
                Write-Info "[DryRun] LB Probe 설정 예정: $probeName (${probeProtocol}:${probePortText})"
            } else {
                $lb | Add-AzLoadBalancerProbeConfig @probeParams | Out-Null
                Write-Info "LB Probe 추가 구성 적용: $probeName"
            }
        }
        if (-not $Context.DryRun -and $currentLbProbes.Count -gt 0) {
            $lb | Set-AzLoadBalancer -ErrorAction Stop | Out-Null
            $lb = Get-AzLoadBalancer -Name $lbName -ResourceGroupName $rgName -ErrorAction Stop
        }

        # 9. 부하 분산 규칙(Rules) 생성 및 최종 바인딩 루프
        $currentLbRules = @($ruleRows | Where-Object { (Get-CellValue -Row $_ -Field 'LBName') -eq $lbName })
        foreach ($rr in $currentLbRules) {
            if (-not (Test-IsEnabledRow -Row $rr)) { continue }
            
            $ruleName = Get-CellValue -Row $rr -Field 'RuleName'
            $feName = Get-CellValue -Row $rr -Field 'FEName'
            $bePoolName = Get-CellValue -Row $rr -Field 'BEPoolName'
            $fePortText = Get-CellValue -Row $rr -Field 'FEPort'
            $bePortText = Get-CellValue -Row $rr -Field 'BEPort'
            $protocol = Convert-LbProtocol -Value (Get-CellValue -Row $rr -Field 'Protocol') -Default 'Tcp'
            
            $idleTimeoutValue = 4
            $idleTimeoutText = Get-CellValue -Row $rr -Field 'IdleTimeouInMin'
            if ($idleTimeoutText) { $idleTimeoutValue = [int]$idleTimeoutText }

            if (-not $ruleName -or -not $fePortText) { continue }

            if ($Context.DryRun) {
                Write-Info "[DryRun] LB Rule 구성 예정: $ruleName ($protocol/FE:$fePortText -> BE:$bePortText)"
            } else {
                $ruleFrontend = $lb.FrontendIpConfigurations | Where-Object Name -eq $feName | Select-Object -First 1
                $ruleBackend = $lb.BackendAddressPools | Where-Object Name -eq $bePoolName | Select-Object -First 1
                
                $ruleParams = @{
                    LoadBalancer = $lb
                    Name = $ruleName
                    Protocol = $protocol
                    FrontendPort = [int]$fePortText
                    BackendPort = if ($bePortText) { [int]$bePortText } else { [int]$fePortText }
                    FrontendIpConfiguration = $ruleFrontend
                    BackendAddressPool = $ruleBackend
                    EnableHAPorts = (Convert-ToBoolean -Value (Get-CellValue -Row $rr -Field 'EnableHAports') -Default $false)
                    EnableTcpReset = (Convert-ToBoolean -Value (Get-CellValue -Row $rr -Field 'EnableTcpReset') -Default $true)
                    FloatingIP = (Convert-ToBoolean -Value (Get-CellValue -Row $rr -Field 'EnableFloatingIP') -Default $false)
                    LoadDistribution = (Convert-LbLoadDistribution -Value (Get-CellValue -Row $rr -Field 'LoadDistribution'))
                    IdleTimeoutInMinutes = $idleTimeoutValue
                }

                $probeName = Get-CellValue -Row $rr -Field 'ProbeName'
                if ($probeName) {
                    $ruleProbe = $lb.Probes | Where-Object Name -eq $probeName | Select-Object -First 1
                    if ($ruleProbe) {
                        $ruleParams['Probe'] = $ruleProbe
                    } else {
                        Write-WarnLog "LB Rule의 Probe를 찾지 못해 바인딩을 건너뜁니다: $probeName"
                    }
                }

                $lb | Add-AzLoadBalancerRuleConfig @ruleParams | Out-Null
                Write-Info "LB Rule 추가 구성 적용: $ruleName"
            }
        }
        if (-not $Context.DryRun -and $currentLbRules.Count -gt 0) {
            $lb | Set-AzLoadBalancer -ErrorAction Stop | Out-Null
            Write-Info "로드밸런서 자식 리소스 재배포 최종 완료: $lbName"
        }
    }
    End-Step 'Deploy-LoadBalancers'
}

function Deploy-Udrs {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-Udrs'
    $rows = Get-SheetRows -Context $Context -SheetCandidates @('UDR','RouteTable','UDR_Route')
    $rows = Get-FilteredRowsByOption -Rows $rows -OptionFilters $Context.VmRoleFilter -Columns @('Option','UDRType','Direction','Role')

    $groups = $rows |
        Where-Object {
            (Test-IsEnabledRow -Row $_ -Default $true) -and
            (Get-CellValueAny -Row $_ -Fields @('UDRName','RouteTableName'))
        } |
        Group-Object -Property { ("{0}|{1}" -f (Get-CellValueAny -Row $_ -Fields @('RGname','RGName')), (Get-CellValueAny -Row $_ -Fields @('UDRName','RouteTableName'))) }

    foreach ($group in $groups) {
        $base = $group.Group | Select-Object -First 1
        $rgName = Get-CellValueAny -Row $base -Fields @('RGname','RGName')
        $location = Get-CellValue -Row $base -Field 'Location'
        $udrName = Get-CellValueAny -Row $base -Fields @('UDRName','RouteTableName')
        $disableBgp = Convert-ToBoolean -Value (Get-CellValue -Row $base -Field 'DisableBGPRoutePropagation') -Default $false

        Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun

        if ($Context.DryRun) {
            Write-Info "[DryRun] UDR 배포 예정: $udrName (RG=$rgName, DisableBGPRoutePropagation=$disableBgp)"
            foreach ($row in $group.Group) {
                if (-not (Test-IsEnabledRow -Row $row -Default $true)) { continue }
                $routeName = Get-CellValue -Row $row -Field 'RouteName'
                if (-not $routeName) { continue }
                $addressPrefix = Get-CellValue -Row $row -Field 'AddressPrefix'
                $nextHopType = Convert-ToRouteNextHopType -Value (Get-CellValue -Row $row -Field 'NextHopType')
                $nextHopIp = Get-CellValue -Row $row -Field 'NextHopIpAddress'
                if ($nextHopType -eq 'VirtualAppliance') {
                    Write-Info "[DryRun] UDR Route 추가/갱신 예정: $udrName/$routeName ($addressPrefix -> $nextHopType $nextHopIp)"
                } else {
                    Write-Info "[DryRun] UDR Route 추가/갱신 예정: $udrName/$routeName ($addressPrefix -> $nextHopType)"
                }

                $associateSubnet = Convert-ToBoolean -Value (Get-CellValue -Row $row -Field 'AssociateToSubnet') -Default $false
                if ($associateSubnet) {
                    $vnetName = Get-CellValueAny -Row $row -Fields @('VnetName','VirtualNetworkName','VNetName')
                    $subnetName = Get-CellValueAny -Row $row -Fields @('SubnetName','Subnet')
                    $vnetRg = Get-CellValueAny -Row $row -Fields @('VnetRG','VNetRG','VirtualNetworkRG')
                    if (-not $vnetRg) { $vnetRg = $rgName }
                    Write-Info "[DryRun] Subnet-UDR 연결 예정: $vnetRg/$vnetName/$subnetName -> $udrName"
                }
            }
            continue
        }

        $routeTable = Get-AzRouteTable -ResourceGroupName $rgName -Name $udrName -ErrorAction SilentlyContinue
        if (-not $routeTable) {
            $routeTable = New-AzRouteTable -ResourceGroupName $rgName -Name $udrName -Location $location -DisableBgpRoutePropagation:$disableBgp -ErrorAction Stop
            Write-Info "UDR 생성 완료: $udrName"
        } else {
            Write-Info "UDR 유지: $udrName"
            if ([bool]$routeTable.DisableBgpRoutePropagation -ne $disableBgp) {
                $routeTable.DisableBgpRoutePropagation = $disableBgp
                $routeTable = $routeTable | Set-AzRouteTable -ErrorAction Stop
                Write-Info "UDR BGP 전파 설정 갱신: $udrName (DisableBGPRoutePropagation=$disableBgp)"
            }
        }

        $tableChanged = $false
        foreach ($row in $group.Group) {
            if (-not (Test-IsEnabledRow -Row $row -Default $true)) { continue }

            $routeName = Get-CellValue -Row $row -Field 'RouteName'
            if (-not $routeName) { continue }
            $addressPrefix = Get-CellValue -Row $row -Field 'AddressPrefix'
            $nextHopType = Convert-ToRouteNextHopType -Value (Get-CellValue -Row $row -Field 'NextHopType')
            $nextHopIp = Get-CellValue -Row $row -Field 'NextHopIpAddress'

            $routeParams = @{
                Name          = $routeName
                AddressPrefix = $addressPrefix
                NextHopType   = $nextHopType
                RouteTable    = $routeTable
            }
            if ($nextHopType -eq 'VirtualAppliance') {
                $routeParams['NextHopIpAddress'] = $nextHopIp
            }

            $existingRoute = $routeTable.Routes | Where-Object { $_.Name -eq $routeName } | Select-Object -First 1
            if ($existingRoute) {
                $routeTable = Set-AzRouteConfig @routeParams
            } else {
                $routeTable = Add-AzRouteConfig @routeParams
            }
            $tableChanged = $true
            Write-Info "UDR Route 반영(대기): $udrName/$routeName"
        }

        if ($tableChanged) {
            $routeTable = $routeTable | Set-AzRouteTable -ErrorAction Stop
            Write-Info "UDR Route 업데이트 완료: $udrName"
        }

        $subnetLinkSet = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $group.Group) {
            if (-not (Test-IsEnabledRow -Row $row -Default $true)) { continue }
            $associateSubnet = Convert-ToBoolean -Value (Get-CellValue -Row $row -Field 'AssociateToSubnet') -Default $false
            if (-not $associateSubnet) { continue }

            $vnetName = Get-CellValueAny -Row $row -Fields @('VnetName','VirtualNetworkName','VNetName')
            $subnetName = Get-CellValueAny -Row $row -Fields @('SubnetName','Subnet')
            $vnetRg = Get-CellValueAny -Row $row -Fields @('VnetRG','VNetRG','VirtualNetworkRG')
            if (-not $vnetRg) { $vnetRg = $rgName }
            $linkKey = "$vnetRg|$vnetName|$subnetName"
            if ($subnetLinkSet.Contains($linkKey)) { continue }
            [void]$subnetLinkSet.Add($linkKey)

            $vnet = Get-AzVirtualNetwork -ResourceGroupName $vnetRg -Name $vnetName -ErrorAction Stop
            $subnet = $vnet.Subnets | Where-Object { $_.Name -eq $subnetName } | Select-Object -First 1
            if (-not $subnet) {
                throw "UDR 연결 대상 Subnet을 찾을 수 없습니다. VNet=$vnetName, Subnet=$subnetName, RG=$vnetRg"
            }

            if ($subnet.RouteTable -and $subnet.RouteTable.Id -eq $routeTable.Id) {
                Write-Info "Subnet-UDR 연결 유지: $vnetRg/$vnetName/$subnetName -> $udrName"
                continue
            }

            $subnet.RouteTable = $routeTable
            $null = Set-AzVirtualNetwork -VirtualNetwork $vnet -ErrorAction Stop
            Write-Info "Subnet-UDR 연결 완료: $vnetRg/$vnetName/$subnetName -> $udrName"
        }
    }

    End-Step 'Deploy-Udrs'
}

function Deploy-Nsgs {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-Nsgs'
    $nsgRows = Get-SheetRows -Context $Context -SheetCandidates @('NSG')
    $ruleRows = Get-SheetRows -Context $Context -SheetCandidates @('NSG_Detail','NSG_Rule','NSG_PRD_Rule') -Optional
    $vnetRows = Get-SheetRows -Context $Context -SheetCandidates @('VNET') -Optional
    $nsgRows = Get-FilteredRowsByOption -Rows $nsgRows -OptionFilters $Context.VmRoleFilter -Columns @('Role','Option')

    $selectedNsgNames = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($nr in $nsgRows) {
        $nn = Get-CellValue -Row $nr -Field 'NSGName'
        if ($nn) { [void]$selectedNsgNames.Add($nn) }
    }
    if ($selectedNsgNames.Count -gt 0) {
        $ruleRows = @($ruleRows | Where-Object {
            $target = Get-CellValue -Row $_ -Field 'NSGName'
            $target -and $selectedNsgNames.Contains($target)
        })
    }

    $vnetRgMap = @{}
    foreach ($vr in $vnetRows) {
        $vnetNameFromSheet = Get-VnetName -Row $vr
        $vnetRgFromSheet = Get-CellValue -Row $vr -Field 'RGname'
        if ($vnetNameFromSheet -and $vnetRgFromSheet -and -not $vnetRgMap.ContainsKey($vnetNameFromSheet)) {
            $vnetRgMap[$vnetNameFromSheet] = $vnetRgFromSheet
        }
    }

    $groups = $nsgRows | Where-Object { Get-CellValue -Row $_ -Field 'NSGName' } | Group-Object -Property NSGName
    foreach ($group in $groups) {
        $base = $group.Group | Select-Object -First 1
        $nsgName = Get-CellValue -Row $base -Field 'NSGName'
        $rgName = Get-CellValue -Row $base -Field 'RG'

        if ($Context.DryRun) {
            Write-Info "[DryRun] NSG 배포 예정: $nsgName (RG=$rgName)"
            foreach ($row in $group.Group) {
                $useNic = (Get-CellValue -Row $row -Field 'UseNIC')
                $nicName = Get-CellValueAny -Row $row -Fields @('NIC','NicName')
                $nicRg = Get-CellValueAny -Row $row -Fields @('NICRG','NicRG','NetworkInterfaceRG')
                if (-not $nicRg) { $nicRg = $rgName }
                $subnetName = Get-CellValueAny -Row $row -Fields @('Subnet','SubnetName')
                $vnetName = Get-CellValueAny -Row $row -Fields @('VirtualNetwork','VNetName','VnetName')
                if ($useNic -and $useNic.ToUpperInvariant() -eq 'O' -and $nicName) {
                    Write-Info "[DryRun] NIC-NSG 연결 예정: $nicName (RG=$nicRg) -> $nsgName"
                } elseif ($subnetName -and $vnetName) {
                    Write-Info "[DryRun] Subnet-NSG 연결 예정: $vnetName/$subnetName -> $nsgName"
                }
            }
            continue
        }

        $nsg = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $rgName -ErrorAction SilentlyContinue
        if (-not $nsg) {
            if ($Context.DryRun) {
                Write-Info "[DryRun] NSG 생성 예정: $nsgName"
            } else {
                $nsg = New-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $rgName -Location 'koreacentral' -Force
                Write-Info "NSG 생성 완료: $nsgName"
            }
        } else {
            Write-Info "NSG 유지: $nsgName"
        }

        foreach ($row in $group.Group) {
            $useNic = (Get-CellValue -Row $row -Field 'UseNIC')
            $nicName = Get-CellValueAny -Row $row -Fields @('NIC','NicName')
            $nicRg = Get-CellValueAny -Row $row -Fields @('NICRG','NicRG','NetworkInterfaceRG')
            if (-not $nicRg) { $nicRg = $rgName }
            $subnetName = Get-CellValueAny -Row $row -Fields @('Subnet','SubnetName')
            $vnetName = Get-CellValueAny -Row $row -Fields @('VirtualNetwork','VNetName','VnetName')

            if ($useNic -and $useNic.ToUpperInvariant() -eq 'O' -and $nicName) {
                if ($Context.DryRun) {
                    Write-Info "[DryRun] NIC-NSG 연결 예정: $nicName (RG=$nicRg) -> $nsgName"
                } else {
                    $nic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $nicRg -ErrorAction SilentlyContinue
                    if ($nic) {
                        $nic.NetworkSecurityGroup = $nsg
                        $nic | Set-AzNetworkInterface | Out-Null
                        Write-Info "NIC-NSG 연결 완료: $nicName (RG=$nicRg) -> $nsgName"
                    } else {
                        Write-WarnLog "NIC를 찾지 못해 NSG 연결을 건너뜁니다. NIC=$nicName, RG=$nicRg, NSG=$nsgName"
                    }
                }
                continue
            }

            if ($subnetName -and $vnetName) {
                if ($Context.DryRun) {
                    Write-Info "[DryRun] Subnet-NSG 연결 예정: $vnetName/$subnetName -> $nsgName"
                } else {
                    $vnetRg = Get-CellValueAny -Row $row -Fields @('VnetRG','VNetRG','VirtualNetworkRG')
                    if (-not $vnetRg -and $vnetRgMap.ContainsKey($vnetName)) {
                        $vnetRg = $vnetRgMap[$vnetName]
                    }
                    if (-not $vnetRg) {
                        $vnetRg = $rgName
                    }

                    $vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $vnetRg -ErrorAction SilentlyContinue
                    if (-not $vnet) {
                        $allByName = @(Get-AzVirtualNetwork -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $vnetName })
                        if ($allByName.Count -eq 1) {
                            $vnet = $allByName[0]
                            Write-WarnLog "NSG 시트 RG와 VNET RG가 달라 자동 보정했습니다. VNet=$vnetName, RG=$($vnet.ResourceGroupName)"
                        } elseif ($allByName.Count -gt 1) {
                            Write-WarnLog "동일 이름 VNET이 여러 개라 서브넷 NSG 연결을 건너뜁니다. VNet=$vnetName"
                            continue
                        }
                    }

                    if ($vnet) {
                        $subnet = $vnet.Subnets | Where-Object Name -eq $subnetName
                        if ($subnet) {
                            $subnet.NetworkSecurityGroup = $nsg
                            $vnet | Set-AzVirtualNetwork | Out-Null
                            Write-Info "Subnet-NSG 연결 완료: $($vnet.Name)/$subnetName -> $nsgName"
                        } else {
                            Write-WarnLog "서브넷을 찾지 못해 NSG 연결을 건너뜁니다. VNet=$($vnet.Name), Subnet=$subnetName, NSG=$nsgName"
                        }
                    } else {
                        Write-WarnLog "VNet을 찾지 못해 NSG 연결을 건너뜁니다. VNet=$vnetName, NSG=$nsgName"
                    }
                }
            }
        }

        if ($Context.DryRun -or -not $nsg) { continue }

        $targetRules = $ruleRows | Where-Object { (Get-CellValue -Row $_ -Field 'NSGName') -eq $nsgName }
        foreach ($rr in $targetRules) {
            $ruleName = Get-CellValue -Row $rr -Field 'Name'
            $priority = Get-CellValue -Row $rr -Field 'Priority'
            if (-not $ruleName -or -not $priority) { continue }

            if ($nsg.SecurityRules | Where-Object Name -eq $ruleName) {
                continue
            }

            $ruleDescription = Get-CellValue -Row $rr -Field 'Description'
            if ([string]::IsNullOrWhiteSpace($ruleDescription)) {
                $ruleDescription = 'No description'
            }

            $params = @{
                Name                     = $ruleName
                Priority                 = [int]$priority
                Direction                = (Get-CellValue -Row $rr -Field 'Direction')
                Access                   = (Get-CellValue -Row $rr -Field 'Action')
                Protocol                 = (Convert-ToRuleValue -Value (Get-CellValue -Row $rr -Field 'Protocol'))
                Description              = $ruleDescription
                SourceAddressPrefix      = (Convert-ToRuleValue -Value (Get-CellValue -Row $rr -Field 'Src Addr'))
                SourcePortRange          = (Convert-ToRuleValue -Value (Get-CellValue -Row $rr -Field 'Src Port'))
                DestinationAddressPrefix = (Convert-ToRuleValue -Value (Get-CellValue -Row $rr -Field 'Dest Addr'))
                DestinationPortRange     = (Convert-ToRuleValue -Value (Get-CellValue -Row $rr -Field 'Dest Port'))
            }

            $nsg | Add-AzNetworkSecurityRuleConfig @params | Out-Null
            Write-Info "NSG Rule 추가: $nsgName/$ruleName"
        }

        $nsg | Set-AzNetworkSecurityGroup | Out-Null
    }

    End-Step 'Deploy-Nsgs'
}

function Resolve-DiskEncryptionSetId {
    param(
        [DeploymentContext]$Context,
        [psobject]$Row,
        [string]$FallbackRgName
    )

    $explicitId = Get-CellValueAny -Row $Row -Fields @('DiskEncryptionSetId','DesResourceId')
    if ($explicitId) { return $explicitId }

    $desName = Get-CellValueAny -Row $Row -Fields @('DESName','DiskEncryptionSetName')
    if (-not $desName) { return $null }
    $desRg = Get-CellValueAny -Row $Row -Fields @('DESRG','DesRG')
    if (-not $desRg) { $desRg = $FallbackRgName }

    if ($Context.DryRun) {
        return "/subscriptions/$($Context.SubscriptionId)/resourceGroups/$desRg/providers/Microsoft.Compute/diskEncryptionSets/$desName"
    }

    $des = Get-AzDiskEncryptionSet -ResourceGroupName $desRg -Name $desName -ErrorAction SilentlyContinue
    if (-not $des) {
        throw "Disk Encryption Set을 찾을 수 없습니다: $desRg/$desName"
    }
    return $des.Id
}

function Get-DiskEncryptionSetByResourceId {
    param([string]$ResourceId)

    if ([string]::IsNullOrWhiteSpace($ResourceId)) { return $null }
    if ($ResourceId -notmatch '/resourceGroups/([^/]+)/providers/Microsoft\.Compute/diskEncryptionSets/([^/]+)$') {
        throw "DiskEncryptionSetId 형식이 올바르지 않습니다: $ResourceId"
    }

    $rgName = $Matches[1]
    $desName = $Matches[2]
    return (Get-AzDiskEncryptionSet -ResourceGroupName $rgName -Name $desName -ErrorAction SilentlyContinue)
}

function Assert-DesDoubleEncryption {
    param(
        [DeploymentContext]$Context,
        [psobject]$Row,
        [string]$FallbackRgName
    )

    $requireDoubleEncryption = Convert-ToBoolean -Value (Get-CellValueAny -Row $Row -Fields @('UseOsDiskDoubleEncryption','EnableOsDiskDoubleEncryption')) -Default $false
    if (-not $requireDoubleEncryption) { return }

    $explicitId = Get-CellValueAny -Row $Row -Fields @('DiskEncryptionSetId','DesResourceId')
    $desName = Get-CellValueAny -Row $Row -Fields @('DESName','DiskEncryptionSetName')
    if (-not $explicitId -and -not $desName) {
        throw "UseOsDiskDoubleEncryption=TRUE 인 경우 DESName 또는 DiskEncryptionSetId가 필요합니다."
    }
    if ($Context.DryRun) { return }

    $des = $null
    if ($explicitId) {
        $des = Get-DiskEncryptionSetByResourceId -ResourceId $explicitId
    } else {
        $desRg = Get-CellValueAny -Row $Row -Fields @('DESRG','DesRG')
        if (-not $desRg) { $desRg = $FallbackRgName }
        $des = Get-AzDiskEncryptionSet -ResourceGroupName $desRg -Name $desName -ErrorAction SilentlyContinue
    }

    if (-not $des) {
        throw "이중 암호화 확인 대상 DES를 찾을 수 없습니다. DESName=$desName, DiskEncryptionSetId=$explicitId"
    }

    $encType = $null
    if ($des.PSObject.Properties['EncryptionType']) {
        $encType = [string]$des.EncryptionType
    }

    if ([string]::IsNullOrWhiteSpace($encType)) {
        Write-WarnLog "DES EncryptionType을 확인하지 못했습니다. DES=$($des.Name)"
        return
    }

    if ($encType -ne 'EncryptionAtRestWithPlatformAndCustomerKeys') {
        throw "DES 이중 암호화 설정이 필요합니다. DES=$($des.Name), EncryptionType=$encType"
    }
}

function Convert-DataDiskValue {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $text = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -in @('-','--','N/A','n/a')) { return $null }
    return $text
}

function Get-DataDiskCellValue {
    param(
        [psobject]$Row,
        [string]$Field
    )

    $value = Get-CellValue -Row $Row -Field $Field
    return (Convert-DataDiskValue -Value $value)
}

function Get-DataDiskCellValueAny {
    param(
        [psobject]$Row,
        [string[]]$Fields
    )

    foreach ($field in $Fields) {
        $value = Get-DataDiskCellValue -Row $Row -Field $field
        if ($value) { return $value }
    }
    return $null
}

function Resolve-DataDiskSku {
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) {
        throw 'Disk type is required. Use StandardSSD_LRS.'
    }

    switch ($InputValue.Trim()) {
        'StandardSSD_LRS' { return 'StandardSSD_LRS' }
        'StandardSSD' { return 'StandardSSD_LRS' }
        'PremiumSSD_LRS'  { return 'Premium_LRS' }
        'Premium'         { return 'Premium_LRS' }
        'PremiumSSD'      { return 'Premium_LRS' } # (선택) 사용자가 이렇게 입력할 수도 있으니 대비
        default { throw "Only StandardSSD_LRS or PremiumSSD_LRS is allowed. Input: $InputValue" }
    }
}

function Resolve-DataDiskHostCaching {
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) { return 'None' }

    switch ($InputValue.Trim().ToLowerInvariant()) {
        'none' { return 'None' }
        'readonly' { return 'ReadOnly' }
        'readwrite' { return 'ReadWrite' }
        default { throw "Invalid HostCaching value: $InputValue (allowed: None/ReadOnly/ReadWrite)" }
    }
}

function Resolve-DataDiskEncryptionSetId {
    param(
        [DeploymentContext]$Context,
        [psobject]$Row
    )

    $explicitId = Get-DataDiskCellValueAny -Row $Row -Fields @('DiskEncryptionSetId','DesResourceId')
    if ($explicitId) { return $explicitId }

    $desName = Get-DataDiskCellValueAny -Row $Row -Fields @('DESName','DiskEncryptionSetName')
    if (-not $desName) {
        throw 'DiskEncryptionSetId or DESName(DiskEncryptionSetName) is required.'
    }

    $desRg = Get-DataDiskCellValueAny -Row $Row -Fields @('DESRG','DesRG','DiskEncryptionSetRG')
    if (-not $desRg) {
        throw 'DESRG is required to resolve DESName.'
    }

    if ($Context.DryRun) {
        return "<DryRun: /resourceGroups/$desRg/providers/Microsoft.Compute/diskEncryptionSets/$desName>"
    }

    $des = Get-AzDiskEncryptionSet -ResourceGroupName $desRg -Name $desName -ErrorAction SilentlyContinue
    if (-not $des) {
        throw "Disk Encryption Set not found. RG=$desRg, Name=$desName"
    }

    return $des.Id
}

function New-DataDiskSpecsFromRow {
    param([psobject]$Row)

    $specs = @()

    $disk1Name = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk1Name','DataDisk1Name','Disk1Name')
    $disk1Type = Get-DataDiskCellValueAny -Row $Row -Fields @('DataDisk1Type','Datadisk1Type','Disk1Type')
    $disk1Size = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk1Size','DataDisk1Size','DataDisk1SizeGB','Disk1SizeGB')
    $disk1Double = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk1DoubleEncryption','DataDisk1DoubleEncryption','Disk1DoubleEncryption','DoubleEncryption1')
    $disk1Caching = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk1HostCaching','DataDisk1HostCaching','Disk1HostCaching')
    $disk1Lun = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk1Lun','DataDisk1Lun','Disk1Lun')

    if ($disk1Name -and $disk1Size) {
        $specs += [pscustomobject]@{
            Slot = 1
            Name = $disk1Name
            Type = $disk1Type
            Size = $disk1Size
            DoubleEncryption = $disk1Double
            HostCaching = $disk1Caching
            Lun = $disk1Lun
        }
    }

    $disk2Name = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk2Name','DataDisk2Name','Disk2Name')
    $disk2Type = Get-DataDiskCellValueAny -Row $Row -Fields @('DataDisk2Type','Datadisk2Type','Disk2Type')
    $disk2Size = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk2Size','DataDisk2Size','DataDisk2SizeGB','Disk2SizeGB')
    $disk2Double = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk2DoubleEncryption','DataDisk2DoubleEncryption','Disk2DoubleEncryption','DoubleEncryption2')
    $disk2Caching = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk2HostCaching','DataDisk2HostCaching','Disk2HostCaching')
    $disk2Lun = Get-DataDiskCellValueAny -Row $Row -Fields @('Datadisk2Lun','DataDisk2Lun','Disk2Lun')

    if ($disk2Name -and $disk2Size) {
        $specs += [pscustomobject]@{
            Slot = 2
            Name = $disk2Name
            Type = $disk2Type
            Size = $disk2Size
            DoubleEncryption = $disk2Double
            HostCaching = $disk2Caching
            Lun = $disk2Lun
        }
    }

    return @($specs)
}

function Resolve-DataDiskLun {
    param(
        [object]$Vm,
        [string]$LunInput
    )

    if (-not [string]::IsNullOrWhiteSpace($LunInput)) {
        $lun = 0
        if (-not [int]::TryParse($LunInput, [ref]$lun)) {
            throw "LUN must be numeric: $LunInput"
        }
        if ($lun -lt 0) {
            throw "LUN must be >= 0: $LunInput"
        }

        $used = @($Vm.StorageProfile.DataDisks | ForEach-Object { [int]$_.Lun })
        if ($used -contains $lun) {
            throw "LUN already in use: $lun"
        }
        return $lun
    }

    $existing = @($Vm.StorageProfile.DataDisks | ForEach-Object { [int]$_.Lun })
    for ($candidate = 0; $candidate -lt 64; $candidate++) {
        if ($existing -notcontains $candidate) { return $candidate }
    }
    throw 'No available LUN (0-63).'
}

function Resolve-VmAdminPassword {
    param(
        [DeploymentContext]$Context,
        [psobject]$Row,
        [string]$VmName
    )

    $useKvPassword = Convert-ToBoolean -Value (Get-CellValue -Row $Row -Field 'UseKeyVaultPassword') -Default $false
    $plainPassword = Get-CellValue -Row $Row -Field 'AdminPassword'
    if (-not $useKvPassword) {
        if (-not $plainPassword) {
            throw "AdminPassword가 비어 있습니다. VM=$VmName"
        }
        return $plainPassword
    }

    if ($Context.DryRun) {
        return 'DryRun!VmPassw0rd'
    }

    $secretUri = Get-CellValue -Row $Row -Field 'AdminPasswordSecretUri'
    $kvName = Get-CellValue -Row $Row -Field 'AdminPasswordKVName'
    $secretName = Get-CellValue -Row $Row -Field 'AdminPasswordSecretName'
    $secretVersion = Get-CellValue -Row $Row -Field 'AdminPasswordSecretVersion'
    if ($secretVersion -and $secretVersion.Trim().ToLowerInvariant() -eq 'latest') { $secretVersion = $null }

    $secret = $null
    if ($secretUri) {
        if ($secretUri -match '/secrets/') {
            $secret = Get-AzKeyVaultSecret -Id $secretUri -ErrorAction Stop
        } else {
            if (-not $kvName) {
                if ($secretUri -match 'https://([^./]+)\.vault\.azure\.net/?$') {
                    $kvName = $Matches[1]
                } else {
                    throw "AdminPasswordSecretUri 형식이 올바르지 않습니다. VM=$VmName"
                }
            }
            if (-not $secretName) {
                throw "AdminPasswordSecretUri가 Vault URL인 경우 AdminPasswordSecretName이 필요합니다. VM=$VmName"
            }
        }
    }

    if (-not $secret) {
        if (-not $kvName -or -not $secretName) {
            throw "UseKeyVaultPassword=Y 인 경우 AdminPasswordSecretUri 또는 AdminPasswordKVName+AdminPasswordSecretName이 필요합니다. VM=$VmName"
        }
        if ($secretVersion) {
            $secret = Get-AzKeyVaultSecret -VaultName $kvName -Name $secretName -Version $secretVersion -ErrorAction Stop
        } else {
            $secret = Get-AzKeyVaultSecret -VaultName $kvName -Name $secretName -ErrorAction Stop
        }
    }

    if (-not $secret -or -not $secret.SecretValue) {
        throw "Key Vault Secret 값을 조회하지 못했습니다. VM=$VmName"
    }

    if ($secret.PSObject.Properties.Name -contains 'SecretValueText' -and $secret.SecretValueText) {
        return [string]$secret.SecretValueText
    }

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret.SecretValue)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function New-VmTemplateParameter {
    param(
        [DeploymentContext]$Context,
        [psobject]$Row,
        [string]$VmName,
        [string]$RgName
    )

    $vmSizeRaw = Get-CellValue -Row $Row -Field 'VmSize'
    if ($vmSizeRaw.StartsWith('Standard_')) {
        $vmSize = $vmSizeRaw
    } else {
        $vmSize = "Standard_$vmSizeRaw"
    }

    $osDiskName = Get-CellValue -Row $Row -Field 'OsDiskName'
    if (-not $osDiskName) {
        $osDiskName = "$VmName-OsDisk"
    }

    $adminPassword = Resolve-VmAdminPassword -Context $Context -Row $Row -VmName $VmName
    Assert-DesDoubleEncryption -Context $Context -Row $Row -FallbackRgName $RgName
    $desId = Resolve-DiskEncryptionSetId -Context $Context -Row $Row -FallbackRgName $RgName
    $vnetRG = Get-CellValue -Row $Row -Field 'VnetRG'
    if (-not $vnetRG) {
        $vnetRG = $RgName
    }
    $enableAcceleratedNetworking = Convert-ToBoolean -Value (Get-CellValue -Row $Row -Field 'EnableAcceleratedNetworking') -Default $true
    $osType = Get-CellValue -Row $Row -Field 'OsType'
    $computerName = $VmName
    if ($osType -and $osType.Trim().ToUpperInvariant() -eq 'WINDOWS') {
        $computerName = Get-WindowsComputerName -VmName $VmName
        if ($computerName -ne $VmName) {
            Write-Info "Windows ComputerName 자동 보정: VM=$VmName -> ComputerName=$computerName"
        }
    }

    $params = @{
        location                      = (Get-CellValue -Row $Row -Field 'Location')
        networkInterfaceName          = (Get-CellValue -Row $Row -Field 'NicName')
        subnetName                    = (Get-CellValue -Row $Row -Field 'SubnetName')
        vnetRG                        = $vnetRG
        virtualNetworkName            = (Get-CellValue -Row $Row -Field 'VnetName')
        privateIP                     = (Get-CellValue -Row $Row -Field 'PrivateIP')
        virtualMachineName            = $VmName
        virtualMachineComputerName    = $computerName
        virtualMachineRG              = $RgName
        virtualMachineSize            = $vmSize
        osDiskType                    = (Convert-OsDiskStorageType -InputValue (Get-CellValue -Row $Row -Field 'OsDiskStorageType'))
        osDiskName                    = $osDiskName
        adminUsername                 = (Get-CellValue -Row $Row -Field 'AdminUsername')
        adminPassword                 = $adminPassword
        enableAcceleratedNetworking   = $enableAcceleratedNetworking
        virtualMachineZone            = (Get-CellValue -Row $Row -Field 'Zones')
        ResourceGroupName             = $RgName
    }

    if ($desId) {
        $params.diskEncryptionSetId = $desId
    }

    if (-not $params.virtualMachineZone) {
        $params.virtualMachineZone = '1'
    }

    $sourceType = Get-VmSourceType -Row $Row
    if ($sourceType -eq 'CustomImage') {
        $params.imageResourceId = (Get-CellValue -Row $Row -Field 'ImageResourceId')
    } else {
        $params.publisher = (Get-CellValue -Row $Row -Field 'Publisher')
        $params.offer = (Get-CellValue -Row $Row -Field 'Offer')
        $params.sku = (Get-CellValue -Row $Row -Field 'Sku')
        $params.version = (Get-CellValue -Row $Row -Field 'Version')
    }

    return $params
}

function Get-VmTemplateFile {
    param([psobject]$Row)

    $osType = (Get-CellValue -Row $Row -Field 'OsType')
    $sourceType = Get-VmSourceType -Row $Row

    if ($osType -eq 'Linux' -and $sourceType -eq 'Marketplace') { return '.\템플릿\VM\template-VM-LinuxZone.json' }
    if ($osType -eq 'Linux' -and $sourceType -eq 'CustomImage') { return '.\템플릿\VM\template-VM-LinuxZone_Image.json' }
    if ($osType -eq 'Windows' -and $sourceType -eq 'Marketplace') { return '.\템플릿\VM\template-VM-WindowsZone.json' }
    if ($osType -eq 'Windows' -and $sourceType -eq 'CustomImage') { return '.\템플릿\VM\template-VM-WindowsZone_Image.json' }

    throw "지원하지 않는 VM 조합입니다. OsType=$osType, SourceType=$sourceType"
}

function Get-TemplateParameterNameSet {
    param([string]$TemplateFile)

    $fullTemplatePath = if ([System.IO.Path]::IsPathRooted($TemplateFile)) {
        $TemplateFile
    } else {
        Join-Path -Path $PSScriptRoot -ChildPath $TemplateFile
    }

    if (-not (Test-Path -LiteralPath $fullTemplatePath)) {
        throw "템플릿 파일을 찾을 수 없습니다: $fullTemplatePath"
    }

    $templateJson = Get-Content -Path $fullTemplatePath -Raw | ConvertFrom-Json -Depth 20
    $set = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $templateJson.parameters.PSObject.Properties.Name) {
        [void]$set.Add($p)
    }
    return $set
}

function Filter-TemplateParameters {
    param(
        [hashtable]$Parameters,
        [string]$TemplateFile
    )

    $allowed = Get-TemplateParameterNameSet -TemplateFile $TemplateFile
    $filtered = @{}
    $dropped = New-Object System.Collections.Generic.List[string]

    foreach ($k in $Parameters.Keys) {
        if ($allowed.Contains($k)) {
            $filtered[$k] = $Parameters[$k]
        } else {
            [void]$dropped.Add($k)
        }
    }

    if ($dropped.Count -gt 0) {
        Write-WarnLog "템플릿 미정의 파라미터 제외: $($dropped -join ', ') (Template=$TemplateFile)"
    }

    return $filtered
}

function Deploy-Vms {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-Vms'
    $rows = Get-SheetRows -Context $Context -SheetCandidates @('VM','VM_PRD')
    $rows = Get-FilteredVmRows -Context $Context -Rows $rows

    foreach ($row in $rows) {
        $vmName = Get-CellValue -Row $row -Field 'Name'
        if (-not $vmName) { continue }

        $rgName = Get-CellValue -Row $row -Field 'RGname'
        $location = Get-CellValue -Row $row -Field 'Location'

        Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun

        $templateFile = Get-VmTemplateFile -Row $row
        $params = New-VmTemplateParameter -Context $Context -Row $row -VmName $vmName -RgName $rgName
        $params = Filter-TemplateParameters -Parameters $params -TemplateFile $templateFile

        if ($Context.DryRun) {
            Write-Info "[DryRun] VM 배포 예정: $vmName (Template=$templateFile)"
            continue
        }

        $existing = Get-AzVM -ResourceGroupName $rgName -Name $vmName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Info "VM 유지: $vmName"
            continue
        }

        $result = Invoke-WithRetry -Operation "New-AzResourceGroupDeployment:$vmName" -ScriptBlock {
            New-AzResourceGroupDeployment -ResourceGroupName $rgName -TemplateFile $templateFile -TemplateParameterObject $params -ErrorAction Stop
        }
        if ($result.ProvisioningState -eq 'Succeeded') {
            Write-Info "VM 배포 성공: $vmName"
        } else {
            throw "VM 배포 실패: $vmName, ProvisioningState=$($result.ProvisioningState)"
        }
    }

    End-Step 'Deploy-Vms'
}

function Deploy-DataDisks {
    param([DeploymentContext]$Context)

    Start-Step 'Deploy-DataDisks'

    $rows = Get-SheetRows -Context $Context -SheetCandidates @('VM_Datadisk')
    $rows = Get-FilteredRowsByOption -Rows $rows -OptionFilters $Context.VmRoleFilter -Columns @('Option','Role','VmRole')

    $attachToVm = $false
    $createdCount = 0
    $attachedCount = 0
    $skippedCount = 0
    $failedCount = 0
    $rowNo = 2

    foreach ($row in $rows) {
        try {
            $vmName = Get-DataDiskCellValueAny -Row $row -Fields @('VMName','VmName','VirtualMachineName','Name')
            $rgName = Get-DataDiskCellValueAny -Row $row -Fields @('RGname','ResourceGroupName','DataDiskRG','DiskRG')
            $location = Get-DataDiskCellValue -Row $row -Field 'Location'
            $zoneRaw = Get-DataDiskCellValueAny -Row $row -Fields @('Zone','Zones')
            $enableRaw = Get-DataDiskCellValueAny -Row $row -Fields @('Enable','Enabled')
            $rowDoubleRaw = Get-DataDiskCellValueAny -Row $row -Fields @('DoubleEncryption','UseDoubleEncryption','EncryptionAtRestWithPlatformAndCustomerKeys')
            $rowHostCachingRaw = Get-DataDiskCellValueAny -Row $row -Fields @('HostCaching','DataDiskHostCaching')

            $diskSpecs = @(New-DataDiskSpecsFromRow -Row $row)

            if (-not (Convert-ToBoolean -Value $enableRaw -Default $true)) {
                Write-Info "[SKIP] Row $rowNo disabled"
                $skippedCount++
                continue
            }

            if (-not $vmName -and -not $rgName -and -not $location -and $diskSpecs.Count -eq 0) {
                continue
            }

            if ($diskSpecs.Count -eq 0) {
                Write-Info "[SKIP] Row $rowNo no disk input (VM=$vmName)"
                $skippedCount++
                continue
            }

            if (-not $vmName) { throw 'VMName is required.' }
            if (-not $rgName) { throw 'RGname (or DataDiskRG/DiskRG) is required.' }
            if (-not $location) { throw 'Location is required.' }
            if (-not $zoneRaw) { throw 'Zone/Zones is required.' }
            $zone = $zoneRaw.Trim()

            $desId = Resolve-DataDiskEncryptionSetId -Context $Context -Row $row

            Ensure-ResourceGroup -Name $rgName -Location $location -DryRun:$Context.DryRun

            $vm = $null
            $vmChanged = $false
            if ($attachToVm -and -not $Context.DryRun) {
                $vm = Get-AzVM -ResourceGroupName $rgName -Name $vmName -ErrorAction SilentlyContinue
                if (-not $vm) {
                    throw "VM not found for attach. RG=$rgName, VM=$vmName"
                }
            }

            foreach ($disk in $diskSpecs) {
                if (-not $disk.Name) { throw "Datadisk$($disk.Slot)Name is required. VM=$vmName" }
                if (-not $disk.Size) { throw "Datadisk$($disk.Slot)Size is required. VM=$vmName" }

                $diskSizeGB = 0
                if (-not [int]::TryParse([string]$disk.Size, [ref]$diskSizeGB)) {
                    throw "Datadisk$($disk.Slot)Size must be numeric. VM=$vmName, value=$($disk.Size)"
                }
                if ($diskSizeGB -le 0) {
                    throw "Datadisk$($disk.Slot)Size must be >= 1. VM=$vmName, value=$($disk.Size)"
                }

                $diskSku = Resolve-DataDiskSku -InputValue ([string]$disk.Type)

                $doubleInput = if ($disk.DoubleEncryption) { [string]$disk.DoubleEncryption } else { $rowDoubleRaw }
                $doubleEnabled = Convert-ToBoolean -Value $doubleInput -Default $true
                $encryptionType = if ($doubleEnabled) {
                    'EncryptionAtRestWithPlatformAndCustomerKeys'
                } else {
                    'EncryptionAtRestWithCustomerKey'
                }

                $cachingInput = if ($disk.HostCaching) { [string]$disk.HostCaching } else { $rowHostCachingRaw }
                $hostCaching = Resolve-DataDiskHostCaching -InputValue $cachingInput

                if ($Context.DryRun) {
                    Write-Info "[DryRun] Create plan: VM=$vmName, Disk=$($disk.Name), RG=$rgName, SizeGB=$diskSizeGB, Sku=$diskSku, Zone=$zone, HostCaching=$hostCaching, DoubleEncryption=$doubleEnabled, EncryptionType=$encryptionType, DES=$desId"
                    if ($attachToVm) {
                        Write-Info "[DryRun] Attach plan: VM=$vmName, Disk=$($disk.Name), HostCaching=$hostCaching, LUN=$($disk.Lun)"
                    }
                    $createdCount++
                    continue
                }

                $createdDisk = Get-AzDisk -ResourceGroupName $rgName -DiskName $disk.Name -ErrorAction SilentlyContinue
                if (-not $createdDisk) {
                    $diskConfig = New-AzDiskConfig `
                        -Location $location `
                        -CreateOption Empty `
                        -DiskSizeGB $diskSizeGB `
                        -SkuName $diskSku `
                        -Zone $zone `
                        -DiskEncryptionSetId $desId `
                        -EncryptionType $encryptionType

                    $createdDisk = New-AzDisk -ResourceGroupName $rgName -DiskName $disk.Name -Disk $diskConfig -ErrorAction Stop
                    Write-Info "Created: $($disk.Name) (RG=$rgName, Zone=$zone, DoubleEncryption=$doubleEnabled)"
                    $createdCount++
                } else {
                    Write-Info "[SKIP] Already exists: $($disk.Name) (RG=$rgName)"
                    $skippedCount++
                }

                if ($attachToVm) {
                    $alreadyAttached = @($vm.StorageProfile.DataDisks | Where-Object { $_.Name -eq $disk.Name -or $_.ManagedDisk.Id -eq $createdDisk.Id })
                    if ($alreadyAttached.Count -gt 0) {
                        Write-Info "[SKIP] Already attached: VM=$vmName, Disk=$($disk.Name)"
                        $skippedCount++
                        continue
                    }

                    $lun = Resolve-DataDiskLun -Vm $vm -LunInput ([string]$disk.Lun)
                    $vm = Add-AzVMDataDisk -VM $vm -Name $disk.Name -ManagedDiskId $createdDisk.Id -Lun $lun -Caching $hostCaching -CreateOption Attach
                    $vmChanged = $true
                    $attachedCount++
                    Write-Info "Attach queued: VM=$vmName, Disk=$($disk.Name), LUN=$lun, HostCaching=$hostCaching"
                }
            }

            if ($attachToVm -and $vmChanged -and -not $Context.DryRun) {
                Update-AzVM -ResourceGroupName $rgName -VM $vm -ErrorAction Stop | Out-Null
                Write-Info "VM update completed: $vmName"
            }
        } catch {
            $failedCount++
            Write-WarnLog "[FAIL] Row ${rowNo}: $($_.Exception.Message)"
        } finally {
            $rowNo++
        }
    }

    Write-Info "DataDisk Summary: Created=$createdCount, Attached=$attachedCount, Skipped=$skippedCount, Failed=$failedCount"
    if ($failedCount -gt 0) {
        throw "Failed rows exist: $failedCount"
    }

    End-Step 'Deploy-DataDisks'
}

function Initialize-Modules {
    Start-Step 'Initialize-Modules'

    foreach ($module in @('ImportExcel','PSOutLog')) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Import-Module $module -ErrorAction Stop
    }

    foreach ($module in @('Az.Accounts','Az.Resources','Az.Network','Az.Storage','Az.Compute','Az.KeyVault')) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "필수 Az 모듈이 없습니다: $module (Install-Module $module 필요)"
        }
        Import-Module $module -ErrorAction Stop
    }

    End-Step 'Initialize-Modules'
}

function Resolve-ExcelFullPath {
    param(
        [string]$Path,
        [switch]$AllowFallback
    )

    $candidatePath = $Path
    if (-not [System.IO.Path]::IsPathRooted($candidatePath)) {
        $candidatePath = Join-Path -Path $PSScriptRoot -ChildPath $candidatePath
    }

    if (Test-Path -LiteralPath $candidatePath) {
        return (Resolve-Path -LiteralPath $candidatePath).Path
    }

    if (-not $AllowFallback) {
        throw "Excel 파일을 찾을 수 없습니다: $Path"
    }

    # 기본 파일명이 바뀌는 경우를 대비해 서버정보 폴더의 최신 리소스배포 파일로 보완
    $serverInfoDir = if ([System.IO.Path]::IsPathRooted($Path)) {
        Split-Path -Path $candidatePath -Parent
    } else {
        Join-Path -Path $PSScriptRoot -ChildPath '서버정보'
    }

    if (Test-Path -LiteralPath $serverInfoDir) {
        $fallback = Get-ChildItem -Path $serverInfoDir -File -Filter '*.xlsx' -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d{8}_리소스배포_.*\.xlsx$' } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($fallback) {
            Write-WarnLog "지정한 Excel 파일을 찾지 못해 최신 파일을 사용합니다: $($fallback.FullName)"
            return $fallback.FullName
        }
    }

    throw "Excel 파일을 찾을 수 없습니다: $Path"
}

function Ensure-AzSession {
    param(
        [switch]$ConnectAccount,
        [switch]$DryRun
    )

    Start-Step 'Ensure-AzSession'

    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx -and $ConnectAccount) {
        Connect-AzAccount -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext -ErrorAction Stop
    }

    # Context가 있어도 토큰 만료로 실제 호출이 실패할 수 있어 유효성을 확인한다.
    if ($ctx) {
        $sessionValid = $true
        try {
            Get-AzSubscription -SubscriptionId $ctx.Subscription.Id -ErrorAction Stop | Out-Null
        } catch {
            $sessionValid = $false
        }

        if (-not $sessionValid -and $ConnectAccount) {
            Write-WarnLog 'Azure 세션이 만료되어 재로그인을 시도합니다.'
            Connect-AzAccount -ErrorAction Stop | Out-Null
            $ctx = Get-AzContext -ErrorAction Stop
        } elseif (-not $sessionValid -and -not $DryRun) {
            throw 'Azure 인증이 만료되었습니다. -ConnectAccount 옵션으로 재로그인해 주세요.'
        }
    }

    if (-not $ctx -and -not $DryRun) {
        throw 'Azure 세션이 없습니다. -ConnectAccount 옵션을 사용하거나 사전에 로그인해 주세요.'
    }

    if (-not $ctx -and $DryRun) {
        End-Step 'Ensure-AzSession'
        return '00000000-0000-0000-0000-000000000000'
    }

    End-Step 'Ensure-AzSession'
    return $ctx.Subscription.Id
}

function Run-Main {
    Start-Step 'Script'

    Initialize-Modules

    $allowFallback = -not $script:IsExcelPathExplicit
    $resolvedExcelPath = Resolve-ExcelFullPath -Path $ExcelPath -AllowFallback:$allowFallback
    $subscriptionId = Ensure-AzSession -ConnectAccount:$ConnectAccount -DryRun:$DryRun

    $context = [DeploymentContext]::new($resolvedExcelPath, $DeployType, $Option, [bool]$DryRun, $subscriptionId, [bool]$LbResourceOnly)

    $issues = @(Validate-Inputs -Context $context)
    $blockingIssues = @($issues | Where-Object { $_.Message -notlike '권장 필드*' })

    if ($issues.Count -gt 0) {
        foreach ($issue in $issues) {
            if ($issue.Message -like '권장 필드*') {
                Write-WarnLog "[$($issue.Type)] Row=$($issue.Row), Name=$($issue.ResourceName), Field=$($issue.Field): $($issue.Message)"
            } else {
                Write-ErrorLog "[$($issue.Type)] Row=$($issue.Row), Name=$($issue.ResourceName), Field=$($issue.Field): $($issue.Message)"
            }
        }
    }

    if ($blockingIssues.Count -gt 0) {
        throw '입력 데이터 검증에 실패했습니다. Excel 데이터를 수정한 후 다시 실행해 주세요.'
    }

    $deploymentOrder = @('RG','VNET','UDR','LB','LB_PROBE','LB_RULE','VM','DATADISK','NSG')
    $lbDeployHandled = $false
    foreach ($type in $deploymentOrder) {
        if ($context.DeployType -notcontains $type) { continue }
        switch ($type) {
            'RG' { Deploy-ResourceGroups -Context $context }
            'VNET' { Deploy-Vnets -Context $context }
            'UDR' { Deploy-Udrs -Context $context }
            'STORAGE' { Deploy-Storages -Context $context }
            'KV' { Deploy-KeyVaults -Context $context }
            'DES' { Deploy-DiskEncryptionSets -Context $context }
            'LB' {
                if (-not $lbDeployHandled) {
                    Deploy-LoadBalancers -Context $context
                    $lbDeployHandled = $true
                }
            }
            'LB_PROBE' {
                if (-not $lbDeployHandled) {
                    Deploy-LoadBalancers -Context $context
                    $lbDeployHandled = $true
                }
            }
            'LB_RULE' {
                if (-not $lbDeployHandled) {
                    Deploy-LoadBalancers -Context $context
                    $lbDeployHandled = $true
                }
            }
            'NSG' { Deploy-Nsgs -Context $context }
            'VM' { Deploy-Vms -Context $context }
            'DATADISK' { Deploy-DataDisks -Context $context }
        }
    }

    End-Step 'Script'
}

try {
    Run-Main
} catch {
    Write-ErrorLog "치명적 오류: $($_.Exception.Message)"
    throw
}
