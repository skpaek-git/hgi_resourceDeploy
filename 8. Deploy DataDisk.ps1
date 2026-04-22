[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExcelPath = '.\서버정보\20260422_리소스배포_종합.xlsx',

    [Parameter()]
    [string]$WorksheetName = 'VM_Datadisk',

    [Parameter()]
    [Alias('VmRole','Role')]
    [string[]]$Option = @(),

    [Parameter()]
    [switch]$ConnectAccount,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [switch]$AttachToVm
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function Write-WarnLog {
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

    $text = $prop.Value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -in @('-','--','N/A','n/a')) { return $null }
    return $text
}

function Get-CellValueAny {
    param(
        [psobject]$Row,
        [string[]]$Fields
    )

    foreach ($field in $Fields) {
        $v = Get-CellValue -Row $Row -Field $field
        if ($v) { return $v }
    }
    return $null
}

function Convert-ToBoolean {
    param(
        [string]$Value,
        [bool]$Default = $true
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    $raw = $Value.Trim().ToLowerInvariant()
    if ($raw -in @('1','true','t','yes','y','o')) { return $true }
    if ($raw -in @('0','false','f','no','n','x')) { return $false }
    return $Default
}

function Resolve-DiskSku {
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) {
        throw 'Disk type is required. Use StandardSSD_LRS.'
    }

    switch ($InputValue.Trim()) {
        'StandardSSD_LRS' { return 'StandardSSD_LRS' }
        'StandardSSD' { return 'StandardSSD_LRS' }
        default { throw "Only StandardSSD_LRS is allowed. Input: $InputValue" }
    }
}

function Resolve-HostCaching {
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) {
        return 'None'
    }

    switch ($InputValue.Trim().ToLowerInvariant()) {
        'none' { return 'None' }
        'readonly' { return 'ReadOnly' }
        'readwrite' { return 'ReadWrite' }
        default { throw "Invalid HostCaching value: $InputValue (allowed: None/ReadOnly/ReadWrite)" }
    }
}

function Resolve-WorksheetName {
    param(
        [string]$Path,
        [string]$Requested
    )

    $sheetInfos = Get-ExcelSheetInfo -Path $Path
    $exact = $sheetInfos | Where-Object { $_.Name -eq $Requested } | Select-Object -First 1
    if ($exact) { return $exact.Name }

    $wild = $sheetInfos | Where-Object { $_.Name -like "$Requested*" } | Select-Object -First 1
    if ($wild) { return $wild.Name }

    $allNames = $sheetInfos.Name -join ', '
    throw "Sheet not found: $Requested (available: $allNames)"
}

function Get-FilteredRowsByOption {
    param(
        [object[]]$Rows,
        [string[]]$FilterValues
    )

    if ($FilterValues.Count -eq 0) { return @($Rows) }

    $normalized = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $FilterValues) {
        if ([string]::IsNullOrWhiteSpace($f)) { continue }
        [void]$normalized.Add($f.Trim())
    }
    if ($normalized.Count -eq 0) { return @($Rows) }

    $columns = @('Option','Role','VmRole')
    $availableColumns = @()
    if ($Rows.Count -gt 0) {
        $names = @($Rows[0].PSObject.Properties.Name)
        foreach ($c in $columns) {
            if ($names -contains $c) { $availableColumns += $c }
        }
    }

    if ($availableColumns.Count -eq 0) {
        throw "Datadisk 시트에 'Option'(또는 Role/VmRole) 컬럼이 없습니다. -Option 필터를 사용하려면 컬럼이 필요합니다."
    }

    $filtered = foreach ($r in $Rows) {
        $rowValue = $null
        foreach ($col in $availableColumns) {
            $rowValue = Get-CellValue -Row $r -Field $col
            if ($rowValue) { break }
        }

        if ($rowValue -and $normalized.Contains($rowValue)) {
            $r
        }
    }

    return @($filtered)
}

function Resolve-DiskEncryptionSetId {
    param(
        [psobject]$Row,
        [bool]$IsDryRun
    )

    $explicitId = Get-CellValueAny -Row $Row -Fields @('DiskEncryptionSetId','DesResourceId')
    if ($explicitId) { return $explicitId }

    $desName = Get-CellValueAny -Row $Row -Fields @('DESName','DiskEncryptionSetName')
    if (-not $desName) {
        throw 'DiskEncryptionSetId or DESName(DiskEncryptionSetName) is required.'
    }

    $desRg = Get-CellValueAny -Row $Row -Fields @('DESRG','DesRG','DiskEncryptionSetRG')
    if (-not $desRg) {
        throw 'DESRG is required to resolve DESName.'
    }

    if ($IsDryRun) {
        return "<DryRun: /resourceGroups/$desRg/providers/Microsoft.Compute/diskEncryptionSets/$desName>"
    }

    $des = Get-AzDiskEncryptionSet -ResourceGroupName $desRg -Name $desName -ErrorAction SilentlyContinue
    if (-not $des) {
        throw "Disk Encryption Set not found. RG=$desRg, Name=$desName"
    }

    return $des.Id
}

function Ensure-AzSession {
    param([switch]$ShouldConnect)

    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx -and $ShouldConnect) {
        Connect-AzAccount -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext -ErrorAction Stop
    }

    if (-not $ctx) {
        throw 'Azure session not found. Use -ConnectAccount.'
    }

    Get-AzSubscription -SubscriptionId $ctx.Subscription.Id -ErrorAction Stop | Out-Null
}

function Ensure-ResourceGroup {
    param(
        [string]$Name,
        [string]$Location,
        [switch]$IsDryRun
    )

    if ($IsDryRun) {
        Write-Info "[DryRun] RG check/create: $Name ($Location)"
        return
    }

    $rg = Get-AzResourceGroup -Name $Name -ErrorAction SilentlyContinue
    if ($rg) {
        Write-Info "RG exists: $Name"
        return
    }

    New-AzResourceGroup -Name $Name -Location $Location -ErrorAction Stop | Out-Null
    Write-Info "RG created: $Name"
}

function New-DiskSpecsFromRow {
    param([psobject]$Row)

    $specs = @()

    $disk1Name = Get-CellValueAny -Row $Row -Fields @('Datadisk1Name','DataDisk1Name','Disk1Name')
    $disk1Type = Get-CellValueAny -Row $Row -Fields @('DataDisk1Type','Datadisk1Type','Disk1Type')
    $disk1Size = Get-CellValueAny -Row $Row -Fields @('Datadisk1Size','DataDisk1Size','DataDisk1SizeGB','Disk1SizeGB')
    $disk1Double = Get-CellValueAny -Row $Row -Fields @('Datadisk1DoubleEncryption','DataDisk1DoubleEncryption','Disk1DoubleEncryption','DoubleEncryption1')
    $disk1Caching = Get-CellValueAny -Row $Row -Fields @('Datadisk1HostCaching','DataDisk1HostCaching','Disk1HostCaching')
    $disk1Lun = Get-CellValueAny -Row $Row -Fields @('Datadisk1Lun','DataDisk1Lun','Disk1Lun')

    if ($disk1Name -or $disk1Type -or $disk1Size) {
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

    $disk2Name = Get-CellValueAny -Row $Row -Fields @('Datadisk2Name','DataDisk2Name','Disk2Name')
    $disk2Type = Get-CellValueAny -Row $Row -Fields @('DataDisk2Type','Datadisk2Type','Disk2Type')
    $disk2Size = Get-CellValueAny -Row $Row -Fields @('Datadisk2Size','DataDisk2Size','DataDisk2SizeGB','Disk2SizeGB')
    $disk2Double = Get-CellValueAny -Row $Row -Fields @('Datadisk2DoubleEncryption','DataDisk2DoubleEncryption','Disk2DoubleEncryption','DoubleEncryption2')
    $disk2Caching = Get-CellValueAny -Row $Row -Fields @('Datadisk2HostCaching','DataDisk2HostCaching','Disk2HostCaching')
    $disk2Lun = Get-CellValueAny -Row $Row -Fields @('Datadisk2Lun','DataDisk2Lun','Disk2Lun')

    if ($disk2Name -or $disk2Type -or $disk2Size) {
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

function Resolve-Lun {
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

Import-Module ImportExcel -ErrorAction Stop

if (-not $DryRun) {
    foreach ($module in @('Az.Accounts','Az.Resources','Az.Compute')) {
        Import-Module $module -ErrorAction Stop
    }
    Ensure-AzSession -ShouldConnect:$ConnectAccount
}

$excelCandidatePath = $ExcelPath
if (-not [System.IO.Path]::IsPathRooted($excelCandidatePath)) {
    $excelCandidatePath = Join-Path -Path $PSScriptRoot -ChildPath $excelCandidatePath
}
if (-not (Test-Path -LiteralPath $excelCandidatePath)) {
    $serverInfoDir = Join-Path -Path $PSScriptRoot -ChildPath '서버정보'
    $fallback = Get-ChildItem -Path $serverInfoDir -File -Filter '*.xlsx' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{8}_리소스배포_.*\.xlsx$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($fallback) {
        Write-WarnLog "Excel file not found: $excelCandidatePath"
        Write-WarnLog "Fallback Excel selected: $($fallback.FullName)"
        $excelCandidatePath = $fallback.FullName
    } else {
        throw "Excel file not found: $excelCandidatePath"
    }
}
$resolvedExcelPath = [string](Resolve-Path -LiteralPath $excelCandidatePath)

$sheetName = Resolve-WorksheetName -Path $resolvedExcelPath -Requested $WorksheetName
$rows = @(Import-Excel -Path $resolvedExcelPath -WorksheetName $sheetName)
$rows = @(Get-FilteredRowsByOption -Rows $rows -FilterValues $Option)

if ($Option.Count -gt 0) {
    Write-Info "Sheet loaded: $sheetName / rows: $($rows.Count) (Option filter: $($Option -join ', '))"
} else {
    Write-Info "Sheet loaded: $sheetName / rows: $($rows.Count)"
}

$createdCount = 0
$attachedCount = 0
$skippedCount = 0
$failedCount = 0
$rowNo = 2

foreach ($row in $rows) {
    try {
        $vmName = Get-CellValueAny -Row $row -Fields @('VMName','VmName','VirtualMachineName','Name')
        $rgName = Get-CellValueAny -Row $row -Fields @('RGname','ResourceGroupName','DataDiskRG','DiskRG')
        $location = Get-CellValue -Row $row -Field 'Location'
        $zoneRaw = Get-CellValueAny -Row $row -Fields @('Zone','Zones')
        $enableRaw = Get-CellValueAny -Row $row -Fields @('Enable','Enabled')
        $rowDoubleRaw = Get-CellValueAny -Row $row -Fields @('DoubleEncryption','UseDoubleEncryption','EncryptionAtRestWithPlatformAndCustomerKeys')
        $rowHostCachingRaw = Get-CellValueAny -Row $row -Fields @('HostCaching','DataDiskHostCaching')

        $diskSpecs = @(New-DiskSpecsFromRow -Row $row)

        if (-not (Convert-ToBoolean -Value $enableRaw -Default $true)) {
            Write-Info "[SKIP] Row $rowNo disabled"
            $skippedCount++
            $rowNo++
            continue
        }

        if (-not $vmName -and -not $rgName -and -not $location -and $diskSpecs.Count -eq 0) {
            $rowNo++
            continue
        }

        if ($diskSpecs.Count -eq 0) {
            Write-Info "[SKIP] Row $rowNo no disk input (VM=$vmName)"
            $skippedCount++
            $rowNo++
            continue
        }

        if (-not $vmName) { throw 'VMName is required.' }
        if (-not $rgName) { throw 'RGname (or DataDiskRG/DiskRG) is required.' }
        if (-not $location) { throw 'Location is required.' }
        if (-not $zoneRaw) { throw 'Zone/Zones is required.' }
        $zone = $zoneRaw.Trim()

        $desId = Resolve-DiskEncryptionSetId -Row $row -IsDryRun:$DryRun

        Ensure-ResourceGroup -Name $rgName -Location $location -IsDryRun:$DryRun

        $vm = $null
        $vmChanged = $false
        if ($AttachToVm -and -not $DryRun) {
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

            $diskSku = Resolve-DiskSku -InputValue ([string]$disk.Type)

            $doubleInput = if ($disk.DoubleEncryption) { [string]$disk.DoubleEncryption } else { $rowDoubleRaw }
            $doubleEnabled = Convert-ToBoolean -Value $doubleInput -Default $true
            $encryptionType = if ($doubleEnabled) {
                'EncryptionAtRestWithPlatformAndCustomerKeys'
            } else {
                'EncryptionAtRestWithCustomerKey'
            }

            $cachingInput = if ($disk.HostCaching) { [string]$disk.HostCaching } else { $rowHostCachingRaw }
            $hostCaching = Resolve-HostCaching -InputValue $cachingInput

            if ($DryRun) {
                Write-Info "[DryRun] Create plan: VM=$vmName, Disk=$($disk.Name), RG=$rgName, SizeGB=$diskSizeGB, Sku=$diskSku, Zone=$zone, HostCaching=$hostCaching, DoubleEncryption=$doubleEnabled, EncryptionType=$encryptionType, DES=$desId"
                if ($AttachToVm) {
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

            if ($AttachToVm) {
                $alreadyAttached = @($vm.StorageProfile.DataDisks | Where-Object { $_.Name -eq $disk.Name -or $_.ManagedDisk.Id -eq $createdDisk.Id })
                if ($alreadyAttached.Count -gt 0) {
                    Write-Info "[SKIP] Already attached: VM=$vmName, Disk=$($disk.Name)"
                    $skippedCount++
                    continue
                }

                $lun = Resolve-Lun -Vm $vm -LunInput ([string]$disk.Lun)
                $vm = Add-AzVMDataDisk -VM $vm -Name $disk.Name -ManagedDiskId $createdDisk.Id -Lun $lun -Caching $hostCaching -CreateOption Attach
                $vmChanged = $true
                $attachedCount++
                Write-Info "Attach queued: VM=$vmName, Disk=$($disk.Name), LUN=$lun, HostCaching=$hostCaching"
            }
        }

        if ($AttachToVm -and $vmChanged -and -not $DryRun) {
            Update-AzVM -ResourceGroupName $rgName -VM $vm -ErrorAction Stop | Out-Null
            Write-Info "VM update completed: $vmName"
        }
    }
    catch {
        $failedCount++
        Write-WarnLog "[FAIL] Row ${rowNo}: $($_.Exception.Message)"
    }
    finally {
        $rowNo++
    }
}

Write-Host ''
Write-Host '========== Summary ==========' -ForegroundColor Cyan
Write-Host "Created  : $createdCount"
Write-Host "Attached : $attachedCount"
Write-Host "Skipped  : $skippedCount"
Write-Host "Failed   : $failedCount"

if ($failedCount -gt 0) {
    throw "Failed rows exist: $failedCount"
}
