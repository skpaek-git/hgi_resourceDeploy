#region | ★★ CustomImage(Linux) ★★ |

# Login
Connect-AzAccount

# Module
Install-Module ImportExcel
Import-Module ImportExcel

# Import Excel
$xls = Import-Excel ".\서버정보\20260422_리소스배포_종합.xlsx" -WorksheetName "VM" # CSV 파일 체크 필요

# Linux VM 배포
$VMs = $xls | Where-Object {$_.OsType -eq "Linux"}
$excelVmNames = $VMs | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } | ForEach-Object { $_.Name.Trim() }

Write-Host "`n[프로세스 시작] 인프라 동기화 작업을 시작합니다." -ForegroundColor White -BackgroundColor DarkBlue

foreach ($VM in $VMs) {
    if ([string]::IsNullOrWhiteSpace($VM.Stage) -and [string]::IsNullOrWhiteSpace($VM.Role)) { continue }
    if ([string]::IsNullOrWhiteSpace($VM.Name)) { continue }

    $vmName = $VM.Name.Trim()
    $rgName = $VM.RGName.Trim()
    
    Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
    Write-Host "검사 대상: $vmName" -ForegroundColor Cyan

    # 1. 기존 VM 존재 여부 확인
    $existingVM = Get-AzVM -ResourceGroupName $rgName -Name $vmName -ErrorAction SilentlyContinue
    if ($existingVM) {
        Write-Host "-> [SKIP] $vmName 리소스가 이미 존재하여 배포를 건너뜁니다." -ForegroundColor Yellow
        continue 
    }

    # 2. 리소스 그룹 확인
    New-AzResourceGroup -Name $rgName -Location $VM.location.Trim() -Force -ErrorAction SilentlyContinue | Out-Null

    # 3. 파라미터 구성 (데이터 디스크 항목 완전 제거)
    try {
        $paramObj = @{}
        $paramObj["location"] = $VM.location.Trim()
        $paramObj["networkInterfaceName"] = $VM.NicName.Trim()
        $paramObj["subnetName"] = $VM.SubnetName.Trim()
        $paramObj["vnetRG"] = $VM.VnetRG.Trim()
        $paramObj["virtualNetworkName"] = $VM.VnetName.Trim()
        $paramObj["privateIP"] = $VM.PrivateIP.Trim()
        $paramObj["virtualMachineName"] = $vmName
        $paramObj["virtualMachineComputerName"] = $vmName
        $paramObj["virtualMachineRG"] = $rgName
        $paramObj["virtualMachineSize"] = "Standard_" + $VM.VmSize.Trim()
        $paramObj["osDiskType"] = $VM.OsDiskStorageType.Trim()
        $paramObj["osDiskName"] = $VM.OsDiskName.Trim()
        $paramObj["imageResourceId"] = $VM.ImageResourceId.Trim()
        $paramObj["adminUsername"] = $VM.AdminUsername.Trim()
        $paramObj["adminPassword"] = (New-Object -TypeName PSCredential -ArgumentList 'id', ($VM.AdminPassword.ToString().Trim() | ConvertTo-SecureString -AsPlainText -Force)).Password
        
        $subId = (Get-AzContext).Subscription.Id
        $desRg = if (-not [string]::IsNullOrWhiteSpace($VM.DesRG)) { $VM.DesRG.Trim() } else { $rgName }
        if (-not [string]::IsNullOrWhiteSpace($VM.DiskEncryptionSetId)) {
            $paramObj["diskEncryptionSetId"] = $VM.DiskEncryptionSetId.Trim()
        } elseif (-not [string]::IsNullOrWhiteSpace($VM.DESName)) {
            $paramObj["diskEncryptionSetId"] = "/subscriptions/$subId/resourceGroups/$desRg/providers/Microsoft.Compute/diskEncryptionSets/$($VM.DESName.Trim())"
        } else {
            throw "DESName 또는 DiskEncryptionSetId 값이 필요합니다. VM=$vmName"
        }
        $paramObj["diagnosticsStorageAccountName"] = $VM.DiagStrName.Trim()
        $paramObj["diagnosticsStorageAccountId"]   = "/subscriptions/$subId/resourceGroups/$($VM.DiagStrRG.Trim())/providers/Microsoft.Storage/storageAccounts/$($VM.DiagStrName.Trim())"
        $paramObj["virtualMachineZone"] = [string]$VM.Zones
        $paramObj["ResourceGroupName"] = $rgName

        # 4. ARM 배포 실행
        $deployResult = New-AzResourceGroupDeployment -ResourceGroupName $rgName `
            -TemplateFile ".\템플릿\VM\template-VM-LinuxZone_image.json" `
            -TemplateParameterObject $paramObj -ErrorAction Stop

        if ($deployResult.ProvisioningState -eq "Succeeded") {
            Write-Host "-> [SUCCESS] $vmName 배포 성공!" -ForegroundColor Green
        }
    } catch {
        Write-Host "-> [ERROR] $vmName 처리 중 오류: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
Write-Host "[삭제 검토] 엑셀 시트에서 제외된 리소스를 확인합니다." -ForegroundColor Yellow

$targetRGs = $VMs.RGName | Where-Object { $_ } | Select-Object -Unique
foreach ($rg in $targetRGs) {
    $azureVMs = Get-AzVM -ResourceGroupName $rg.Trim() -ErrorAction SilentlyContinue
    if ($null -ne $azureVMs) {
        foreach ($azVM in $azureVMs) {
            if ($excelVmNames -notcontains $azVM.Name) {
                Write-Host "`n[!] 감지됨: 엑셀에서 삭제된 VM ($($azVM.Name))" -ForegroundColor Magenta
                $confirmation = Read-Host "해당 VM과 관련 리소스(OS디스크, NIC)를 삭제하시겠습니까? (y/n)"
                
                if ($confirmation -eq 'y') {
                    $osDiskName = $azVM.StorageProfile.OsDisk.Name
                    $nicIds = $azVM.NetworkProfile.NetworkInterfaces.Id

                    # 순차 삭제 (VM -> Disk -> NIC)
                    Write-Host "-> VM 삭제 중..." -ForegroundColor Red
                    Remove-AzVM -ResourceGroupName $rg.Trim() -Name $azVM.Name -Force

                    Write-Host "-> OS 디스크($osDiskName) 삭제 중..." -ForegroundColor Red
                    Remove-AzDisk -ResourceGroupName $rg.Trim() -DiskName $osDiskName -Force

                    foreach ($nicId in $nicIds) {
                        $nicName = ($nicId -split '/')[-1]
                        Write-Host "-> NIC($nicName) 삭제 중..." -ForegroundColor Red
                        Remove-AzNetworkInterface -ResourceGroupName $rg.Trim() -Name $nicName -Force
                    }
                    Write-Host "-> 정리 완료." -ForegroundColor Gray
                }
            }
        }
    }
}

Write-Host "`n[전체 종료] 인프라 동기화가 완료되었습니다." -ForegroundColor White -BackgroundColor DarkBlue
#endregion

####################################################################################################################################################################
####################################################################################################################################################################
####################################################################################################################################################################
####################################################################################################################################################################

#region | ★★ CustomImage(Windows) ★★ |

# Windows VM 배포
$VMs = $xls | Where-Object {$_.OsType -eq "Windows"}
$excelVmNames = $VMs | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } | ForEach-Object { $_.Name.Trim() }

Write-Host "`n[프로세스 시작] 인프라 동기화 작업을 시작합니다." -ForegroundColor White -BackgroundColor DarkBlue

foreach ($VM in $VMs) {
    if ([string]::IsNullOrWhiteSpace($VM.Stage) -and [string]::IsNullOrWhiteSpace($VM.Role)) { continue }
    if ([string]::IsNullOrWhiteSpace($VM.Name)) { continue }

    $vmName = $VM.Name.Trim()
    $rgName = $VM.RGName.Trim()
    
    Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
    Write-Host "검사 대상: $vmName" -ForegroundColor Cyan

    # 1. 기존 VM 존재 여부 확인
    $existingVM = Get-AzVM -ResourceGroupName $rgName -Name $vmName -ErrorAction SilentlyContinue
    if ($existingVM) {
        Write-Host "-> [SKIP] $vmName 리소스가 이미 존재하여 배포를 건너뜁니다." -ForegroundColor Yellow
        continue 
    }

    # 2. 리소스 그룹 확인
    New-AzResourceGroup -Name $rgName -Location $VM.location.Trim() -Force -ErrorAction SilentlyContinue | Out-Null

    # 3. 파라미터 구성 (데이터 디스크 항목 완전 제거)
    try {
        $paramObj = @{}
        $paramObj["location"] = $VM.location.Trim()
        $paramObj["networkInterfaceName"] = $VM.NicName.Trim()
        $paramObj["subnetName"] = $VM.SubnetName.Trim()
        $paramObj["vnetRG"] = $VM.VnetRG.Trim()
        $paramObj["virtualNetworkName"] = $VM.VnetName.Trim()
        $paramObj["privateIP"] = $VM.PrivateIP.Trim()
        $paramObj["virtualMachineName"] = $vmName
        $paramObj["virtualMachineComputerName"] = $vmName
        $paramObj["virtualMachineRG"] = $rgName
        $paramObj["virtualMachineSize"] = "Standard_" + $VM.VmSize.Trim()
        $paramObj["osDiskType"] = $VM.OsDiskStorageType.Trim()
        $paramObj["osDiskName"] = $VM.OsDiskName.Trim()
        $paramObj["imageResourceId"] = $VM.ImageResourceId.Trim()
        $paramObj["adminUsername"] = $VM.AdminUsername.Trim()
        $paramObj["adminPassword"] = (New-Object -TypeName PSCredential -ArgumentList 'id', ($VM.AdminPassword.ToString().Trim() | ConvertTo-SecureString -AsPlainText -Force)).Password
        
        $subId = (Get-AzContext).Subscription.Id
        $desRg = if (-not [string]::IsNullOrWhiteSpace($VM.DesRG)) { $VM.DesRG.Trim() } else { $rgName }
        if (-not [string]::IsNullOrWhiteSpace($VM.DiskEncryptionSetId)) {
            $paramObj["diskEncryptionSetId"] = $VM.DiskEncryptionSetId.Trim()
        } elseif (-not [string]::IsNullOrWhiteSpace($VM.DESName)) {
            $paramObj["diskEncryptionSetId"] = "/subscriptions/$subId/resourceGroups/$desRg/providers/Microsoft.Compute/diskEncryptionSets/$($VM.DESName.Trim())"
        } else {
            throw "DESName 또는 DiskEncryptionSetId 값이 필요합니다. VM=$vmName"
        }
        $paramObj["diagnosticsStorageAccountName"] = $VM.DiagStrName.Trim()
        $paramObj["diagnosticsStorageAccountId"]   = "/subscriptions/$subId/resourceGroups/$($VM.DiagStrRG.Trim())/providers/Microsoft.Storage/storageAccounts/$($VM.DiagStrName.Trim())"
        $paramObj["virtualMachineZone"] = [string]$VM.Zones
        $paramObj["ResourceGroupName"] = $rgName

        # 4. ARM 배포 실행
        $deployResult = New-AzResourceGroupDeployment -ResourceGroupName $rgName `
            -TemplateFile ".\템플릿\VM\template-VM-WindowsZone_image.json" `
            -TemplateParameterObject $paramObj -ErrorAction Stop

        if ($deployResult.ProvisioningState -eq "Succeeded") {
            Write-Host "-> [SUCCESS] $vmName 배포 성공!" -ForegroundColor Green
        }
    } catch {
        Write-Host "-> [ERROR] $vmName 처리 중 오류: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
Write-Host "[삭제 검토] 엑셀 시트에서 제외된 리소스를 확인합니다." -ForegroundColor Yellow

$targetRGs = $VMs.RGName | Where-Object { $_ } | Select-Object -Unique
foreach ($rg in $targetRGs) {
    $azureVMs = Get-AzVM -ResourceGroupName $rg.Trim() -ErrorAction SilentlyContinue
    if ($null -ne $azureVMs) {
        foreach ($azVM in $azureVMs) {
            if ($excelVmNames -notcontains $azVM.Name) {
                Write-Host "`n[!] 감지됨: 엑셀에서 삭제된 VM ($($azVM.Name))" -ForegroundColor Magenta
                $confirmation = Read-Host "해당 VM과 관련 리소스(OS디스크, NIC)를 삭제하시겠습니까? (y/n)"
                
                if ($confirmation -eq 'y') {
                    $osDiskName = $azVM.StorageProfile.OsDisk.Name
                    $nicIds = $azVM.NetworkProfile.NetworkInterfaces.Id

                    # 순차 삭제 (VM -> Disk -> NIC)
                    Write-Host "-> VM 삭제 중..." -ForegroundColor Red
                    Remove-AzVM -ResourceGroupName $rg.Trim() -Name $azVM.Name -Force

                    Write-Host "-> OS 디스크($osDiskName) 삭제 중..." -ForegroundColor Red
                    Remove-AzDisk -ResourceGroupName $rg.Trim() -DiskName $osDiskName -Force

                    foreach ($nicId in $nicIds) {
                        $nicName = ($nicId -split '/')[-1]
                        Write-Host "-> NIC($nicName) 삭제 중..." -ForegroundColor Red
                        Remove-AzNetworkInterface -ResourceGroupName $rg.Trim() -Name $nicName -Force
                    }
                    Write-Host "-> 정리 완료." -ForegroundColor Gray
                }
            }
        }
    }
}

Write-Host "`n[전체 종료] 인프라 동기화가 완료되었습니다." -ForegroundColor White -BackgroundColor DarkBlue

#endregion

####################################################################################################################################################################
####################################################################################################################################################################
####################################################################################################################################################################
####################################################################################################################################################################

#region | Marketplace(Linux) |
<#
# Login
Connect-AzAccount

# Module
Install-Module ImportExcel
Import-Module ImportExcel
$xls = Import-Excel ".\서버정보\20260318_샘플_리소스배포_CustomImageVM_NSG_RG_VNET_LB.xlsx" -WorksheetName "VM" # CSV 파일 체크 필요
$VMs = $xls
$excelVmNames = $VMs | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } | ForEach-Object { $_.Name.Trim() }

Write-Host "`n[프로세스 시작] 인프라 동기화 작업을 시작합니다." -ForegroundColor White -BackgroundColor DarkBlue

foreach ($VM in $VMs) {
    if ([string]::IsNullOrWhiteSpace($VM.Stage) -and [string]::IsNullOrWhiteSpace($VM.Role)) { continue }
    if ([string]::IsNullOrWhiteSpace($VM.Name)) { continue }

    $vmName = $VM.Name.Trim()
    $rgName = $VM.RGName.Trim()
    
    Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
    Write-Host "검사 대상: $vmName" -ForegroundColor Cyan

    # 1. 기존 VM 존재 여부 확인
    $existingVM = Get-AzVM -ResourceGroupName $rgName -Name $vmName -ErrorAction SilentlyContinue
    if ($existingVM) {
        Write-Host "-> [SKIP] $vmName 리소스가 이미 존재하여 배포를 건너뜁니다." -ForegroundColor Yellow
        continue 
    }

    # 2. 리소스 그룹 확인
    New-AzResourceGroup -Name $rgName -Location $VM.location.Trim() -Force -ErrorAction SilentlyContinue | Out-Null

    # 3. 파라미터 구성 (데이터 디스크 항목 완전 제거)
    try {
        $paramObj = @{}
        $paramObj["location"] = $VM.location.Trim()
        $paramObj["networkInterfaceName"] = $VM.NicName.Trim()
        $paramObj["subnetName"] = $VM.SubnetName.Trim()
        $paramObj["vnetRG"] = $VM.VnetRG.Trim()
        $paramObj["virtualNetworkName"] = $VM.VnetName.Trim()
        $paramObj["privateIP"] = $VM.PrivateIP.Trim()
        $paramObj["virtualMachineName"] = $vmName
        $paramObj["virtualMachineComputerName"] = $vmName
        $paramObj["virtualMachineRG"] = $rgName
        $paramObj["virtualMachineSize"] = "Standard_" + $VM.VmSize.Trim()
        $paramObj["osDiskType"] = $VM.OsDiskStorageType.Trim()
        $paramObj["osDiskName"] = $VM.OsDiskName.Trim()
        $paramObj["imageResourceId"] = $VM.ImageResourceId.Trim()
        $paramObj["adminUsername"] = $VM.AdminUsername.Trim()
        $paramObj["adminPassword"] = (New-Object -TypeName PSCredential -ArgumentList 'id', ($VM.AdminPassword.ToString().Trim() | ConvertTo-SecureString -AsPlainText -Force)).Password
        
        $subId = (Get-AzContext).Subscription.Id
        $desRg = if (-not [string]::IsNullOrWhiteSpace($VM.DesRG)) { $VM.DesRG.Trim() } else { $rgName }
        if (-not [string]::IsNullOrWhiteSpace($VM.DiskEncryptionSetId)) {
            $paramObj["diskEncryptionSetId"] = $VM.DiskEncryptionSetId.Trim()
        } elseif (-not [string]::IsNullOrWhiteSpace($VM.DESName)) {
            $paramObj["diskEncryptionSetId"] = "/subscriptions/$subId/resourceGroups/$desRg/providers/Microsoft.Compute/diskEncryptionSets/$($VM.DESName.Trim())"
        } else {
            throw "DESName 또는 DiskEncryptionSetId 값이 필요합니다. VM=$vmName"
        }
        $paramObj["diagnosticsStorageAccountName"] = $VM.DiagStrName.Trim()
        $paramObj["diagnosticsStorageAccountId"]   = "/subscriptions/$subId/resourceGroups/$($VM.DiagStrRG.Trim())/providers/Microsoft.Storage/storageAccounts/$($VM.DiagStrName.Trim())"
        $paramObj["virtualMachineZone"] = [string]$VM.Zones
        $paramObj["ResourceGroupName"] = $rgName

        # 4. ARM 배포 실행
        $deployResult = New-AzResourceGroupDeployment -ResourceGroupName $rgName `
            -TemplateFile ".\템플릿\VM\template-VM-LinuxZone.json" `
            -TemplateParameterObject $paramObj -ErrorAction Stop

        if ($deployResult.ProvisioningState -eq "Succeeded") {
            Write-Host "-> [SUCCESS] $vmName 배포 성공!" -ForegroundColor Green
        }
    } catch {
        Write-Host "-> [ERROR] $vmName 처리 중 오류: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n" + ("=" * 60) -ForegroundColor Gray
Write-Host "[삭제 검토] 엑셀 시트에서 제외된 리소스를 확인합니다." -ForegroundColor Yellow

$targetRGs = $VMs.RGName | Where-Object { $_ } | Select-Object -Unique
foreach ($rg in $targetRGs) {
    $azureVMs = Get-AzVM -ResourceGroupName $rg.Trim() -ErrorAction SilentlyContinue
    if ($null -ne $azureVMs) {
        foreach ($azVM in $azureVMs) {
            if ($excelVmNames -notcontains $azVM.Name) {
                Write-Host "`n[!] 감지됨: 엑셀에서 삭제된 VM ($($azVM.Name))" -ForegroundColor Magenta
                $confirmation = Read-Host "해당 VM과 관련 리소스(OS디스크, NIC)를 삭제하시겠습니까? (y/n)"
                
                if ($confirmation -eq 'y') {
                    $osDiskName = $azVM.StorageProfile.OsDisk.Name
                    $nicIds = $azVM.NetworkProfile.NetworkInterfaces.Id

                    # 순차 삭제 (VM -> Disk -> NIC)
                    Write-Host "-> VM 삭제 중..." -ForegroundColor Red
                    Remove-AzVM -ResourceGroupName $rg.Trim() -Name $azVM.Name -Force

                    Write-Host "-> OS 디스크($osDiskName) 삭제 중..." -ForegroundColor Red
                    Remove-AzDisk -ResourceGroupName $rg.Trim() -DiskName $osDiskName -Force

                    foreach ($nicId in $nicIds) {
                        $nicName = ($nicId -split '/')[-1]
                        Write-Host "-> NIC($nicName) 삭제 중..." -ForegroundColor Red
                        Remove-AzNetworkInterface -ResourceGroupName $rg.Trim() -Name $nicName -Force
                    }
                    Write-Host "-> 정리 완료." -ForegroundColor Gray
                }
            }
        }
    }
}

Write-Host "`n[전체 종료] 인프라 동기화가 완료되었습니다." -ForegroundColor White -BackgroundColor DarkBlue
#>
#endregion

####################################################################################################################################################################
####################################################################################################################################################################
####################################################################################################################################################################
####################################################################################################################################################################

#region | Marketplace(Windows) |
<#
$VMs = $xls | Where-Object {$_.Role -eq "WIN"} # Role 수정 필요
foreach ($VM in $VMs) {
New-AzResourceGroup -Name $VM.RGName -Location $VM.location -Force -ErrorAction:SilentlyContinue
    $paramObj = @{}
    $paramObj["location"] = $VM.location
    #NIC
    $paramObj["networkInterfaceName"] = $VM.Name + "-nic" # 리소스 네이밍룰 확인 필요
    $paramObj["subnetName"] = $VM.SubnetName
    $paramObj["virtualNetworkName"] = $VM.VnetName
    $paramObj["privateIP"] = $VM.PrivateIP
    #VM
    $paramObj["virtualMachineName"] = $VM.Name
    $paramObj["virtualMachineComputerName"] = $VM.Name
    $paramObj["virtualMachineRG"] = $VM.RGName
    $paramObj["virtualMachineSize"] = "Standard_" + $VM.VmSize
    #OS
    $osDiskType = switch ($VM.OsDiskStorageType) {
        PremiumSSD { "Premium_LRS" }
        StandardSSD { "StandardSSD_LRS" }
        StandardHDD { "Standard_LRS" }
        Default { "Standard_LRS" }
    }
    $paramObj["osDiskType"] = $osDiskType
  
    #ImageRef
    $paramObj["publisher"] = $VM.Publisher
    $paramObj["offer"] = $VM.Offer
    $paramObj["sku"] = $VM.Sku
    $paramObj["version"] = $VM.Version

    #Account
    $paramObj["adminUsername"] = $VM.AdminUsername
    $paramObj["adminPassword"] = (New-Object -TypeName PSCredential -ArgumentList 'id', ($VM.AdminPassword | ConvertTo-SecureString -AsPlainText -Force)).Password
    # Ref: https://github.com/Azure/azure-powershell/issues/12792 ## <- This seems to work, rather than passing password directly
#     
    #Diagnostics
    $subId = (Get-AzContext).Subscription.Id
    $storageRG = "TEST-Network-RG" # 부트진단 스토리지 계정 리소스 그룹명 확인 필요
    $diagStorageName = "20260313teststroage" # 부트진단 스토리지 계정 리소스명 확인 필요
    $paramObj["diagnosticsStorageAccountName"] = "20260313teststroage" #
    # $paramObj["diagnosticsStorageAccountId"] = "Microsoft.Storage/storageAccounts/20260313teststroage"
    $paramObj["diagnosticsStorageAccountId"] = "/subscriptions/$subId/resourceGroups/$storageRG/providers/Microsoft.Storage/storageAccounts/$diagStorageName"
    
    #AvailibilityZone
    $paramObj["virtualMachineZone"] = [string]$VM.Zones


     # $adminPw = $VM.AdminPassword | ConvertTo-SecureString -AsPlainText -Force

    New-AzResourceGroupDeployment -ResourceGroupName $VM.RGName -TemplateFile ".\템플릿\VM\template-VM-WindowsZone.json" -TemplateParameterObject $paramObj -Verbose #ARM 템플릿 체크 필요
    }
#>
#endregion
