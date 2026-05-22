# Login
Connect-AzAccount

#region | Deploy Storage |
    # Module
    Install-Module ImportExcel
    Import-Module ImportExcel
    $storageListData = Import-Excel ".\서버정보\20260422_리소스배포_종합.xlsx" -WorksheetName "Storage"

    # Deploy Storage(Boot)
    Write-Host "----------------------------------------------------" -ForegroundColor Cyan
    Write-Host "[프로세스 시작] Azure 스토리지 계정 배포 검토"

    foreach ($row in $storageListData) {
        # 필수 정보인 StorageName이 없으면 건너뜀
        if ([string]::IsNullOrWhiteSpace($row."StorageName")) { continue }
        
        $rgName      = $row.RGname.Trim()
        $location    = $row.Location.Trim()
        $storageName = $row."StorageName".Trim().ToLower() # 스토리지 이름은 반드시 소문자여야 함

        # 1. 기존 리소스 존재 여부 확인
        $existingStorage = Get-AzStorageAccount -ResourceGroupName $rgName -Name $storageName -ErrorAction SilentlyContinue

        if ($null -ne $existingStorage) {
            Write-Host "-> '$storageName' 리소스가 이미 존재합니다. (작업 건너뜀)" -ForegroundColor Gray
            continue
        }

        Write-Host "----------------------------------------------------"
        Write-Host "[신규 배포 대상] Storage: $storageName" -ForegroundColor Yellow

        # 2. SKU 및 Kind 설정 (CSV에 없을 경우를 대비한 기본값 설정)
        $sku  = if (-not [string]::IsNullOrWhiteSpace($row.SkuName)) { $row.SkuName.Trim() } else { "Standard_LRS" }
        $kind = if (-not [string]::IsNullOrWhiteSpace($row.Kind)) { $row.Kind.Trim() } else { "StorageV2" }

        # 3. 스토리지 계정 생성 실행
        try {
            $storageParams = @{
                ResourceGroupName = $rgName
                Name              = $storageName
                Location          = $location
                SkuName           = $sku
                Kind              = $kind
                EnableHttpsTrafficOnly = $true # 보안 베스트 프랙티스 적용
            }

            New-AzStorageAccount @storageParams | Out-Null
            Write-Host "[성공] '$storageName' 배포 완료" -ForegroundColor Green
        } catch {
            Write-Error "!! 스토리지 생성 실패: $storageName - $($_.Exception.Message)"
        }
    }

    Write-Host "----------------------------------------------------"
    Write-Host "모든 스토리지 배포 검토 작업이 종료되었습니다." -ForegroundColor White -BackgroundColor DarkGreen

#endregion



