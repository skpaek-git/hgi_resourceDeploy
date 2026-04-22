# Login
Connect-AzAccount

#region | Deploy RG |
    # Module
    Install-Module ImportExcel
    Import-Module ImportExcel
    $rgListData = Import-Excel ".\서버정보\20260422_리소스배포_종합.xlsx" -WorksheetName "RG"

    # Deploy RG
    Write-Host "----------------------------------------------------" -ForegroundColor Cyan
    Write-Host "[프로세스 시작] 리소스 그룹 배포 검토"

    foreach ($row in $rgListData) {
        # 필수 값(RGname)이 없거나 공백이면 건너뜀
        if ([string]::IsNullOrWhiteSpace($row.RGname)) { continue }
        
        $rgName = $row.RGname.Trim()
        # Location 정보가 없을 경우를 대비해 기본값(koreacentral) 설정 또는 엑셀 값 사용
        $location = if (-not [string]::IsNullOrWhiteSpace($row.Location)) { $row.Location.Trim() } else { "koreacentral" }

        # 1. 리소스 그룹 존재 여부 확인
        $existingRG = Get-AzResourceGroup -Name $rgName -ErrorAction SilentlyContinue

        if ($null -ne $existingRG) {
            # 이미 존재할 경우 메시지 출력 후 즉시 다음 행으로 이동 (연결 로직 등 모두 스킵)
            Write-Host "-> '$rgName' 리소스 그룹이 이미 생성되어 있습니다. (작업 건너뜀)" -ForegroundColor Gray
            continue
        }

        # 2. 신규 리소스 그룹 생성 (존재하지 않을 경우에만 실행)
        try {
            Write-Host "-> [신규 생성] 리소스 그룹($rgName) 생성 중..." -ForegroundColor Yellow
            
            # 순수하게 이름과 위치 정보만 사용하여 생성
            New-AzResourceGroup -Name $rgName -Location $location -Force | Out-Null
            
            Write-Host "[성공] '$rgName' 생성 완료 (위치: $location)" -ForegroundColor Green
        } catch {
            Write-Error "!! 리소스 그룹 생성 실패: $rgName - $($_.Exception.Message)"
        }
    }

    Write-Host "----------------------------------------------------"
    Write-Host "모든 리소스 그룹 배포 작업이 종료되었습니다." -ForegroundColor White -BackgroundColor DarkGreen
#endregion
