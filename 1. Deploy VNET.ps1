# Login
Connect-AzAccount

#region | Deploy VNET |
    # Module
    Install-Module ImportExcel
    Import-Module ImportExcel
    $vnetListData = Import-Excel ".\서버정보\20260422_리소스배포_종합.xlsx" -WorksheetName "VNET"

    # Deploy VNET
    Write-Host "----------------------------------------------------" -ForegroundColor Cyan
    Write-Host "[프로세스 시작] 가상 네트워크(VNet) 설정 비교 및 업데이트 검토"

    foreach ($row in $vnetListData) {
        if ([string]::IsNullOrWhiteSpace($row."VNet1 Name")) { continue }
        
        $rgName = $row.RGname.Trim(); $vnetName = $row."VNet1 Name".Trim(); $location = $row.Location.Trim()

        # 1. 엑셀 데이터 수집 (현재 기준점)
        $excelVnetAddrs = $row.PSObject.Properties | Where-Object { $_.Name -like "VNet* Address" -and -not [string]::IsNullOrWhiteSpace($_.Value) } | ForEach-Object { $_.Value.ToString().Trim() }
        $excelSubnets = @{}
        $i = 1
        while ($true) {
            $sNameCol = "Subnet$i Name"; $sAddrCol = "Subnet$i Address"
            if (-not ($row.PSObject.Properties.Name -contains $sNameCol)) { break }
            if (-not [string]::IsNullOrWhiteSpace($row.$sNameCol)) { $excelSubnets[$row.$sNameCol.Trim()] = $row.$sAddrCol.Trim() }
            $i++
        }

        # 2. Azure 리소스 확인
        $vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $rgName -ErrorAction SilentlyContinue
        if ($null -eq $vnet) { <# 신규 생성 로직 생략 #> continue }

        Write-Host "`n[분석 중] VNet: $vnetName" -ForegroundColor Cyan
        $isChanged = $false

        # --- [삭제 감지 로직] ---
        
        # A. 서브넷 삭제 감지
        # Azure에는 있지만 엑셀 목록에는 없는 서브넷을 찾음
        $subnetsToRemove = $vnet.Subnets | Where-Object { -not $excelSubnets.ContainsKey($_.Name) }
        
        foreach ($sToRemove in $subnetsToRemove) {
            $confirm = Read-Host "!! 경고: 서브넷 '$($sToRemove.Name)'이 엑셀에 없습니다. Azure에서 삭제할까요? (y/n)"
            if ($confirm -eq 'y') {
                $vnet.Subnets.Remove($sToRemove)
                Write-Host "   - 서브넷 제거 예약: $($sToRemove.Name)" -ForegroundColor Red
                $isChanged = $true
            }
        }

        # B. 주소 대역 삭제 감지
        # Azure에는 설정되어 있지만 엑셀 주소 목록에는 없는 대역을 찾음
        $addrsToRemove = $vnet.AddressSpace.AddressPrefixes | Where-Object { $excelVnetAddrs -notcontains $_ }
        
        foreach ($aToRemove in $addrsToRemove) {
            $confirm = Read-Host "!! 경고: 주소 대역 '$aToRemove'이 엑셀에 없습니다. 삭제할까요? (y/n)"
            if ($confirm -eq 'y') {
                $vnet.AddressSpace.AddressPrefixes.Remove($aToRemove)
                Write-Host "   - 주소 대역 제거 예약: $aToRemove" -ForegroundColor Red
                $isChanged = $true
            }
        }

        # --- [추가 감지 로직] ---
        # (이전 스크립트와 동일하게 신규 항목 추가)
        foreach ($addr in $excelVnetAddrs) {
            if ($vnet.AddressSpace.AddressPrefixes -notcontains $addr) {
                $vnet.AddressSpace.AddressPrefixes.Add($addr)
                Write-Host "   + 주소 대역 추가: $addr" -ForegroundColor Green
                $isChanged = $true
            }
        }
        foreach ($sName in $excelSubnets.Keys) {
            if ($null -eq ($vnet.Subnets | Where-Object { $_.Name -eq $sName })) {
                Add-AzVirtualNetworkSubnetConfig -VirtualNetwork $vnet -Name $sName -AddressPrefix $excelSubnets[$sName] | Out-Null
                Write-Host "   + 서브넷 추가: $sName" -ForegroundColor Green
                $isChanged = $true
            }
        }

        # 최종 반영
        if ($isChanged) {
            $vnet | Set-AzVirtualNetwork | Out-Null
            Write-Host "=> 업데이트 완료" -ForegroundColor White -BackgroundColor DarkBlue
        } else {
            Write-Host "-> 변경 사항 없음" -ForegroundColor Gray
        }
    }
#endregion



