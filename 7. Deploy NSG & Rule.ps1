# Login
Connect-AzAccount

#region | Deploy NSG |
    # Module
    Install-Module ImportExcel
    Import-Module ImportExcel
    $ruleListData = Import-Excel ".\서버정보\20260422_리소스배포_종합.xlsx" -WorksheetName "NSG"

    # Deploy NSG
        $nsgGroups = $nsgListData <#| Where-Object {$_.Role -eq "WEB3"}#>| Group-Object NSGName 
        foreach ($group in $nsgGroups) {
            $row = $group.Group[0] 
            $rgName = $row.RG.Trim()
            $location = "koreacentral"
            $nsgName = $group.Name.Trim()

            Write-Host "----------------------------------------------------" -ForegroundColor Cyan
            Write-Host "[프로세스 시작] NSG: $nsgName"

            # 1. NSG 리소스 존재 여부 확인
            $nsg = Get-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $rgName -ErrorAction SilentlyContinue
            
            if ($null -ne $nsg) {
                # [수정] 이미 존재할 경우 메시지 출력 후 즉시 다음 group(NSG)으로 넘어감
                Write-Host "-> '$nsgName' 리소스가 이미 생성되어 있습니다. (작업 건너뜀)" -ForegroundColor Gray
                continue 
            }

            # 2. 신규 NSG 생성 (존재하지 않을 경우에만 실행됨)
            Write-Host "-> [신규 생성] NSG가 존재하지 않아 생성을 시작합니다..." -ForegroundColor Yellow
            $nsg = New-AzNetworkSecurityGroup -Name $nsgName -ResourceGroupName $rgName -Location $location -Force
            Write-Host "-> NSG 생성 완료" -ForegroundColor Green

            # 3. 연결 타겟 자동 판단 로직
            $linkTarget = ""
            if ($null -ne $row.UseNIC -and $row.UseNIC.Trim().ToUpper() -eq "O") {
                $linkTarget = "O" # NIC 연결 모드
            } 
            elseif (-not [string]::IsNullOrWhiteSpace($row.Subnet)) {
                $linkTarget = "X" # Subnet 연결 모드
            }

            # 4. 분기 처리 수행 (신규 생성된 경우에만 연결 수행)
            if ($linkTarget -eq "O") {
                Write-Host "-> 타겟 분석: NIC ($($row.NIC)) 연결 시도"
                $nic = Get-AzNetworkInterface -Name $row.NIC.Trim() -ResourceGroupName $rgName -ErrorAction SilentlyContinue
                
                if ($null -ne $nic) {
                    $nic.NetworkSecurityGroup = $nsg
                    $nic | Set-AzNetworkInterface | Out-Null
                    Write-Host "[성공] NIC($($row.NIC)) 연결 완료" -ForegroundColor Green
                } else {
                    Write-Warning "[실패] NIC를 찾을 수 없습니다: $($row.NIC)"
                }
            } 
            elseif ($linkTarget -eq "X") {
                Write-Host "-> 타겟 분석: Subnet ($($row.Subnet)) 자동 연결 시도"
                $vnet = Get-AzVirtualNetwork | Where-Object { $_.Name -eq $row.VirtualNetwork.Trim() }
                
                if ($null -ne $vnet) {
                    $subnetFound = $false
                    foreach ($sub in $vnet.Subnets) {
                        if ($sub.Name -eq $row.Subnet.Trim()) {
                            $sub.NetworkSecurityGroup = $nsg
                            $subnetFound = $true
                            break
                        }
                    }

                    if ($subnetFound) {
                        $vnet | Set-AzVirtualNetwork | Out-Null
                        Write-Host "[성공] 서브넷($($row.Subnet)) 연결 완료" -ForegroundColor Green
                    } else {
                        Write-Warning "[실패] VNet 내에 서브넷($($row.Subnet))이 존재하지 않습니다."
                    }
                } else {
                    Write-Warning "[실패] 가상 네트워크($($row.VirtualNetwork))를 찾을 수 없습니다."
                }
            }
        }
        Write-Host "----------------------------------------------------"
        Write-Host "모든 작업이 완료되었습니다." -ForegroundColor White -BackgroundColor DarkGreen


#endregion

############################################################################################################################################################################
############################################################################################################################################################################
############################################################################################################################################################################

#region | ADD NSG Rules |
    # Module
    Install-Module ImportExcel
    Import-Module ImportExcel
    $ruleListData = Import-Excel ".\서버정보\20260422_리소스배포_종합.xlsx" -WorksheetName "NSG_Rule"

    $ruleGroups = $ruleListData <#| Where-Object {$_.Role -eq "WEB3"} #>| Group-Object NSGName
    foreach ($group in $ruleGroups) {
        $nsgName = $group.Name.Trim()
        $nsg = Get-AzNetworkSecurityGroup | Where-Object { $_.Name -eq $nsgName }
        if ($null -eq $nsg) { continue }

        Write-Host "----------------------------------------------------" -ForegroundColor Cyan
        Write-Host "[규칙 검토 및 동기화] NSG: $nsgName"
        
        $isChanged = $false
        $excelRuleNames = $group.Group | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } | ForEach-Object { $_.Name.Trim() }

        # --- [1. 삭제 감지 로직] ---
        $rulesToRemove = $nsg.SecurityRules | Where-Object { ($excelRuleNames -notcontains $_.Name) -and ($_.Priority -lt 65000) }
        foreach ($ruleToDelete in $rulesToRemove) {
            Write-Host "!! 경고: 규칙 '$($ruleToDelete.Name)'이 엑셀에 없습니다." -ForegroundColor Red
            $confirm = Read-Host "-> Azure에서 삭제할까요? (y/n)"
            if ($confirm -eq 'y') {
                $nsg.SecurityRules.Remove($ruleToDelete)
                $isChanged = $true
                Write-Host "   - 제거 예약됨" -ForegroundColor Gray
            }
        }

        # --- [2. 추가 및 업데이트 로직] ---
        foreach ($ruleRow in $group.Group) {
            if ([string]::IsNullOrWhiteSpace($ruleRow.Priority) -or [string]::IsNullOrWhiteSpace($ruleRow.Name)) { continue }

            $rName = $ruleRow.Name.Trim()
            $existingRule = $nsg.SecurityRules | Where-Object { $_.Name -eq $rName }
            if ($null -ne $existingRule) {
                Write-Host "-> [유지] '$rName'" -ForegroundColor Gray
                continue
            }

            # [에러 수정] 배열 변환 함수 보정
            $getArrayVal = {
                param($val)
                # 빈 값, Any, * 인 경우 확실하게 "*" 반환 (Null 에러 방지)
                if ([string]::IsNullOrWhiteSpace($val) -or $val.ToString().Trim().ToUpper() -eq "ANY" -or $val.ToString().Trim() -eq "*") { 
                    return "*" 
                }
                
                $rawStr = $val.ToString().Replace(" ", "").Trim()
                # 오타 교정
                if ($rawStr -eq "Internt") { $rawStr = "Internet" }
                if ($rawStr -eq "VirtualNetwork") { $rawStr = "VirtualNetwork" }

                # [문법 수정] return 뒤에 if를 직접 쓰지 않고 결과를 변수에 담아 반환
                if ($rawStr.Contains(",")) { return $rawStr.Split(",") }
                return $rawStr
            }

            $rProtocol = if ([string]::IsNullOrWhiteSpace($ruleRow.Protocol) -or $ruleRow.Protocol.Trim().ToUpper() -eq "ANY") { "*" } else { $ruleRow.Protocol.Trim() }
            $rSrcAddr  = &$getArrayVal $ruleRow."Src Addr"
            $rSrcPort  = &$getArrayVal $ruleRow."Src Port"
            $rDstAddr  = &$getArrayVal $ruleRow."Dest Addr"
            $rDstPort  = &$getArrayVal $ruleRow."Dest Port"
            $rDesc     = if ($null -ne $ruleRow.Description) { $ruleRow.Description.Trim() } else { " " }

            Write-Host "-> [신규 등록] $rName (Priority: $($ruleRow.Priority))" -ForegroundColor Yellow

            $params = @{
                Name                     = $rName
                Priority                 = $ruleRow.Priority
                Direction                = $ruleRow.Direction.Trim()
                Access                   = $ruleRow.Action.Trim()
                Protocol                 = $rProtocol
                Description              = $rDesc
                SourceAddressPrefix      = $rSrcAddr
                SourcePortRange          = $rSrcPort
                DestinationAddressPrefix = $rDstAddr
                DestinationPortRange     = $rDstPort
            }

            try {
                $nsg | Add-AzNetworkSecurityRuleConfig @params | Out-Null
                $isChanged = $true
            } catch {
                Write-Error "!! 규칙 추가 실패: $rName - $($_.Exception.Message)"
            }
        }

        if ($isChanged) {
            $nsg | Set-AzNetworkSecurityGroup | Out-Null
            Write-Host "[완료] NSG($nsgName) 동기화 완료" -ForegroundColor Green
        } else {
            Write-Host "[완료] 변경 사항 없음" -ForegroundColor Gray
        }
    }


#endregion
