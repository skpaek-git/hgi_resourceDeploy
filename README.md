# Azure Resource Deployment Script

## 개요
`99. Deploy Resources.ps1`는 Excel 시트(`RG_*`, `VNET_*`, `UDR_*`, `LB_*`, `NSG_*`, `VM_*`, `VM_Datadisk`)를 읽어 리소스를 배포합니다.

현재 운영 기준:
- `STORAGE`, `KV`, `DES`는 통합 배포 대상에서 제외
- 기본 배포 타입은 `RG`, `VNET`, `UDR`, `LB`, `NSG`, `VM`, `DATADISK`
- LB Probe에서 `Port`가 비어 있는 행은 오류가 아닌 스킵 처리

## 보조 스크립트
- `5. LEG_Deploy LB.ps1`

`5` 스크립트는 내부적으로 `99. Deploy Resources.ps1`를 호출하며 동일한 파라미터(`-ExcelPath`, `-ConnectAccount`, `-DryRun`)를 사용합니다.
`3. LEG_Deploy KV.ps1`, `4. LEG_Deploy DES.ps1`는 현재 운영 기준에서는 사용하지 않습니다(레거시).
`0/1/2/6/7` 레거시 스크립트는 과거 단독 배포용이며, 현재 통합 운영 기준에서는 `99` 중심 실행을 권장합니다.

DataDisk 배포는 `99. Deploy Resources.ps1`의 `DeployType DATADISK`로 통합 처리됩니다(외부 `8` 스크립트 의존 없음).
- 시트 스키마(현재): `VMName`, `RGname`, `Location`, `Zone(or Zones)`, `DESRG`, `DESName`, `DoubleEncryption`, `HostCaching`, `Datadisk1Name`, `DataDisk1Type`, `Datadisk1Size`, `Datadisk2Name`, `DataDisk2Type`, `Datadisk2Size`
- 필수(디스크별): `Datadisk{N}Name`, `DataDisk{N}Type=StandardSSD_LRS`, `Datadisk{N}Size`
- 필수(행 공통): `VMName`, `RGname`(또는 `DataDiskRG`/`DiskRG`), `Location`, `Zone`, `DESName`+`DESRG` 또는 `DiskEncryptionSetId`
- 이중 암호화 컬럼: `Datadisk1DoubleEncryption`, `Datadisk2DoubleEncryption` 또는 행 공통 `DoubleEncryption` (`TRUE/FALSE`)
- 이중 암호화 컬럼이 없으면 기본 `TRUE`로 처리되어 `EncryptionAtRestWithPlatformAndCustomerKeys`가 적용됩니다.
- `HostCaching` 기본값은 `None`이며, 허용값은 `None/ReadOnly/ReadWrite`입니다.
- 디스크 입력이 없는 행(`Datadisk1/2` 관련 값이 모두 공백 또는 `-`)은 자동 스킵됩니다.

## 필수 모듈
- ImportExcel
- PSOutLog
- Az.Accounts
- Az.Resources
- Az.Network
- Az.Storage
- Az.Compute
- Az.KeyVault

## 실행 예시
```powershell
# 주의: 파일명에 공백이 있으면 반드시 호출 연산자(&) + 따옴표를 사용해야 합니다.
# 잘못된 예: .\99. Deploy Resources.ps1 ...
# 올바른 예: & ".\99. Deploy Resources.ps1" ...

# Load Balancer만
& ".\5. LEG_Deploy LB.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -ConnectAccount

# Load Balancer만 (Role/Option 필터)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB') -Option 'EXT' -ConnectAccount

# LB 리소스만 선배포(Probe/Rule 제외)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB') -Option 'DTG' -LbResourceOnly -ConnectAccount

# LB Probe만 배포
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB_PROBE') -Option 'DTG' -ConnectAccount

# LB Rule만 배포
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB_RULE') -Option 'DTG' -ConnectAccount

# LB 분리 배포 가이드
# 1) LB 리소스(Frontend/BackendPool)만 먼저 생성
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB') -Option 'DTG' -LbResourceOnly -ConnectAccount

# 2) 상태 프로브만 별도 반영(사전 협의 후)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB_PROBE') -Option 'DTG' -ConnectAccount

# 3) 부하분산 규칙만 별도 반영(사전 협의 후)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB_RULE') -Option 'DTG' -ConnectAccount

# 4) 프로브 + 룰을 한 번에 반영
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB_PROBE','LB_RULE') -Option 'DTG' -ConnectAccount

# 5) 기존 방식(리소스 + 프로브 + 룰 일괄)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('LB') -Option 'DTG' -ConnectAccount

# VM만
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('VM') -ConnectAccount

# NSG만
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('NSG') -ConnectAccount

# UDR만 (Option 필터: EXT/INT)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('UDR') -Option 'EXT' -ConnectAccount

# 통합 배포 (권장 순서)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('RG','VNET','UDR','LB','NSG','VM','DATADISK')

# 통합 스크립트로 Data Disk만 배포 (Option 필터)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('DATADISK') -Option 'TEST' -ConnectAccount
```

## 실행 환경 권장
- `powershell`(5.1)보다 `pwsh`(7+) 사용 권장
- 예시:
```powershell
pwsh -NoProfile -File ".\99. Deploy Resources.ps1" -ExcelPath ".\서버정보\배포파일.xlsx" -DeployType @('RG','VNET','UDR','LB','NSG','VM','DATADISK')
```

## 핵심 동작 규칙
- 입력 검증 실패 시 배포 중단
- 실행 순서는 내부적으로 고정: `RG -> VNET -> UDR -> LB -> LB_PROBE -> LB_RULE -> VM -> DATADISK -> NSG`
- 로그는 `PSOutLog` 기반으로 `logs\99. Deploy Resources.log`에 기록

## CMK 연동 규칙 (현재)
- VM OS 디스크 CMK는 `diskEncryptionSetId` 템플릿 파라미터로 적용
- VM 시트에서 아래 중 하나 필요
  - `DiskEncryptionSetId`
  - `DESName` (+ `DESRG` 권장)
- DES는 외부에 사전 생성된 리소스를 참조하는 방식으로 운영
- `UseOsDiskDoubleEncryption=TRUE`이면 연결되는 DES의 암호화 타입이 `EncryptionAtRestWithPlatformAndCustomerKeys`여야 합니다.
- `EnableAcceleratedNetworking` 컬럼으로 NIC 가속 네트워킹을 제어합니다(기본 `TRUE`).
- 포털 기준 해석:
  - `EnableAcceleratedNetworking=TRUE` -> `사용`(명시적 활성화)
  - `EnableAcceleratedNetworking=FALSE` -> `사용 안 함`
  - ARM bool 파라미터 특성상 `자동(권장)`을 직접 지정하는 값은 없음

## VM Admin Password 규칙
- `UseKeyVaultPassword=Y`: Key Vault Secret 사용
- 빈값/N: `AdminPassword` 사용
- 우선순위:
  1) `AdminPasswordSecretUri`
  2) `AdminPasswordKVName + AdminPasswordSecretName (+ AdminPasswordSecretVersion)`
  3) `AdminPassword`

참고:
- VM이 Key Vault를 직접 조회하는 구조가 아니라, 배포 실행 계정이 Secret을 조회해 VM 생성 파라미터로 전달
- Key Vault가 Private Endpoint 전용이면, 배포 실행 위치에서 Private Endpoint + Private DNS 접근 가능해야 Secret 조회 성공

## 데이터 작성 체크리스트
- VM CMK 적용 시:
  - `DiskEncryptionSetId`를 쓰거나, `DESName`(+선택 `DesRG`)을 정확히 입력(외부에 사전 생성된 DES 기준)
- VM 비밀번호를 Key Vault Secret으로 쓸 경우:
  - `UseKeyVaultPassword=Y`
  - `AdminPasswordSecretUri` 또는 `AdminPasswordKVName + AdminPasswordSecretName` 필수(외부/기존 Key Vault 기준)
- LB 배포 시:
  - Internal LB는 `FEVNetRG/FEVNetName/FESubnetName` 정확히 입력
  - `FEVNetRG` 비우면 LB RG를 기본값으로 사용함
  - `-Option` 사용 시 `LB`, `LB_Probe`, `LB_Rule` 시트 모두 `Role`(또는 `Option`) 값으로 필터링됨
  - `LB_Probe`에서 `ProbeName` 또는 `Port`가 비어 있으면 해당 행은 스킵됨

## 자주 발생하는 오류
- `Resource ... virtualNetworks/... was not found`
  - VM/LB의 `VnetName` 오타 또는 `VnetRG` 불일치
  - 혹은 Command 실행 환경에서 Context Setting이 잘못되어 있을수 있음
    - Bash(Azure CLI) : az account set -s {Subscription id}
    - PowerShell(Azure Powershell) : Set-azContext -Subscription {Subscription id}
- `The term 'if' is not recognized...`
  - 구버전 스크립트에서 발생하던 문법 이슈이며, 현재 버전에서는 수정됨
- `SubscriptionId ... must match ... contained in the Key Vault Id`
  - DES와 Source Key Vault의 구독 불일치(서비스 제약)

## Excel 컬럼 가이드
- VM 시트(권장/신규 포함):
  - `Name`, `RGname`, `Location`, `Zones`
  - `EnableAcceleratedNetworking` (`TRUE/FALSE`, 기본 `TRUE`)
  - `UseOsDiskDoubleEncryption` (`TRUE/FALSE`)
  - `DESName`, `DESRG` 또는 `DiskEncryptionSetId`

- VM_Datadisk 시트(데이터디스크 배포):
  - 공통: `VMName`, `RGname`, `Location`, `Zone(or Zones)`, `DESName`, `DESRG`(또는 `DiskEncryptionSetId`)
  - 디스크: `Datadisk1Name`, `DataDisk1Type`, `Datadisk1Size`, `Datadisk2Name`, `DataDisk2Type`, `Datadisk2Size`
  - 옵션: `DoubleEncryption`, `HostCaching`, `Datadisk{N}Lun`

- UDR 시트(UDR 배포):
  - 공통: `Enable`, `Option(EXT/INT)`(또는 `UDRType`/`Direction`), `RGname`, `Location`, `UDRName`(또는 `RouteTableName`), `DisableBGPRoutePropagation`
  - 라우트: `RouteName`, `AddressPrefix`, `NextHopType`, `NextHopIpAddress(NextHopType=VirtualAppliance 시 필수)`
  - 연결(선택): `AssociateToSubnet(TRUE/FALSE)`, `VnetName`, `SubnetName`, `VnetRG`(미입력 시 `RGname` 사용)

## Release Note
### 2026-03-20 (CMK/KV 기반 초기 확장)
- 통합 스크립트에 KV 처리 확장
- VM 템플릿 4종에 `diskEncryptionSetId` 파라미터 반영
- VM 스크립트에 CMK 전달 로직 반영

### 2026-03-24 (LB 통합)
- LB 배포 기능 통합 (`LB`, `LB_Probe`, `LB_Rule`)
- `5. Deploy LB.ps1` 추가
- NSG 상세 시트명 호환 확장(`NSG_Detail`, `NSG_Rule`, `NSG_PRD_Rule`)

### 2026-03-27 (KV Secret + VM Password 연동)
- KV 시트에서 Key/Secret 동시 처리 지원
- `Enable` 행 스킵 처리, Secret 선택 컬럼 빈값 허용
- VM에서 `UseKeyVaultPassword` 분기 기반 Secret 조회 지원

### 2026-04-03 (DES 자동배포 복구)
- DES 배포 타입(`DES`) 재활성화
- `DES`/`DES_PRD` 시트 검증 및 배포 함수 복구
- VM의 DES 해석 로직(`DESName`/`DesRG`) 복구
- `4. Deploy DES.ps1` 복구

### 2026-04-20 (VM NIC/OS 암호화 + DataDisk 분리 배포 강화)
- VM 템플릿 4종에 `enableAcceleratedNetworking` 파라미터 반영
- VM 시트의 `EnableAcceleratedNetworking` 값을 NIC에 적용
- VM 시트의 `UseOsDiskDoubleEncryption=TRUE` 검증 로직 추가(DES 암호화 타입 검사)
- DataDisk 스크립트에서 타 시트 값 보완 제거(행 내 컬럼만 사용)
- DataDisk 스크립트에서 빈 디스크 행 자동 스킵 및 `HostCaching` 컬럼 처리
- DataDisk 스크립트 `-AttachToVm`로 생성 후 VM 자동 연결 지원
- `99. Deploy Resources.ps1`에 `-DeployType DATADISK` 추가
- `99`의 `-Option` 필터를 DataDisk 시트(`Option/Role/VmRole` 컬럼)에도 동일 적용

### 2026-04-23 (UDR 배포 추가)
- `99. Deploy Resources.ps1`에 `-DeployType UDR` 추가
- `UDR` 시트 기준 Route Table 생성/갱신, Route 등록, Subnet 연결 지원
- `-Option` 필터로 `EXT`/`INT` 분리 실행 지원

### 2026-04-28 (LB Role/Option 필터 + UDR 안정화)
- `LB` 배포(검증/실행)에서 `-Option` 필터를 `Role`/`Option` 컬럼 기준으로 적용
- `LB`, `LB_Probe`, `LB_Rule` 각각에 `-Option` 필터를 직접 적용하고, 이후 선택된 LB에 매핑된 Probe/Rule만 처리하도록 보강
- UDR 라우트 반영 시 신규 Route는 `Add`, 기존 Route는 `Set`으로 분기 처리

### 2026-05-14 (운영 안정화/단계 배포 강화)
- Windows VM 배포 시 `computerName`을 자동 보정하도록 개선 (VM 이름 15자 초과 시 앞 15자로 자동 적용)
- Windows VM `computerName` 규칙 사전 검증 로직 보강 (길이/허용문자/숫자-only 체크)
- ARM 호출 일시 오류(예: `An error occurred while sending the request`, timeout/429)에 대한 재시도 로직 추가 (`Get/New-AzResourceGroup`, `New-AzResourceGroupDeployment`)
- DataDisk 시트에서 `-`/빈값 처리 시 디스크 슬롯 인식 기준을 강화 (`Name`과 `Size`가 모두 있을 때만 배포 대상으로 간주)
- LB 단계 배포를 위해 `-LbResourceOnly` 옵션 추가 (LB 리소스만 선배포, Probe/Rule 제외)
- `DeployType`에 `LB_PROBE`, `LB_RULE` 추가 (시트 단위 실행 지원)
- `LB_PROBE`/`LB_RULE` 단독 실행 시 해당 항목 검증/처리만 수행하도록 분기 로직 정비

### 2026-05-22 (운영 정책 반영)
- `99. Deploy Resources.ps1`에서 `STORAGE`, `KV`, `DES`를 배포 타입에서 제외
  - `DeployType` ValidateSet/기본값/실행순서 동기화
- LB Rule 프로토콜 컬럼을 `FEProtocol` 또는 `Protocol`로 읽도록 유지
- LB Probe `Port` 미입력 행은 검증 오류로 중단하지 않고 스킵 처리
  - 실제 배포 시 `[SKIP]` 로그 출력
  - DryRun에서도 `[DryRun][SKIP]`로 동일하게 표시
- DataDisk 배포 로직을 `99` 내부로 통합
  - `8. Deploy DataDisk.ps1` 외부 호출 의존 제거
  - `DeployType DATADISK` 실행만으로 `VM_Datadisk` 처리 가능
- VNET 시트의 정의 서브넷 생성/갱신 시 정책 강제
  - `Private Subnet` 기준으로 `DefaultOutboundAccess=false` 적용
  - `PrivateEndpointNetworkPolicies=Enabled` 적용(NSG/UDR 정책 적용 가능 상태)
