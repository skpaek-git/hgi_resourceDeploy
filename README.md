# Azure Resource Deployment Script

## 개요
`99. Deploy Resources.ps1`는 Excel 시트(`RG`, `VNET`, `Storage`, `KV`, `DES`, `LB`, `LB_Probe`, `LB_Rule`, `NSG`, `NSG_Detail/NSG_PRD_Rule`, `VM/VM_PRD`)를 읽어 리소스를 배포합니다.

현재 운영 기준:
- DES 자동 배포/수동 배포 병행 가능
- VM은 `DiskEncryptionSetId` 또는 `DESName` 기준으로 CMK 연결
- KV는 Key/Secret을 같은 시트에서 함께 처리 가능

## 보조 스크립트
- `3. Deploy KV.ps1`
- `4. Deploy DES.ps1`
- `5. Deploy LB.ps1`

두 스크립트는 내부적으로 `99. Deploy Resources.ps1`를 호출하며 동일한 파라미터(`-ExcelPath`, `-ConnectAccount`, `-DryRun`)를 사용합니다.

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

# Key Vault만
& ".\3. Deploy KV.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -ConnectAccount

# DES만
& ".\4. Deploy DES.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -ConnectAccount

# Load Balancer만
& ".\5. Deploy LB.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -ConnectAccount

# VM만
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('VM') -ConnectAccount

# NSG만
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('NSG') -ConnectAccount

# 통합 배포 (권장 순서)
& ".\99. Deploy Resources.ps1" -ExcelPath '.\서버정보\배포파일.xlsx' -DeployType @('RG','VNET','STORAGE','KV','DES','LB','NSG','VM') -ConnectAccount
```

## 실행 환경 권장
- `powershell`(5.1)보다 `pwsh`(7+) 사용 권장
- 예시:
```powershell
pwsh -NoProfile -File ".\99. Deploy Resources.ps1" -ExcelPath ".\서버정보\배포파일.xlsx" -DeployType @('KV','DES','VM') -ConnectAccount
```

## 핵심 동작 규칙
- 입력 검증 실패 시 배포 중단
- 실행 순서는 내부적으로 고정: `RG -> VNET -> STORAGE -> KV -> DES -> LB -> VM -> NSG`
- 로그는 `PSOutLog` 기반으로 `logs\99. Deploy Resources.log`에 기록

## CMK 연동 규칙 (현재)
- VM OS 디스크 CMK는 `diskEncryptionSetId` 템플릿 파라미터로 적용
- VM 시트에서 아래 중 하나 필요
  - `DiskEncryptionSetId`
  - `DESName` (+ 선택 `DesRG`)
- DES 시트(`DES`/`DES_PRD`) 자동 배포 지원

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
- DES 자동배포를 쓸 경우:
  - `DES_PRD`의 `KVName`이 `KV_PRD`에 반드시 존재해야 함
  - `DES_PRD`의 `KeyName`이 실제 Key Vault에 존재해야 함
- VM CMK 적용 시:
  - `DiskEncryptionSetId`를 쓰거나, `DESName`(+선택 `DesRG`)을 정확히 입력
- VM 비밀번호를 Key Vault Secret으로 쓸 경우:
  - `UseKeyVaultPassword=Y`
  - `AdminPasswordSecretUri` 또는 `AdminPasswordKVName + AdminPasswordSecretName` 필수
- LB 배포 시:
  - Internal LB는 `FEVNetRG/FEVNetName/FESubnetName` 정확히 입력
  - `FEVNetRG` 비우면 LB RG를 기본값으로 사용함

## 자주 발생하는 오류
- `KV 시트에 존재하지 않는 Key Vault 참조입니다`
  - DES 시트의 `KVName`과 KV 시트 `KVName` 불일치
- `Resource ... virtualNetworks/... was not found`
  - VM/LB의 `VnetName` 오타 또는 `VnetRG` 불일치
- `The term 'if' is not recognized...`
  - 구버전 스크립트에서 발생하던 문법 이슈이며 현재 버전에서는 수정됨
- `SubscriptionId ... must match ... contained in the Key Vault Id`
  - DES와 Source Key Vault의 구독 불일치(서비스 제약)

## Excel 컬럼 가이드
- KV 시트 필수:
  - `KVName`, `RGname`, `Location`
  - `KeyName` 또는 `SecretName`
- KV 시트 Secret 관련 선택:
  - `SecretValue`
  - `SecretContentType` (또는 오타 호환 `SecretContenctType`)
  - `SecretExpiresOn`
  - `SecretNotBefore`
  - `Enable`
- VM 시트 CMK 필수:
  - `DiskEncryptionSetId` 또는 `DESName`(+선택 `DesRG`)

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
