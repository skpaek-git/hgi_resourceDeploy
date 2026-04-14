# Agent Best Practices

## 목적
Excel 기반 Azure 배포에서 CMK 포함 리소스 연계를 안전하게 자동화한다.

## 작업 원칙
- `PLAN.md`를 먼저 작성하고 순서대로 진행한다.
- 배포 전 입력 검증 실패 시 즉시 중단한다.
- 시작/종료, 함수 시작/종료, 중요 이벤트는 `PSOutLog`로 기록한다.
- 배포 타입은 파라미터로 선택 실행하고 내부 실행 순서는 의존성 기준으로 고정한다.

## CMK 필수 규칙
- VM 배포 전 Key Vault와 DES가 준비되어야 한다.
- VM 시트에는 `DESName` 또는 `DiskEncryptionSetId`가 반드시 있어야 한다.
- DES 생성 후 시스템할당 ID에 `Key Vault Crypto Service Encryption User` 권한을 부여한다.

## 입력 데이터 규칙
- 시트 자동 탐색: `VM/VM_PRD`, `KV/KeyVault`
- KV 시트 필수: `KVName`, `RGname`, `Location`, `KeyName`
- DES 시트 필수: `DESName`, `RGname`, `Location`, `KeyVaultName`, `KeyName`

## 템플릿 규칙
- VM 템플릿 4종은 `diskEncryptionSetId` 파라미터를 유지해야 한다.
- 하드코딩된 RG/리소스 식별자는 파라미터화한다.

## 테스트 규칙
- 실제 배포 전 `-DryRun` 검증 필수
- 모듈/PowerShell 버전 호환성 확인
- CMK 관련 변경 시 템플릿+스크립트+입력검증을 함께 검토

## 이번 작업 실수 방지 반영
- VM CMK 누락 방지용 입력검증(`DESName`/`DiskEncryptionSetId` 필수) 추가
- 실행 순서 강제(`KV -> DES -> VM`) 반영
- DES 권한 할당 자동화 반영
