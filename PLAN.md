# PLAN

## 목표
- VM OS 디스크 CMK 활성화를 위해 Key Vault, DES, VM 연계를 Excel 기반 자동화로 구현
- `3. Deploy KV.ps1`, `4. Deploy DES.ps1`를 추가하고 통합 스크립트와 동일 파라미터 방식으로 실행
- VM 배포 시 DES를 템플릿 파라미터로 전달하여 CMK가 배포 시점에 적용되도록 구성

## 작업 순서
- [x] 요구사항 분석
- [x] VM 템플릿 4종에 `diskEncryptionSetId` 파라미터 추가
- [x] `99. Deploy Resources.ps1`에 `KV`, `DES` 배포 타입 추가
- [x] `3. Deploy KV.ps1`, `4. Deploy DES.ps1` 작성
- [x] `5. Deploy VM.ps1`에 DES 컬럼 기반 CMK 파라미터 전달 반영
- [x] DryRun/문법 테스트
- [x] README/TODO/Agent 문서 업데이트

## 사전 정의/제안
- Excel에 신규 시트 `KV`, `DES`를 추가하고 아래 컬럼을 운영 표준으로 확정 필요
  - KV: `KVName`, `RGname`, `Location`, `KeyName` (권장: `SkuName`, `KeyType`, `KeySize`)
  - DES: `DESName`, `RGname`, `Location`, `KeyVaultName`, `KeyName` (권장: `KeyVersion`, `KeyVaultRG`)
- VM 시트는 `DESName`(+선택 `DesRG`) 또는 `DiskEncryptionSetId` 중 하나를 필수 입력
