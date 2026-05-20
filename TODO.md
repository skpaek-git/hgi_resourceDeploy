# TODO

## 추가 요구사항/개선 과제
- [x] (1단계) Windows VM 템플릿 VNet RG 하드코딩 제거 (`vnetRG` 파라미터화)
- [x] (2단계) CMK 연동을 위한 Key Vault 배포 스크립트 및 통합 배포 타입(`KV`) 추가
- [x] (3단계) VM OS 디스크 CMK 적용(`diskEncryptionSetId`) 템플릿/스크립트 반영
- [x] (4단계) Load Balancer 시트(`LB/LB_PRD`) 배포 기능 통합
- [x] (5단계) NSG 상세 시트명 호환 확장(`NSG_Detail`, `NSG_Rule`, `NSG_PRD_Rule`)
- [x] (6단계) KV 시트에서 Secret 생성 지원(`SecretName`, `SecretValue`, `Enable`, 만료/시작일 선택)
- [x] (7단계) VM AdminPassword를 Key Vault Secret 참조 방식으로 전환(`UseKeyVaultPassword`)
- [x] (8단계) DES 자동 배포 복구 및 수동/자동 병행 운영 전환

## 현재 운영 가정
- DES는 자동 배포(`DES`/`DES_PRD`) 또는 수동 생성 후 참조를 병행
- VM 시트는 `DiskEncryptionSetId` 또는 `DESName`(+선택 `DesRG`)로 CMK 연결
- VM 비밀번호는 평문(`AdminPassword`) 또는 KV Secret(`UseKeyVaultPassword=Y`) 중 선택

## 차기 후보 과제
- [ ] VM에서 `DiskEncryptionSetId`를 환경별(DEV/STG/PRD) 자동 검증하는 사전 점검 모드 추가
- [ ] KV Secret 회전 시 VM 재배포/비밀번호 갱신 운영 절차 문서화
- [ ] Private Endpoint 기반 KV 사용 시 실행 위치(Cloud Shell/점프박스/에이전트) 체크리스트 추가
