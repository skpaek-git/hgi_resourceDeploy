# PLAN

## 목적
- 현재 배포 운영을 `99. Deploy Resources.ps1` 단일 진입점으로 유지한다.
- Excel 시트 변경이 스크립트 동작에 미치는 영향을 사전에 최소화한다.

## 현행 배포 범위(2026-05-22 기준)
- 포함: `RG`, `VNET`, `UDR`, `LB`, `LB_PROBE`, `LB_RULE`, `VM`, `DATADISK`, `NSG`
- 제외: `STORAGE`, `KV`, `DES`

## 현행 실행/의존 구조
- 메인: `99. Deploy Resources.ps1`
- 래퍼(선택): `5. LEG_Deploy LB.ps1` -> 내부에서 `99` 호출
- DataDisk: `99` 내부 함수(`Deploy-DataDisks`)로 처리, 외부 `8` 의존 없음
- 레거시: `0~8 LEG_*` 파일은 과거 단독 실행 호환 목적 보관

## 운영 계획
- [x] 파일명 체계 정리(`LEG_`, `TEMP_Deploy Resource_OnlyNSG`)
- [x] DataDisk 8->99 통합 및 동작 검증(DryRun)
- [x] LB Probe 포트 누락 스킵 처리 검증
- [x] 이름 변경 후 호출 오류 재점검
- [ ] 레거시 스크립트 폐기 시점 확정(운영 합의 필요)
- [ ] 정기 검증 자동화(예: 주간 DryRun 배치) 여부 결정

## 검증 기준
- 동일 Excel 기준 `99 -DeployType DATADISK -DryRun` 정상 종료(code 0)
- LB 단독/래퍼 실행 정상 종료(code 0)
- 시트 컬럼 변경(`FEProtocol` -> `Protocol`)에도 오류 없이 처리
