# TODO

## 현재 운영 기준 점검 TODO
- [x] `0~8` 레거시 스크립트 파일명에 `LEG_` 약어 적용
- [x] `98` 스크립트명을 `98.TEMP_Deploy Resource_OnlyNSG.ps1`로 정리
- [x] `99. Deploy Resources.ps1`에서 `STORAGE`, `KV`, `DES` 배포 타입 제외
- [x] LB Rule 프로토콜 컬럼 `FEProtocol`/`Protocol` 동시 허용 확인
- [x] LB_Probe `Port` 미입력 행 스킵 처리 반영
- [x] DataDisk 배포 로직을 `99` 내부로 통합(외부 8번 의존 제거)

## 운영 중점 항목
- [ ] `README.md`의 과거 스크립트명(`5. Deploy LB.ps1`, `8. Deploy DataDisk.ps1`) 잔존 문구 정리
- [ ] 샘플 Excel(운영/테스트) 기준 정기 DryRun 체크리스트(월 1회) 문서화
- [ ] `LEG_` 스크립트의 보존 기간/폐기 기준 확정

## 참고
- 현재 권장 실행 경로는 `99. Deploy Resources.ps1` 단일 진입점입니다.
- `3. LEG_Deploy KV.ps1`, `4. LEG_Deploy DES.ps1`는 정책상 미사용(명시적 차단) 상태입니다.
