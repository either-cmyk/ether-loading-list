# 변경 이력

## v1.0.0 - 2026-05-08

### 최초 배포 (플러그인화)

- `loading-list-generator` 스킬을 독립 플러그인으로 분리
- v2.0 합본 방식 적용:
  - 같은 SKU 1줄 합본 (단일+혼적박스 모두 등장 시)
  - 다중 혼적박스 콤마 표기 (예: `12-7, 12-12`)
  - declared_box_count 사용자 지정 지원
  - 자동 혼적 감지 + 화이트리스트 합집합
- 거래명세서 vs S-QED 수량 검증 + ADJUSTMENTS 자동 보정
- 파레트별 1페이지(PDF) / 1시트(XLSX) 동시 생성
- 마운트 폴더에 timestamp suffix 저장 (덮어쓰기 X)
- after-install.sh로 memory 4개 자동 복사
- 폰트: NotoSansKR (Regular/Bold)
- 메모리 4개 통합:
  - `loading_list_format_rules.md` (쿠팡 신양식 필드 매핑)
  - `loading_list_consolidation.md` (합본 방식 v2.0)
  - `loading_list_tray_mixed.md` (트레이/혼적박스 처리, B-004/G-079 케이스)
  - `save_to_mounted_folder.md` (마운트 폴더 timestamp 저장)

### 검증된 발주 (이 플러그인으로 처리한 실제 케이스)

- 130185106 (1파레트 14박스) — G-014 혼적
- 130379426 (13파레트 204박스) — G-079, B-067, B-119, B-124 혼적
- 130490638 (2파레트 28박스) — B-119, B-124 혼적
- 130491856 (11파레트 176박스) — G-079, B-067 혼적
- 130826054 (5파레트 79박스) — 12개 혼적박스 케이스 (가장 복잡)
