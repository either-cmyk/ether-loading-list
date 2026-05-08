---
description: 쿠팡 팔레트 적재리스트 자동 생성 (거래명세서 PDF + CN 출고내역 → XLSX/PDF)
argument-hint: <발주번호?>
---

`skills/loading-list-generator/SKILL.md` 를 따라 실행하라.

$ARGUMENTS 가 있으면 해당 발주번호로 처리. 비어있으면 사용자에게 거래명세서 PDF + CN 출고내역(S-QED) 엑셀 첨부를 요청하고 진행.

핵심 규칙:
- 트레이→파레트 매핑: G트레이 먼저 / 트레이번호 오름차순 (사용자 지정 시 그대로)
- 같은 SKU는 1줄로 합본, 혼적박스 들어간 옵션에만 박스번호 표기
- 다중 혼적박스 걸치면 콤마 표기 (예: `12-7, 12-12`)
- 거래명세서 합계와 100% 일치 검증
- 결과는 마운트 폴더에 timestamp suffix로 저장
