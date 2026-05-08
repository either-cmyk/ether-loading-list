---
name: 적재리스트 출력은 마운트 폴더에 timestamp suffix
description: 사용자가 선택한 폴더에 timestamp 포함하여 저장 (덮어쓰기 X)
type: feedback
---

XLSX/PDF 결과 파일은 outputs 폴더가 아닌 **사용자가 mount한 작업 폴더**에 저장.

**Why:** outputs 폴더는 사용자가 직접 접근 못함. 마운트 폴더에 저장해야 사용자가 열 수 있음.

**How to apply:**
- 파일명: `적재리스트_{발주번호}_{YYYYMMDD_HHMMSS}.xlsx`
- 동일 발주번호 재생성 시에도 timestamp suffix로 새 파일 생성 (덮어쓰기 X)
- 사용자가 실수로 삭제 후 재요청 시에도 안전
- `computer://` 링크로 사용자에게 직접 안내
