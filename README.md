# ether-loading-list

쿠팡 팔레트 적재리스트 자동 생성 플러그인 (Claude Code 마켓플레이스).

## 대상 사용자

이더컴퍼니 및 계열 사업자 (이더컴퍼니/뉴트리정/마인플로/클린인테크/이든코퍼레이션) 입고 운영자.

## 기능 개요

`loading-list-generator` 스킬이 다음을 자동화:

1. 거래명세서 PDF 파싱 (업체명, 발주번호, 센터, 도착예정일, 팔레트수, SKU별 확정수량)
2. CN인사이더 출고내역(S-QED) 엑셀 파싱 (트레이별 박스 + 혼적박스 자동감지)
3. 합본 방식으로 파레트별 행 생성 (같은 SKU 1줄, 혼적 박스번호만 표기)
4. 거래명세서 vs 실측 수량 검증 + 자동 보정
5. 파레트별 1페이지(PDF) / 1시트(XLSX) 생성

## v2.0 핵심 변경점

- **합본 방식**: 같은 SKU가 단일+혼적박스에 모두 있으면 1줄로 합산, 박스번호만 표기
- **다중 혼적 콤마 표기**: 한 SKU가 여러 혼적박스에 걸치면 `12-7, 12-12` 형식
- **declared_box_count**: 사용자가 박스수 더 크게 등록한 경우 그 값 우선 (박스번호도 declared 기준)
- **자동 혼적 감지 + 화이트리스트 합집합**: 같은 (트레이, 박스코드) 다중행 자동감지 + 사용자 명시 추가 가능

## 폴더 구조

```
ether-loading-list/
├── .claude-plugin/
│   ├── marketplace.json     # 마켓플레이스 메타
│   └── plugin.json          # 플러그인 메타
├── skills/
│   └── loading-list-generator/
│       ├── SKILL.md
│       ├── scripts/
│       │   └── build_loading_list.py
│       └── assets/
│           ├── NotoSansKR-Regular.ttf
│           └── NotoSansKR-Bold.ttf
├── memory/
│   ├── MEMORY.md
│   ├── loading_list_format_rules.md
│   ├── loading_list_consolidation.md
│   ├── loading_list_tray_mixed.md
│   └── save_to_mounted_folder.md
├── commands/
│   └── loading-list.md
├── hooks/
│   └── after-install.sh     # 설치 시 memory/ 파일을 사용자 메모리 폴더에 복사
├── CHANGELOG.md
└── README.md
```

## 설치 방법 (PowerShell 1회 실행)

```powershell
$pat_b64 = "Z2l0aHViX3BhdF8xMUNDUDNVWlkwZEJZTXlNWmkzS0dFX3JCdXBaTVhxUDhabWJwWVNQNk9STGpkR1M5SkhrcXR1UktBb2tkcDNzaFk1U05YV1lTRHg1c0VNdnIw"
$pat = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($pat_b64))

# 마켓플레이스 등록
claude plugin marketplace add "https://x-access-token:$pat@github.com/either-cmyk/ether-loading-list.git"

# 플러그인 설치
claude plugin install ether-loading-list
```

설치 후 Cowork Desktop을 완전 종료(트레이 Quit) 후 재시작.

## 업데이트

```
/plugin marketplace update
/plugin update ether-loading-list
```

## 사용

거래명세서 PDF + S-QED 엑셀을 함께 첨부하고 "적재리스트 만들어줘" 라고 입력하면 자동 실행.

또는 슬래시 커맨드:

```
/loading-list 130826054
```

## 업데이트 흐름 (개발자용)

### PC1에서 변경

```bash
cd /path/to/ether-loading-list
# 파일 수정
git add -A
git commit -m "..."
git push
```

### PC2에서 수령

```
/plugin marketplace update
/plugin update ether-loading-list
```

after-install.sh가 자동 실행되어 memory/ 4개 파일이 사용자 Claude 메모리 폴더로 복사됨.
