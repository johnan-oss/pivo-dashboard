# Pivo 재고 대시보드 — Claude 협업 컨텍스트
> 이 문서를 새 Claude 세션 첫 메시지로 붙여넣으면 기존 작업을 바로 이어서 진행할 수 있습니다.
> 최종 업데이트: 2026-05-14

---

## 1. 한 줄 요약

Pivo(카메라 자동 추적 스탠드 회사)의 SKU별 재고 소모 예측 대시보드를 Python으로 생성하고 GitHub Pages에 배포하는 프로젝트입니다. 단품(SA 직접판매) + 번들(Bundle 기여) 판매 데이터를 합산해 채널별(3PL/AMZ/B2B) 월평균 소모량을 계산하고, 재고 소진 시점을 예측합니다.

---

## 2. 파일 위치

| 용도 | 경로 |
|------|------|
| **대시보드 HTML** (메인) | `/Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html` |
| 원천 데이터 엑셀 | `/Users/jian/Documents/Claude/Pivo_재고예측시스템_v5.xlsx` |
| 검증 매뉴얼 | `/Users/jian/Documents/Claude/Pivo_대시보드_검증_매뉴얼.md` |
| 참조 파일 (Google Drive) | File ID: `1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q` |
| 생성기 스크립트 | `/tmp/gen_dashboard_v3.py` (세션 종료시 삭제됨 — 재생성 필요) |

---

## 3. 배포 정보

- **공개 URL**: https://johnan-oss.github.io/pivo-dashboard/
- **GitHub Repo**: https://github.com/johnan-oss/pivo-dashboard
- **계정**: johnan-oss (john.an@3i.ai)

### 업데이트 절차
```bash
cp /Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html /tmp/pivo-pages/index.html
cd /tmp/pivo-pages
git add index.html && git commit -m "Update $(date +%Y-%m-%d)" && git push
# 약 60초 후 URL 반영
```

---

## 4. 대시보드 HTML 구조

HTML은 완전 자기완결형(self-contained)으로, 내부에 JS 변수로 모든 데이터가 임베드됩니다.

```javascript
let SA   = { "SH":{...}, "AMZ":{...}, "B2B":{...} }
// SA[채널][구성품SKU][지역][YYYYMM] = 수량 (직접판매)

const BD = { "SH":{...}, "AMZ":{...}, "B2B":{...} }
// BD[채널][구성품SKU][지역][YYYYMM][번들SKU] = 수량 (번들 기여)

const INV = {}
// INV[SKU]["US/3PL" or "US/AMZ"] = 현재고 수량

const BDM = { "번들SKU": ["구성품1","구성품2",...] }
// 번들 분해 맵 (display용)
```

### Python으로 SA/BD 추출
```python
import re, json
with open('/Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html') as f:
    html = f.read()
SA  = json.loads(re.search(r'let SA\s*=\s*(\{.+?\});', html).group(1))
BD  = json.loads(re.search(r'const BD\s*=\s*(\{.+?\});', html).group(1))
INV = json.loads(re.search(r'const INV\s*=\s*(\{.+?\});', html).group(1))
```

---

## 5. 계산 로직

### 기준기간
| 채널 | 기간 | 개월 |
|------|------|------|
| SH (Shopify/3PL) | 2025-10 ~ 2026-03 | 6 |
| B2B | 2025-10 ~ 2026-03 | 6 |
| AMZ (Amazon) | 2026-02 ~ 2026-03 | 2 |

### 월평균 수요 공식
```
SH 월평균  = (SA['SH'][sku]['US'] 기준기간 합 + BD['SH'][sku]['US'] 기준기간 합) / 6
AMZ 월평균 = (SA['AMZ'][sku]['US'] 기준기간 합 + BD['AMZ'][sku]['US'] 기준기간 합) / 2
```

### 예측 공식
```
월별 예측 수요 = 월평균 × 계절성지수[월] × (1 + 성장률)^n
3PL 잔여재고 = 현재고 - Σ(SH예측 + B2B예측)   ← 3PL 창고 소모
AMZ 잔여재고 = 현재고 - Σ(AMZ예측)             ← Amazon FBA 소모
```

---

## 6. 참조 파일 컬럼 매핑 (검증 시 핵심)

참조 파일(`Pivo_SA_Bundle_탭별.xlsx`)의 각 데이터 행:
```
(SA 직접판매|Bundle 기여), (SH|AMZ|B2B), [번들SKU], [상품명], [구성설명],
[idx0=합계], [idx1=2024-01], ..., [idx12=2024-12],
[idx13=2025-01], ..., [idx24=2025-12],
[idx25=2026-01], [idx26=2026-02], [idx27=2026-03], [idx28=2026-04...]
```

- **SH 기준기간 → idx 22~27** (2025-10 ~ 2026-03)
- **AMZ 기준기간 → idx 26~27** (2026-02 ~ 2026-03)

---

## 7. 현재 US 기준 월평균 (2026-05-14 기준)

| SKU | SH 월평균 | AMZ 월평균 |
|-----|---------|----------|
| SM                   |  348.8 |   34.0 |
| TRPD                 |  271.8 |    0.5 |
| NPVB                 |  189.8 |    0.0 |
| NPVS                 |   49.7 |  103.0 |
| NPV                  |    2.0 |    0.0 |
| TC                   |   84.2 |   28.0 |
| PV-PMBK-1GK          |   71.0 |   21.5 |
| PV-ACM2-1GC          |   63.5 |   13.5 |
| PV-I1R01             |   21.0 |   46.0 |
| PV-ACC3-1GC          |   19.8 |    1.0 |
| PV-AWT1-1GC          |   10.2 |    0.0 |
| PV-EWM1-1GC          |   11.0 |    4.0 |

---

## 8. 인증 구조

- **열람**: 로그인 없이 누구나 가능
- **데이터 업데이트**: 관리자 팝업 인증 필요
  - 등록 이메일: `john.an@3i.ai`, `amber.kim@3i.ai`
  - 기본 비밀번호: `3i&pivo` (SHA-256 해시로 저장)
  - 슈퍼관리자(권한관리 메뉴): `john.an@3i.ai`
- 인증 세션: `sessionStorage['pivo_admin_session']` → 탭 닫으면 자동 로그아웃
- 유저 목록: `localStorage['pivo_users']` → 관리자가 추가/삭제 가능

---

## 9. 주요 수정 이력

### FIX-1 (2026-05-14): TRPD AMZ SA 오류 수정
- **문제**: 엑셀 원천에서 TRPD_FBA(Amazon 별도 SKU) 데이터가 TRPD AMZ 직판으로 잘못 집계
  → TRPD AMZ 월평균 71.5/월로 표시됨 (실제: ~2.0/월)
- **수정**: `SA['AMZ']['TRPD']` 전체 클리어
- **결과**: 71.5/월 → 0.5/월 (참조값 2.0/월에 근접) ✅

### FIX-2 (2026-05-14): STANDPRE9/10 월별 수량 수정
- **문제**: 원천 엑셀에서 STANDPRE9(NPVB+SM+TRPD 번들) 월별 데이터가
  1개월 앞당겨 기록되고, 특정 월에 비정상 수량 발생
  (2025-11: 734.28 → 실제 257, 2026-02: 754.28 → 실제 173)
- **수정 범위**: `BD['SH']['NPVB/SM/TRPD']['US']` 내 STANDPRE9/10 전체 교체
- **수정 후 값 (STANDPRE9 US SH)**:
  - 202506=18, 202507=234, 202508=25, 202509=240
  - 202510=236, 202511=257, 202512=199
  - 202601=42, 202602=173, 202603=2
- **수정 후 값 (STANDPRE10 US SH)**: 202603=219
- **결과**:
  - NPVB: 367.4 → 189.8/월 (참조 190.2/월) ✅
  - TRPD: 449.4 → 271.8/월 (참조 272.0/월) ✅
  - SM: 526.4 → 348.8/월 (참조 343.2/월) ✅

---

## 10. 주요 번들 구성 (STANDPRE 계열)

```python
{
  "GiftboxSP_FBA": [
    "NPV",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "GiftboxSPS_FBA": [
    "NPVS",
    "SM",
    "TC"
  ],
  "GSPR": [
    "NPV",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "GSPS": [
    "NPVS",
    "SM",
    "TC"
  ],
  "GSTANDPR": [
    "NPV",
    "PV-I1R01",
    "SM",
    "TC",
    "TRPD"
  ],
  "GSTANDPS": [
    "NPVS",
    "SM",
    "TC",
    "TRPD"
  ],
  "STANDPR": [
    "PV",
    "SM",
    "TC",
    "TRPD"
  ],
  "STANDPS": [
    "NPVS",
    "SM",
    "TC",
    "TRPD"
  ],
  "STANDP": [
    "NPV",
    "SM",
    "TC",
    "TRPD"
  ],
  "STANDP_FBA": [
    "NPV",
    "SM",
    "TC",
    "TRPD"
  ],
  "STANDPS_FBA": [
    "NPVS",
    "SM",
    "TC",
    "TRPD"
  ],
  "PV-BL0101": [
    "PV-P1L01",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "PV-BL0201": [
    "PV-P1L02",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "PV-BL0301": [
    "PV-P1L03",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "PV-BL0401": [
    "PV-P1L04",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "PV-BL0501": [
    "PV-P1L05",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "PV-BL0608": [
    "PV-P1L06",
    "PV-I1R01",
    "SM",
    "TC"
  ],
  "PV-BL0617": [
    "PV-P1L06",
    "PV-I1R01",
    "SM",
    "TRPD",
    "PV-ACC3-1GC"
  ],
  "PV-BL1006": [
    "PV-P1L01",
    "PV-I1R02",
    "SM",
    "TC"
  ],
  "PV-BL1007": [
    "PV-P1L04",
    "PV-I1R02",
    "SM",
    "TC"
  ],
  "PV-BL1008": [
    "PV-P1L03",
    "PV-I1R02",
    "SM",
    "TC"
  ],
  "PV-BL1009": [
    "PV-P1L05",
    "PV-I1R02",
    "SM",
    "TC"
  ],
  "PV-BL1011": [
    "PV-P1L02",
    "PV-I1R02",
    "SM",
    "TC"
  ],
  "PV-BM0112": [
    "PV-PMBK-1GK",
    "SM",
    "TRPD",
    "PV-ACC3-1GC"
  ],
  "PV-BM0101": [
    "PV-PMBK-1GK",
    "SM",
    "PV-ACM2-1GC"
  ],
  "PV-BM0102": [
    "PV-PMBK-1GK",
    "SM",
    "PV-ACM2-1GC",
    "TRPD"
  ],
  "PV-BM0119": [
    "PV-PMBK-1GK",
    "SM",
    "TRPD",
    "PV-AWT1-1GC",
    "PV-ACC3-1GC"
  ],
  "PV-BM0123": [
    "PV-PMBK-1GK",
    "SM",
    "TRPD",
    "PV-ACC3-1GC"
  ],
  "PV-BM0133": [
    "PV-PMBK-1GK",
    "SM",
    "PV-ACM2-1GC"
  ],
  "PV-BS0101": [
    "NPVS",
    "SM",
    "TC"
  ],
  "PV-BS0102": [
    "NPVS",
    "SM",
    "TC",
    "TRPD"
  ],
  "PV-BS0127": [
    "NPVS",
    "TC",
    "SM",
    "TRPD",
    "PV-AWT1-1GC",
    "PV-EWM1-1GC"
  ],
  "PV-BS0129": [
    "NPVS",
    "SM",
    "TRPD",
    "PV-ACC3-1GC",
    "PV-AWT1-1GC"
  ],
  "PV-BS0130": [
    "NPVS",
    "SM",
    "PV-AWT1-1GC",
    "PV-EWM1-1GC",
    "PV-ACC3-TRPD"
  ],
  "PV-BS0133": [
    "NPVS",
    "SM",
    "TC",
    "TRPD",
    "PV-EWM1-1GC"
  ],
  "PV-BS0135": [
    "NPVS",
    "SM",
    "PV-ACC3-TRPD"
  ],
  "PV-BW0101": [
    "PV-P1W03",
    "PV-I1R02",
    "PV-A1TD2",
    "SM"
  ],
  "PV-BW0104": [
    "PV-P1W01",
    "SM",
    "PV-EVL1-1GC",
    "PV-ERC2-1GK",
    "PV-ACM2-1GC"
  ],
  "STANDPRE9": [
    "NPVB",
    "SM",
    "TRPD"
  ],
  "STANDPRE10": [
    "NPVB",
    "SM",
    "TRPD"
  ],
  "STANDPRE11": [
    "PV-PMBK-1GK",
    "SM",
    "TRPD"
  ],
  "STANDPRE12": [
    "PV-PMBK-1GK",
    "SM",
    "TRPD"
  ]
}
```

---

## 11. 잔여 과제 (향후 진행 필요)

| 항목 | 내용 | 우선순위 |
|------|------|---------|
| PV-ACC3-1GC AMZ | 참조값 12.5/월 vs 대시보드 1.0/월. PV-BM0111 AMZ 번들 기여 누락 | 중 |
| 생성기 스크립트 재작성 | `/tmp/gen_dashboard_v3.py` — 세션 종료로 소실됨. 재구성 필요 | 중 |
| EU/CA/UK 검증 | 현재 US만 검증 완료. 다른 지역 대조 미완 | 낮음 |
| NPVS/TC 소폭 차이 | 각각 +3~5/월 차이. 번들 구성 방식 차이로 추정 | 낮음 |

---

## 12. 작업 시 유용한 스니펫

### 특정 SKU 기준기간 월평균 계산
```python
SH_P  = ['202510','202511','202512','202601','202602','202603']
AMZ_P = ['202602','202603']

def monthly_avg(sku, ch, region='US'):
    period = AMZ_P if ch == 'AMZ' else SH_P
    n = len(period)
    sa = sum(SA.get(ch,{}).get(sku,{}).get(region,{}).get(ym,0) for ym in period)
    bd = sum(sum(BD.get(ch,{}).get(sku,{}).get(region,{}).get(ym,{}).values()) for ym in period)
    return (sa+bd)/n

print(monthly_avg('TRPD', 'SH'))   # → 271.8
print(monthly_avg('TRPD', 'AMZ'))  # → 0.5
```

### 검증 스크립트 (참조파일 vs 대시보드 비교)
```python
# 참조파일 다운로드 (MCP 도구 사용)
# tool: mcp__d399d583-ed07-434b-af3b-9e2400a25d4a__read_file_content
# fileId: 1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q

# 참조파일 합계,US 행 파싱 → VERIFICATION_MANUAL.md 참조
```

### BD 내 특정 번들 기여 확인
```python
sku, ch, region, ym = 'NPVB', 'SH', 'US', '202511'
print(BD[ch][sku][region][ym])
# 예: {'STANDPRE9': 257.0}  ← 수정 후
```

---

## 13. 작업 환경

- **OS**: macOS (arm64)
- **Python**: 3.x + 표준 라이브러리 (openpyxl 별도 설치 시 엑셀 파싱 가능)
- **Git 로컬**: `/tmp/pivo-pages/` (GitHub push 워킹 디렉터리)
- **MCP 도구**: Google Drive 연동 (`mcp__d399d583-ed07-434b-af3b-9e2400a25d4a__*`)
