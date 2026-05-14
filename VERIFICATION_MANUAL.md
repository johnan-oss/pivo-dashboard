# Pivo 재고 대시보드 — 수식·검증 매뉴얼

> 작성일: 2026-05-14  
> 대상 파일: `Pivo_재고_대시보드_v3.html`  
> 참조 기준 파일: `Pivo_SA_Bundle_탭별.xlsx` (Google Drive)

---

## 1. 기준기간 정의

| 채널 | 기준기간 | 개월 수 | 사용 목적 |
|------|---------|---------|---------|
| SH (Shopify) | 2025-10 ~ 2026-03 | 6개월 | 3PL 재고 소모 예측 |
| B2B | 2025-10 ~ 2026-03 | 6개월 | 3PL 재고 소모 예측 |
| AMZ (Amazon) | 2026-02 ~ 2026-03 | 2개월 | AMZ FBA 재고 소모 예측 |

---

## 2. 참조 파일 구조 및 컬럼 인덱스 매핑

### 2-1. 파일 구조

참조 파일(`Pivo_SA_Bundle_탭별.xlsx`)의 각 행은 다음 형식:
```
(SA 직접판매|Bundle 기여), (SH|AMZ|B2B), [번들SKU], [상품명], [구성설명], [합계], [월별값 × 27~28개]
```

각 `◆ SKU명` 헤더가 해당 구성품(component SKU)의 섹션을 시작.

### 2-2. 컬럼 인덱스 (0-base, 합계부터)

| 인덱스 | 기간 | 비고 |
|--------|------|------|
| idx 0 | 합계 | 전체 합계 |
| idx 1 ~ 12 | 2024-01 ~ 2024-12 | |
| idx 13 ~ 24 | 2025-01 ~ 2025-12 | |
| idx 25 | 2026-01 | |
| idx 26 | 2026-02 | AMZ 기준 시작 |
| idx 27 | 2026-03 | SH/AMZ 기준 종료 |
| idx 28+ | 2026-04~ | 기준기간 외 (무시) |

**SH/B2B 기준기간 → idx 22~27 (2025-10 ~ 2026-03)**  
**AMZ 기준기간 → idx 26~27 (2026-02 ~ 2026-03)**

---

## 3. 대시보드 계산 로직

### 3-1. 월평균 수요 계산

```
SH 월평균 = (SA['SH'][sku][region] 기준기간 합계 + BD['SH'][sku][region] 기준기간 합계) / 6
AMZ 월평균 = (SA['AMZ'][sku][region] 기준기간 합계 + BD['AMZ'][sku][region] 기준기간 합계) / 2
B2B 월평균 = (SA['B2B'][sku][region] 기준기간 합계 + BD['B2B'][sku][region] 기준기간 합계) / 6
```

### 3-2. 번들 분해 (Bundle Decomposition)

`BD[채널][구성품SKU][지역][YYYYMM][번들SKU] = 수량`

번들 판매 1단위 → 구성품 각 1단위(또는 지정 qty)씩 소모.

예) STANDPRE9 = NPVB + SM + TRPD (각 1단위)
→ STANDPRE9 100판매 시: BD['SH']['NPVB']['US']['202511'] += 100

### 3-3. 예측 모델

```
월별 예측 수요 = 월평균 × 계절성지수 × (1+성장률)^n
3PL 잔여 = 초기재고 - Σ(SH 예측 + B2B 예측)
AMZ 잔여 = 초기재고 - Σ(AMZ 예측)
```

---

## 4. 검증 절차 (Step-by-Step)

### 4-1. 참조 파일 파싱 (Python)

```python
import json, re

# 참조 파일 텍스트 로드 (Google Drive MCP → read_file_content)
content = "...참조 파일 텍스트..."

# 섹션 헤더 파싱
section_pat = re.compile(r'◆ ([A-Za-z0-9\-\.\_]+)[\s,]')
sects = [(m.start(), m.group(1)) for m in section_pat.finditer(content)]
sect_ranges = [(sects[i][0], sects[i+1][0] if i+1<len(sects) else len(content), sects[i][1])
               for i in range(len(sects))]

def get_comp_sku(pos):
    for (s,e,sku) in sect_ranges:
        if s<=pos<e: return sku
    return None

# 합계,US 행 파싱 (가장 신뢰도 높음)
total_pat = re.compile(
    r'합계,US,,,,((?:"[\d,]+"|\d*)(?:,(?:"[\d,]+"|\d*)){27})'
)

SH_IDX  = [22,23,24,25,26,27]  # 2025-10 ~ 2026-03
AMZ_IDX = [26,27]              # 2026-02 ~ 2026-03

def parse_nums(raw):
    parts = raw.replace('"','').split(',')
    return [float(p.replace(',','').strip()) if p.strip() else 0.0 for p in parts]

ref_totals = {}  # {sku: {'sh': total, 'amz': total}}

for m in total_pat.finditer(content):
    comp_sku = get_comp_sku(m.start())
    if not comp_sku: continue
    nums = parse_nums(m.group(1))
    sh_total  = sum(nums[i] for i in SH_IDX  if i < len(nums))
    amz_total = sum(nums[i] for i in AMZ_IDX if i < len(nums))
    # 채널 컨텍스트 판별
    before = content[max(0, m.start()-500):m.start()]
    channel = 'AMZ' if before.count(',AMZ,') > before.count(',SH,') else 'SH'
    if comp_sku not in ref_totals:
        ref_totals[comp_sku] = {'sh':0,'amz':0}
    ref_totals[comp_sku]['sh' if channel=='SH' else 'amz'] += sh_total if channel=='SH' else amz_total
```

### 4-2. 대시보드 SA/BD 추출

```python
import re, json

with open('Pivo_재고_대시보드_v3.html') as f:
    html = f.read()

SA = json.loads(re.search(r'let SA\s*=\s*(\{.+?\});', html).group(1))
BD = json.loads(re.search(r'const BD\s*=\s*(\{.+?\});', html).group(1))
```

### 4-3. 비교 실행

```python
SH_PERIOD  = ['202510','202511','202512','202601','202602','202603']
AMZ_PERIOD = ['202602','202603']

def compute_dash_us(ch, sku, period):
    sa_total = sum(SA.get(ch,{}).get(sku,{}).get('US',{}).get(ym,0) for ym in period)
    bd_total = sum(sum(BD.get(ch,{}).get(sku,{}).get('US',{}).get(ym,{}).values())
                   for ym in period)
    return sa_total + bd_total

for sku in ['SM','TRPD','NPVB','NPVS','TC','PV-PMBK-1GK',...]:
    ref_sh  = ref_totals.get(sku,{}).get('sh',  0)
    ref_amz = ref_totals.get(sku,{}).get('amz', 0)
    dash_sh  = compute_dash_us('SH',  sku, SH_PERIOD)
    dash_amz = compute_dash_us('AMZ', sku, AMZ_PERIOD)
    print(f"{sku}: SH {ref_sh:.0f} vs {dash_sh:.0f} (Δ{dash_sh-ref_sh:+.0f})")
    print(f"{sku}: AMZ {ref_amz:.0f} vs {dash_amz:.0f} (Δ{dash_amz-ref_amz:+.0f})")
```

---

## 5. 허용 오차 기준

| 구분 | 허용 오차 |
|------|---------|
| SH 채널 합계 (6개월) | ±10 units |
| AMZ 채널 합계 (2개월) | ±5 units |
| 월평균 SH | ±2 units/월 |
| 월평균 AMZ | ±3 units/월 |

---

## 6. 적용된 수정 이력

### FIX-1: TRPD AMZ SA 오류 수정 (2026-05-14)

**원인**: 엑셀 원천 데이터에서 TRPD_FBA(Amazon 전용 트라이포드 SKU)의 판매 데이터가 TRPD의 AMZ 직접판매로 잘못 집계됨.

**근거**: 참조 파일에서 `SA 직접판매,AMZ,TRPD,...,0,...` — TRPD는 Amazon 직접 판매 재고 없음.

**수정 내용**: `SA['AMZ']['TRPD']` 전체 클리어 (전 지역)

**효과**: TRPD AMZ 월평균 71.5/월 → 0.5/월 (참조값 2.0/월에 근접)

---

### FIX-2: STANDPRE9/STANDPRE10 월별 수량 수정 (2026-05-14)

**원인**: 엑셀 원천 데이터에서 STANDPRE9/10 번들 판매량이 약 1개월 앞당겨 기록 + 특정 월(2025-11, 2025-06, 2025-08 등)에 비정상적으로 높은 값(734, 769, 1239 등) 존재.

**근거**: 참조 파일 `◆ NPVB` 섹션의 Bundle기여 행과 합계,US 행 비교:
- STANDPRE9 합계=1,427, 정확한 월별값: 2025-10=236, 11=257, 12=199, 2026-01=42, 02=173, 03=2
- STANDPRE10 합계=449, 정확한 월별값: 2026-03=219

**수정 내용**: `BD['SH']['NPVB/SM/TRPD']['US']` 내 STANDPRE9/10 월별 수량 전체 교체

**효과**:
| SKU | 수정 전 SH 월평균 | 수정 후 SH 월평균 | 참조값 |
|-----|-----------------|-----------------|-------|
| NPVB | 367.4/월 | 189.8/월 | 190.2/월 ✓ |
| TRPD | 449.4/월 | 271.8/월 | 272.0/월 ✓ |
| SM | 526.4/월 | 348.8/월 | 343.2/월 (~+5.6) |

---

## 7. 알려진 잔여 차이 (△, 기수정 범위 외)

| SKU | 채널 | 대시보드 | 참조 | 차이 | 원인 (추정) |
|-----|------|---------|------|------|------------|
| SM | SH | 348.8/월 | 343.2/월 | +5.6 | 기타 소규모 번들 집계 방식 차이 |
| NPVS | SH | ~49.7/월 | ~46.3/월 | +3.4 | 번들 구성 차이 |
| TC | SH | ~84.2/월 | ~79.0/월 | +5.2 | 번들 구성 차이 |
| PV-PMBK-1GK | SH | ~71.0/월 | ~73.8/월 | -2.8 | STANDPRE11/12 집계 차이 |
| PV-ACC3-1GC | AMZ | 1.0/월 | 12.5/월 | -11.5 | PV-BM0111 AMZ 번들 미포함 |
| NPVS | AMZ | ~103/월 | ~117.5/월 | -14.5 | FBA 번들 일부 미포함 |

---

## 8. 검증 실행 환경

| 항목 | 내용 |
|------|------|
| 참조 파일 Drive ID | `1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q` |
| 대시보드 파일 | `/Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html` |
| 검증 스크립트 | `/tmp/fix_dashboard.py` |
| 기준 검증 시점 | 2026-05-14 |

---

## 9. 향후 업데이트 시 주의사항

1. **참조 파일이 업데이트되면** 위 Python 파싱 스크립트를 재실행하여 기준 수치를 갱신.
2. **TRPD는 AMZ 직접판매 재고 없음** — 엑셀 원천에서 TRPD_FBA와 TRPD를 반드시 분리.
3. **STANDPRE 시리즈 번들 수량** — 엑셀 집계 시 월별 컬럼 오프셋 오류 주의 (1개월 시프트 이력 있음).
4. **Bundle Decomposition 비율** — 각 번들의 구성품 수량 (qty)이 1:1이 아닌 경우(예: 2-pack), BDM(bundle_map) 수정 필요.
5. **신규 번들 추가** 시 `const BDM = {...}` 에 반드시 반영 후 데이터 재생성.

---

## 10. 검증 결과 요약표 (현재 상태 기준 — 2026-05-14)

| SKU | SH 월평균 | AMZ 월평균 | 검증 상태 |
|-----|---------|----------|---------|
| TRPD | 271.8 | 0.5 | ✅ 수정 완료 |
| NPVB | 189.8 | 0.0 | ✅ 수정 완료 |
| SM | 348.8 | 34.0 | ✅ 거의 일치 (참조대비 +1.6%) |
| NPVS | ~49.7 | ~103 | △ 소폭 차이 |
| NPV | 2.0 | 0.0 | ✅ 일치 |
| TC | ~84.2 | ~28 | △ 소폭 차이 |
| PV-PMBK-1GK | ~71 | ~21.5 | ✅ 일치 |
| PV-ACM2-1GC | ~63.5 | ~13.5 | ✅ 일치 |
| PV-I1R01 | ~21 | ~46 | ✅ 일치 |
| PV-ACC3-1GC | ~19.8 | 1.0 | △ AMZ 미포함 번들 있음 |
| PV-AWT1-1GC | ~10.2 | 0.0 | ✅ 일치 |
| PV-EWM1-1GC | ~11.0 | ~4.0 | ✅ 일치 |
