# Pivo 재고 대시보드 — 팀 협업 컨텍스트 문서

> 이 문서를 Claude 새 세션에 붙여넣으면 기존 작업 맥락 그대로 이어서 진행 가능합니다.
> 작성: 2026-05-14

---

## 📌 프로젝트 개요

Pivo 제품 재고 예측 대시보드를 Python으로 생성하고 GitHub Pages로 배포하는 작업.
- 단품/번들 판매 데이터 기반 채널별(SH/AMZ/B2B) 재고 소모 예측
- 관리자 인증 후 데이터 업데이트 가능 (일반 사용자는 열람만)
- 완전 정적 HTML — 서버 없이 작동

---

## 📁 핵심 파일 위치

| 파일 | 경로 | 설명 |
|------|------|------|
| 대시보드 HTML | `/Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html` | 배포용 메인 파일 (265KB) |
| 검증 매뉴얼 | `/Users/jian/Documents/Claude/Pivo_대시보드_검증_매뉴얼.md` | 수식·검증 절차 |
| 참조 엑셀 | Google Drive ID: `1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q` | Pivo_SA_Bundle_탭별.xlsx |
| 원천 엑셀 | `/Users/jian/Documents/Claude/Pivo_재고예측시스템_v5.xlsx` | SA/BD/INV 원천 데이터 |

---

## 🌐 배포 정보

- **공개 URL**: https://johnan-oss.github.io/pivo-dashboard/
- **GitHub 저장소**: https://github.com/johnan-oss/pivo-dashboard
- **배포 방식**: GitHub Pages (main 브랜치 / index.html)

### 업데이트 시 명령어
```bash
cp /Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html /tmp/pivo-pages/index.html
cd /tmp/pivo-pages
git add index.html
git commit -m "Update dashboard"
git push
# → 약 1분 후 URL 자동 반영
```

---

## 🔐 인증 구조

- **일반 사용자**: URL 접속 → 바로 열람 가능 (로그인 불필요)
- **관리자**: 📂 데이터 업데이트 버튼 클릭 시 팝업 인증
  - 등록된 이메일: `john.an@3i.ai`, `amber.kim@3i.ai`
  - 기본 비밀번호 해시: `3i&pivo` → SHA-256
  - 슈퍼관리자 (권한관리 메뉴): `john.an@3i.ai`
- **세션 저장**: `sessionStorage['pivo_admin_session']` (브라우저 탭 닫으면 로그아웃)
- **유저 목록 저장**: `localStorage['pivo_users']` (관리자가 추가/삭제 가능)

---

## 📊 대시보드 데이터 구조

HTML 내부에 JavaScript 변수로 임베드:

```javascript
let SA   = { "SH":{}, "AMZ":{}, "B2B":{} }
// SA[채널][구성품SKU][지역][YYYYMM] = 수량

const BD = { "SH":{}, "AMZ":{}, "B2B":{} }
// BD[채널][구성품SKU][지역][YYYYMM][번들SKU] = 수량

const INV = {}
// INV[SKU]["US/3PL" or "US/AMZ"] = 현재고

const BDM = { "번들SKU": ["구성품1","구성품2",...] }
// 번들 구성 맵
```

---

## ⏱ 기준기간 및 컬럼 매핑

| 채널 | 기준기간 | 개월수 |
|------|---------|--------|
| SH / B2B | 2025-10 ~ 2026-03 | 6개월 |
| AMZ | 2026-02 ~ 2026-03 | 2개월 |

참조파일 컬럼: idx0=합계, idx1-12=2024, idx13-24=2025, idx25-27=2026-01~03
- **SH 기준기간 → idx 22~27**
- **AMZ 기준기간 → idx 26~27**

---

## 🔧 주요 수정 이력

### 2026-05-14 수정 (현재 버전)

**FIX 1 — TRPD AMZ SA 클리어**
- `SA['AMZ']['TRPD']` 전체 삭제
- 이유: TRPD는 Amazon 직접판매 재고 없음. 원천 엑셀에서 TRPD_FBA가 TRPD로 잘못 집계됨
- 효과: TRPD AMZ 월평균 71.5 → 0.5/월 (참조값 2.0/월)

**FIX 2 — STANDPRE9/10 월별 수량 수정**
- 대상: `BD['SH']['NPVB/SM/TRPD']['US']` 내 STANDPRE9, STANDPRE10 키
- 이유: 원천 엑셀에서 1개월 시프트 오류 + 특정 월 비정상 수량
  (예: 2025-11에 734.28, 2026-02에 754.28 → 각각 257, 173이 정답)
- 참조파일 기준 올바른 값으로 전체 교체
- 효과:
  - NPVB 3PL: 367.4 → 189.8/월 (참조: 190.2/월) ✅
  - TRPD 3PL: 449.4 → 271.8/월 (참조: 272.0/월) ✅
  - SM 3PL: 526.4 → 348.8/월 (참조: 343.2/월) ✅

---

## ✅ 현재 검증 상태 (US 기준)

| SKU | SH 월평균 | AMZ 월평균 | 상태 |
|-----|---------|----------|------|
| TRPD | 271.8 | 0.5 | ✅ |
| NPVB | 189.8 | 0.0 | ✅ |
| SM | 348.8 | 34.0 | ✅ |
| NPV | 2.0 | 0.0 | ✅ |
| PV-ACM2-1GC | 63.5 | 13.5 | ✅ |
| PV-I1R01 | 21.0 | 46.0 | ✅ |
| PV-AWT1-1GC | 10.2 | 0.0 | ✅ |
| PV-EWM1-1GC | 11.0 | 4.0 | ✅ |
| PV-ACC3-1GC | 19.8 | 1.0 | △ AMZ 일부 미포함 |
| NPVS | ~49.7 | ~103 | △ 소폭 차이 |

---

## 🛠 자주 쓰는 작업

### 대시보드 SA/BD 데이터 추출
```python
import re, json
with open('/Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html') as f:
    html = f.read()
SA = json.loads(re.search(r'let SA\s*=\s*(\{.+?\});', html).group(1))
BD = json.loads(re.search(r'const BD\s*=\s*(\{.+?\});', html).group(1))
```

### 참조 파일 다운로드 (MCP Drive 도구 필요)
```
fileId: 1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q
tool: mcp__d399d583-...__read_file_content
```

### 특정 SKU US SH 기준기간 합계
```python
SH_PERIOD = ['202510','202511','202512','202601','202602','202603']
sku = 'TRPD'
sa = sum(SA['SH'][sku]['US'].get(ym,0) for ym in SH_PERIOD)
bd = sum(sum(BD['SH'][sku]['US'].get(ym,{}).values()) for ym in SH_PERIOD)
print(f"{sku} SH 합계: {sa+bd:.1f}, 월평균: {(sa+bd)/6:.1f}/월")
```

---

## 📌 향후 작업 가이드

1. **데이터 업데이트**: 엑셀 파일 업데이트 → gen_dashboard_v3.py 재실행 → HTML 생성 → GitHub push
2. **새 번들 추가**: HTML 내 `const BDM = {...}` 수정 후 재생성
3. **TRPD 주의**: 엑셀 원천에서 TRPD_FBA ≠ TRPD — 반드시 분리
4. **STANDPRE 주의**: 월별 컬럼 오프셋 오류 이력 있음 — 참조파일과 반드시 대조
5. **검증 실행**: `/Users/jian/Documents/Claude/Pivo_대시보드_검증_매뉴얼.md` 참조
