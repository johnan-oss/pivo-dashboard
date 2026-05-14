# Pivo 재고 대시보드 — 자동 업데이트 매뉴얼

> 작성일: 2026-05-15  
> 버전: v1.0  
> 대상: 대시보드 관리자 (john.an@3i.ai, AMBKim)

---

## 목차

1. [개요 — 업데이트 방식 비교](#1-개요)
2. [사전 준비](#2-사전-준비)
3. [방법 A — 수동 업데이트 (즉시 사용 가능)](#3-방법-a--수동-업데이트)
4. [방법 B — 반자동 (GitHub Actions 월간 스케줄)](#4-방법-b--반자동-github-actions)
5. [방법 C — 완전 자동 (Google Sheets 직접 연동)](#5-방법-c--완전-자동-google-sheets-연동)
6. [데이터 형식 가이드 — Excel vs Google Sheets](#6-데이터-형식-가이드)
7. [월간 업데이트 체크리스트](#7-월간-업데이트-체크리스트)
8. [문제 해결 (Troubleshooting)](#8-문제-해결)

---

## 1. 개요

### 업데이트 방식 비교

| 방식 | 난이도 | 자동화 수준 | 소요 시간 | 권장 대상 |
|------|--------|------------|----------|----------|
| **A. 수동** | ⭐ 쉬움 | 없음 (수동 실행) | 10~15분 | 즉시 시작, 학습 목적 |
| **B. 반자동** | ⭐⭐ 보통 | GitHub Actions 스케줄 | 초기 설정 2시간, 이후 0분 | **권장** — 월간 자동 배포 |
| **C. 완전자동** | ⭐⭐⭐ 높음 | Google Sheets 실시간 연동 | 초기 설정 4시간+ | 고급 사용자 |

**권장 경로**: 방법 A로 흐름 파악 → 방법 B 설정하여 월간 자동화

---

## 2. 사전 준비

### 2-1. Google Drive 파일 구조 설정

Google Drive에 다음과 같이 폴더를 구성하세요:

```
📁 Pivo_Dashboard_Data/
├── 📄 Pivo_재고예측시스템_v5.xlsx    ← 매월 업데이트할 원본 데이터
├── 📄 Pivo_재고예측시스템_v6.xlsx    ← 다음 버전 (업데이트 시 교체)
└── 📁 archive/
    ├── 📄 2026-04_snapshot.xlsx
    └── 📄 2026-03_snapshot.xlsx
```

**현재 Drive에 있는 파일**:
- 파일 ID: `1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q`
- URL: `https://docs.google.com/spreadsheets/d/1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q`

### 2-2. 필요한 계정/도구

| 도구 | 용도 | 보유 여부 |
|------|------|----------|
| Google Drive | 데이터 파일 보관 | ✅ 있음 |
| GitHub 계정 | 코드 및 Pages 배포 | ✅ `johnan-oss` |
| GitHub PAT | API 인증 토큰 | ✅ 있음 (만료 전 재발급 필요) |
| Python 3.8+ | 로컬 스크립트 실행 | 확인 필요 |

### 2-3. Python 패키지 설치

```bash
pip install openpyxl requests google-auth google-auth-httplib2 google-api-python-client
```

---

## 3. 방법 A — 수동 업데이트

### 흐름도

```
① Excel 업데이트 → ② Drive에 업로드 → ③ Python 스크립트 실행 → ④ git push → ⑤ 자동 배포
```

### Step 1: 데이터 파일 업데이트

매월 데이터 담당자가 Excel 파일에 신규 월 데이터를 입력합니다.

**열 매핑 규칙** (반드시 준수):
```
idx0  = 합계
idx1  = 2024-01
idx2  = 2024-02
...
idx13 = 2025-01
...
idx25 = 2026-01
idx26 = 2026-02
idx27 = 2026-03
idx28 = 2026-04   ← 매월 새 열 추가
idx29 = 2026-05   ← 5월 업데이트 시 추가
```

> ⚠️ **중요**: 새 월 데이터는 항상 기존 열 순서 뒤에 추가. 중간 삽입 금지.

### Step 2: Google Drive에 업로드

1. [Google Drive](https://drive.google.com) 접속
2. `Pivo_Dashboard_Data/` 폴더로 이동
3. 이전 파일을 `archive/` 폴더로 이동 (백업)
4. 새 파일 드래그 앤 드롭 업로드
5. 파일명을 `Pivo_재고예측시스템_최신.xlsx`로 통일 (스크립트가 이 이름 참조)

### Step 3: 로컬에서 생성 스크립트 실행

아래 스크립트를 저장하고 실행합니다:

**파일**: `~/Documents/Claude/update_dashboard.py`

```python
#!/usr/bin/env python3
"""
Pivo 대시보드 업데이트 스크립트
사용법: python update_dashboard.py --month 202605
"""

import argparse
import json
import re
import subprocess
from pathlib import Path
import openpyxl

# ─── 설정 ────────────────────────────────────────────
LOCAL_EXCEL = Path("~/Downloads/Pivo_재고예측시스템_최신.xlsx").expanduser()
DASHBOARD_HTML = Path("~/Documents/Claude/Pivo_재고_대시보드_v3.html").expanduser()
PAGES_DIR = Path("/tmp/pivo-pages")  # git repo 경로

# 열 인덱스 매핑 (0-based, idx0=합계 제외하고 계산)
# idx1=202401, idx13=202501, idx25=202601 ...
BASE_IDX = 1   # 202401이 col index 1
BASE_YEAR_MONTH = 202401

def col_idx_to_ym(idx):
    """열 인덱스를 YYYYMM 문자열로 변환"""
    # idx1 = 202401
    total_months = idx - 1  # 0-based offset from 202401
    year = 2024 + total_months // 12
    month = 1 + total_months % 12
    return f"{year}{month:02d}"

def ym_to_col_idx(yyyymm):
    """YYYYMM을 열 인덱스로 변환"""
    y, m = int(yyyymm[:4]), int(yyyymm[4:])
    return (y - 2024) * 12 + m  # idx1 = 202401

def parse_excel(path: Path):
    """Excel에서 SA/BD 데이터 파싱"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    
    SA = {}  # SA[channel][sku][region][yyyymm] = qty
    BD = {}  # BD[channel][sku][region][yyyymm][bundle_sku] = qty
    BDM = {} # BDM[bundle_sku] = [component_skus]
    
    rows = list(ws.iter_rows(values_only=True))
    
    for row in rows:
        if not row[0]:
            continue
        
        label = str(row[0]).strip()
        parts = label.split(',')
        
        if len(parts) < 3:
            continue
            
        row_type = parts[0].strip()
        channel = parts[1].strip() if len(parts) > 1 else ''
        sku = parts[2].strip() if len(parts) > 2 else ''
        region = parts[3].strip() if len(parts) > 3 else 'US'
        
        # 합계 행 스킵 (검증용으로만 사용)
        if row_type == '합계':
            continue
        
        # 데이터 파싱 (idx1부터 끝까지)
        monthly_data = {}
        for col_i in range(1, len(row)):
            val = row[col_i]
            if val is None or val == '':
                continue
            try:
                qty = float(val)
                if qty == 0:
                    continue
                ym = col_idx_to_ym(col_i)
                monthly_data[ym] = qty
            except (TypeError, ValueError):
                continue
        
        if not monthly_data:
            continue
            
        if row_type == 'SA직접판매':
            SA.setdefault(channel, {}).setdefault(sku, {}).setdefault(region, {}).update(monthly_data)
            
        elif row_type == 'Bundle기여':
            bundle_sku = parts[4].strip() if len(parts) > 4 else ''
            if not bundle_sku:
                continue
            BD.setdefault(channel, {}).setdefault(sku, {}).setdefault(region, {}).setdefault(bundle_sku, {}).update(monthly_data)
            
            # BDM 업데이트
            if bundle_sku not in BDM:
                BDM[bundle_sku] = []
            if sku not in BDM[bundle_sku]:
                BDM[bundle_sku].append(sku)
    
    return SA, BD, BDM

def update_html(SA, BD, BDM, html_path: Path, output_path: Path):
    """HTML 파일의 SA/BD/BDM 데이터를 새 데이터로 교체"""
    content = html_path.read_text(encoding='utf-8')
    
    # SA 교체
    sa_json = json.dumps(SA, ensure_ascii=False, separators=(',', ':'))
    content = re.sub(r'let SA\s*=\s*\{.*?\};', f'let SA = {sa_json};', content, flags=re.DOTALL)
    
    # BD 교체
    bd_json = json.dumps(BD, ensure_ascii=False, separators=(',', ':'))
    content = re.sub(r'const BD\s*=\s*\{.*?\};', f'const BD = {bd_json};', content, flags=re.DOTALL)
    
    # BDM 교체
    bdm_json = json.dumps(BDM, ensure_ascii=False, separators=(',', ':'))
    content = re.sub(r'const BDM\s*=\s*\{.*?\};', f'const BDM = {bdm_json};', content, flags=re.DOTALL)
    
    output_path.write_text(content, encoding='utf-8')
    print(f"✅ HTML 업데이트 완료: {output_path}")

def git_push(pages_dir: Path, month: str, html_path: Path):
    """변경사항 git commit & push"""
    import shutil
    
    # index.html 복사
    shutil.copy(html_path, pages_dir / "index.html")
    
    cmds = [
        ["git", "-C", str(pages_dir), "add", "index.html"],
        ["git", "-C", str(pages_dir), "commit", "-m", f"Data update: {month}"],
        ["git", "-C", str(pages_dir), "push", "origin", "main"],
    ]
    
    for cmd in cmds:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"❌ 오류: {' '.join(cmd)}\n{result.stderr}")
            return False
        print(f"✅ {' '.join(cmd[-2:])}")
    
    return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--month', required=True, help='업데이트 기준월 (예: 202605)')
    parser.add_argument('--excel', help='Excel 파일 경로 (기본값: ~/Downloads/Pivo_재고예측시스템_최신.xlsx)')
    parser.add_argument('--dry-run', action='store_true', help='git push 없이 HTML만 생성')
    args = parser.parse_args()
    
    excel_path = Path(args.excel) if args.excel else LOCAL_EXCEL
    
    print(f"📊 Excel 파싱 중: {excel_path}")
    SA, BD, BDM = parse_excel(excel_path)
    print(f"   SA: {sum(len(v) for v in SA.values())} SKUs, BD: {sum(len(v) for v in BD.values())} SKUs")
    
    output_html = DASHBOARD_HTML.parent / "Pivo_재고_대시보드_v3.html"
    update_html(SA, BD, BDM, DASHBOARD_HTML, output_html)
    
    if not args.dry_run:
        print(f"\n🚀 GitHub Pages 배포 중...")
        success = git_push(PAGES_DIR, args.month, output_html)
        if success:
            print(f"\n✅ 완료! 대시보드 업데이트됨")
            print(f"   URL: https://johnan-oss.github.io/pivo-dashboard/")
        else:
            print(f"\n⚠️  git push 실패. 수동으로 push 해주세요.")
    else:
        print(f"\n✅ Dry-run 완료. HTML 파일 확인 후 git push 하세요.")

if __name__ == '__main__':
    main()
```

### Step 4: 스크립트 실행

```bash
# Drive에서 파일 다운로드 후
python ~/Documents/Claude/update_dashboard.py --month 202605

# 또는 dry-run (HTML만 생성, push 안함)
python ~/Documents/Claude/update_dashboard.py --month 202605 --dry-run
```

### Step 5: 결과 확인

- 배포 URL: https://johnan-oss.github.io/pivo-dashboard/
- GitHub Actions: https://github.com/johnan-oss/pivo-dashboard/actions

---

## 4. 방법 B — 반자동 (GitHub Actions 월간 스케줄)

> 매월 특정 날짜에 자동으로 데이터를 읽어 대시보드를 업데이트합니다.  
> **전제조건**: Google Drive API 인증 설정 필요 (1회)

### 전체 흐름

```
[매월 5일 오전 9시]
GitHub Actions 트리거
    ↓
Google Drive API로 최신 Excel 다운로드
    ↓
Python 스크립트로 데이터 파싱 + HTML 생성
    ↓
index.html 자동 commit & push
    ↓
GitHub Pages 자동 배포 완료
    ↓
[선택] 이메일/Slack 알림
```

### Step 1: Google Drive API 설정 (1회만)

#### 1-1. Google Cloud Console에서 서비스 계정 생성

1. [Google Cloud Console](https://console.cloud.google.com) 접속
2. 새 프로젝트 생성: `pivo-dashboard`
3. **APIs & Services** → **Enable APIs** → `Google Drive API` 활성화
4. **Credentials** → **Create Credentials** → **Service Account**
   - 이름: `pivo-dashboard-bot`
   - 역할: `Viewer` (읽기 전용)
5. 생성된 서비스 계정 클릭 → **Keys** → **Add Key** → **JSON**
6. JSON 키 파일 다운로드 (`pivo-dashboard-bot-key.json`)

#### 1-2. Drive 파일에 서비스 계정 공유

1. Google Drive에서 `Pivo_재고예측시스템_최신.xlsx` 우클릭 → **공유**
2. 서비스 계정 이메일 추가 (예: `pivo-dashboard-bot@pivo-dashboard.iam.gserviceaccount.com`)
3. 권한: **뷰어**로 설정

#### 1-3. GitHub Secrets에 키 저장

```bash
# JSON 키 파일을 base64로 인코딩
base64 -i pivo-dashboard-bot-key.json | tr -d '\n'
```

복사한 값을 GitHub 레포 Secrets에 저장:
1. [GitHub 레포](https://github.com/johnan-oss/pivo-dashboard) → **Settings** → **Secrets and variables** → **Actions**
2. **New repository secret**:
   - Name: `GDRIVE_SERVICE_ACCOUNT_KEY`
   - Value: (위에서 base64 인코딩한 값)

또한 Excel 파일 ID도 저장:
- Name: `GDRIVE_FILE_ID`  
- Value: `1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q`

### Step 2: GitHub Actions Workflow 파일 생성

GitHub 레포에 다음 파일을 생성합니다:  
**경로**: `.github/workflows/monthly-update.yml`

```yaml
name: Monthly Dashboard Update

on:
  # 매월 5일 오전 9시 (KST = UTC+9, 즉 UTC 0시)
  schedule:
    - cron: '0 0 5 * *'
  
  # 수동 실행도 가능
  workflow_dispatch:
    inputs:
      month:
        description: '업데이트 기준월 (예: 202605)'
        required: false
        default: ''

jobs:
  update-dashboard:
    runs-on: ubuntu-latest
    
    steps:
      - name: Checkout repository
        uses: actions/checkout@v4
      
      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.11'
      
      - name: Install dependencies
        run: |
          pip install openpyxl google-auth google-auth-httplib2 google-api-python-client
      
      - name: Set current month
        id: month
        run: |
          if [ -n "${{ github.event.inputs.month }}" ]; then
            echo "month=${{ github.event.inputs.month }}" >> $GITHUB_OUTPUT
          else
            echo "month=$(date +%Y%m)" >> $GITHUB_OUTPUT
          fi
      
      - name: Download Excel from Google Drive
        env:
          GDRIVE_SERVICE_ACCOUNT_KEY: ${{ secrets.GDRIVE_SERVICE_ACCOUNT_KEY }}
          GDRIVE_FILE_ID: ${{ secrets.GDRIVE_FILE_ID }}
        run: |
          echo "$GDRIVE_SERVICE_ACCOUNT_KEY" | base64 -d > /tmp/service_account.json
          python << 'EOF'
          import os
          import json
          from google.oauth2 import service_account
          from googleapiclient.discovery import build
          from googleapiclient.http import MediaIoBaseDownload
          import io
          
          creds = service_account.Credentials.from_service_account_file(
              '/tmp/service_account.json',
              scopes=['https://www.googleapis.com/auth/drive.readonly']
          )
          service = build('drive', 'v3', credentials=creds)
          
          file_id = os.environ['GDRIVE_FILE_ID']
          request = service.files().get_media(fileId=file_id)
          
          fh = io.BytesIO()
          downloader = MediaIoBaseDownload(fh, request)
          done = False
          while not done:
              _, done = downloader.next_chunk()
          
          with open('/tmp/pivo_data.xlsx', 'wb') as f:
              f.write(fh.getvalue())
          
          print("✅ Drive에서 Excel 다운로드 완료")
          EOF
      
      - name: Generate dashboard HTML
        run: |
          python << 'PYEOF'
          import json, re, openpyxl
          from pathlib import Path
          
          def col_idx_to_ym(idx):
              total_months = idx - 1
              year = 2024 + total_months // 12
              month = 1 + total_months % 12
              return f"{year}{month:02d}"
          
          def parse_excel(path):
              wb = openpyxl.load_workbook(path, data_only=True)
              ws = wb.active
              SA, BD, BDM = {}, {}, {}
              
              for row in ws.iter_rows(values_only=True):
                  if not row[0]:
                      continue
                  label = str(row[0]).strip()
                  parts = [p.strip() for p in label.split(',')]
                  if len(parts) < 3:
                      continue
                  
                  row_type, channel, sku = parts[0], parts[1], parts[2]
                  region = parts[3] if len(parts) > 3 else 'US'
                  
                  if row_type == '합계':
                      continue
                  
                  monthly_data = {}
                  for col_i in range(1, len(row)):
                      val = row[col_i]
                      if val is None or val == '' or val == 0:
                          continue
                      try:
                          qty = float(val)
                          if qty != 0:
                              monthly_data[col_idx_to_ym(col_i)] = qty
                      except (TypeError, ValueError):
                          continue
                  
                  if not monthly_data:
                      continue
                  
                  if row_type == 'SA직접판매':
                      SA.setdefault(channel, {}).setdefault(sku, {}).setdefault(region, {}).update(monthly_data)
                  elif row_type == 'Bundle기여':
                      bundle_sku = parts[4] if len(parts) > 4 else ''
                      if bundle_sku:
                          BD.setdefault(channel, {}).setdefault(sku, {}).setdefault(region, {}).setdefault(bundle_sku, {}).update(monthly_data)
                          BDM.setdefault(bundle_sku, [])
                          if sku not in BDM[bundle_sku]:
                              BDM[bundle_sku].append(sku)
              
              return SA, BD, BDM
          
          print("📊 Excel 파싱 중...")
          SA, BD, BDM = parse_excel('/tmp/pivo_data.xlsx')
          
          html_path = Path('index.html')
          content = html_path.read_text(encoding='utf-8')
          
          sa_json = json.dumps(SA, ensure_ascii=False, separators=(',', ':'))
          bd_json = json.dumps(BD, ensure_ascii=False, separators=(',', ':'))
          bdm_json = json.dumps(BDM, ensure_ascii=False, separators=(',', ':'))
          
          content = re.sub(r'let SA\s*=\s*\{.*?\};', f'let SA = {sa_json};', content, flags=re.DOTALL)
          content = re.sub(r'const BD\s*=\s*\{.*?\};', f'const BD = {bd_json};', content, flags=re.DOTALL)
          content = re.sub(r'const BDM\s*=\s*\{.*?\};', f'const BDM = {bdm_json};', content, flags=re.DOTALL)
          
          html_path.write_text(content, encoding='utf-8')
          print("✅ HTML 업데이트 완료")
          PYEOF
      
      - name: Commit and push
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add index.html
          
          if git diff --cached --quiet; then
            echo "변경사항 없음. 스킵."
          else
            git commit -m "Auto update: $(date +'%Y-%m') data refresh"
            git push
            echo "✅ 배포 완료"
          fi
      
      - name: Notify on success
        if: success()
        run: |
          echo "🎉 대시보드 업데이트 완료!"
          echo "URL: https://johnan-oss.github.io/pivo-dashboard/"
```

### Step 3: Workflow 파일 업로드

```bash
# 로컬에서 실행
mkdir -p /tmp/pivo-pages/.github/workflows
# 위 YAML 내용을 파일로 저장
cat > /tmp/pivo-pages/.github/workflows/monthly-update.yml << 'EOF'
(위 YAML 내용 붙여넣기)
EOF

cd /tmp/pivo-pages
git add .github/
git commit -m "Add monthly auto-update workflow"
git push origin main
```

### Step 4: 수동 테스트 실행

1. GitHub 레포 → **Actions** 탭
2. **Monthly Dashboard Update** → **Run workflow**
3. 실행 결과 확인

---

## 5. 방법 C — 완전 자동 (Google Sheets 직접 연동)

> Google Sheets에서 직접 데이터 편집 시 대시보드가 자동 업데이트됩니다.

### 구성도

```
Google Sheets (데이터 편집)
    ↓ Apps Script (변경 감지)
    ↓ GitHub API (push trigger)
    ↓ GitHub Actions
    ↓ GitHub Pages 배포
```

### Google Sheets Apps Script 설정

Sheets에서 **확장 프로그램** → **Apps Script** → 새 스크립트:

```javascript
// Google Sheets Apps Script
// 시트 변경 시 GitHub Actions 트리거

const GITHUB_TOKEN = '여기에_PAT_토큰';  // GitHub Settings에서 생성
const REPO_OWNER = 'johnan-oss';
const REPO_NAME = 'pivo-dashboard';

function onEdit(e) {
  // 편집 발생 시 5분 후 트리거 (즉시 아닌 이유: 여러 셀 편집 배치 처리)
  ScriptApp.newTrigger('triggerGitHubAction')
    .timeBased()
    .after(5 * 60 * 1000)  // 5분
    .create();
}

function triggerGitHubAction() {
  const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/actions/workflows/monthly-update.yml/dispatches`;
  
  const payload = {
    ref: 'main',
    inputs: {
      month: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMM')
    }
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  UrlFetchApp.fetch(url, options);
  console.log('GitHub Actions 트리거 완료');
  
  // 중복 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'triggerGitHubAction') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
```

> ⚠️ **주의**: 방법 C는 Google Sheets 데이터 형식이 현재 Excel 형식과 동일해야 합니다.  
> 현재 Excel 형식 유지 권장. Google Sheets 전환 시 [6장](#6-데이터-형식-가이드) 참조.

---

## 6. 데이터 형식 가이드

### 현재 Excel 형식 (권장 유지)

```
행 레이블 형식: {타입},{채널},{SKU},{지역}[,{번들SKU}]

예시:
SA직접판매,SH,SM,US         → SM의 SH 직접판매 (US)
SA직접판매,AMZ,NPVB,US      → NPVB의 AMZ 직접판매
Bundle기여,SH,SM,US,STANDPRE9  → STANDPRE9 번들의 SM 기여분
합계,SH,SM,US               → 검증용 합계 (스크립트가 스킵)
```

### Excel vs Google Sheets 비교

| 항목 | Excel (.xlsx) | Google Sheets |
|------|--------------|---------------|
| 데이터 입력 | 로컬 앱 | 웹 브라우저 |
| 공동 편집 | ❌ (한 명씩) | ✅ (실시간) |
| 자동화 | Python openpyxl | Apps Script |
| API 접근 | Drive API | Sheets API |
| 현재 지원 | ✅ 스크립트 있음 | ⚠️ 추가 설정 필요 |

### Google Sheets 전환 시 주의사항

현재 Excel → Google Sheets로 전환하려면:

1. Excel 파일을 Drive에 올린 후 "Google Sheets로 열기" 클릭
2. 수식이 있는 경우 값으로 붙여넣기 확인
3. Python 파서를 `gspread` 라이브러리로 교체:

```python
# Google Sheets 파서 (openpyxl 대체)
import gspread
from google.oauth2 import service_account

def parse_google_sheets(sheet_id, creds_path):
    creds = service_account.Credentials.from_service_account_file(
        creds_path,
        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)
    ws = sh.get_worksheet(0)
    
    rows = ws.get_all_values()
    # 이후 로직은 Excel 파서와 동일
    ...
```

---

## 7. 월간 업데이트 체크리스트

### 매월 데이터 업데이트 시 (방법 A 기준)

```
□ 1. Excel 파일 열기
□ 2. 새 월 데이터 열 추가 (idx 순서 확인)
□ 3. 모든 SKU 행에 데이터 입력
□ 4. 합계 열 수식 확인 (합계 = sum of all months)
□ 5. STANDPRE9/10 번들 데이터 검증 (과거 오류 재발 방지)
□ 6. 파일 저장 후 Google Drive 업로드
□ 7. update_dashboard.py --dry-run 으로 확인
□ 8. 검증 통과 후 실제 실행 (push)
□ 9. 배포 URL 접속하여 최종 확인
□ 10. 검증 매뉴얼에 업데이트 이력 기록
```

### 검증 기준 (Pivo_대시보드_검증_매뉴얼.md 참조)

| SKU | SH 허용 오차 | AMZ 허용 오차 |
|-----|------------|--------------|
| SM | ±10 units | ±5 units |
| TRPD | ±10 units | ±5 units |
| NPVB | ±10 units | ±5 units |
| 기타 | ±5 units | ±3 units |

---

## 8. 문제 해결

### Q1: Google Drive API 인증 오류

```
Error: invalid_grant: Invalid JWT Signature
```

**해결**: 서비스 계정 JSON 키 재발급 → GitHub Secrets 업데이트

---

### Q2: Excel 파일 파싱 오류

```
openpyxl.utils.exceptions.InvalidFileException
```

**해결**: 파일이 .xlsx 형식인지 확인. .xls 구버전 파일은 변환 필요:
```bash
# LibreOffice로 변환
libreoffice --headless --convert-to xlsx 파일.xls
```

---

### Q3: GitHub Actions 실패

1. Actions 탭 → 실패한 실행 클릭 → 로그 확인
2. 주요 원인:
   - Secrets 미설정 → Settings → Secrets 확인
   - Drive 파일 공유 안 됨 → 서비스 계정에 공유 재확인
   - Python 의존성 오류 → requirements 확인

---

### Q4: 대시보드 데이터가 업데이트 안 됨

1. 브라우저 캐시 삭제 후 새로고침 (Ctrl+Shift+R)
2. GitHub Pages 배포 상태 확인: https://github.com/johnan-oss/pivo-dashboard/deployments
3. `index.html` 파일의 마지막 수정일 확인

---

### Q5: 특정 SKU 값이 이상함

검증 매뉴얼(`Pivo_대시보드_검증_매뉴얼.md`) 참조하여:
1. Excel 원본 해당 행 확인
2. 열 인덱스 매핑 재확인
3. STANDPRE9/10 등 번들 데이터 별도 검증

---

## 부록: 관련 파일 위치

| 파일 | 경로 |
|------|------|
| 대시보드 HTML | `/Users/jian/Documents/Claude/Pivo_재고_대시보드_v3.html` |
| 검증 매뉴얼 | `/Users/jian/Documents/Claude/Pivo_대시보드_검증_매뉴얼.md` |
| 팀 컨텍스트 | `/Users/jian/Documents/Claude/Pivo_대시보드_팀협업_컨텍스트.md` |
| 이 파일 | `/Users/jian/Documents/Claude/Pivo_대시보드_자동업데이트_매뉴얼.md` |
| GitHub 레포 | `https://github.com/johnan-oss/pivo-dashboard` |
| 배포 URL | `https://johnan-oss.github.io/pivo-dashboard/` |
| Google Drive 원본 | `https://drive.google.com/file/d/1zV0gJuS9Rmh1RhejNvOGJ6djL2eE3Z0Q` |

---

*이 매뉴얼은 Claude (Anthropic)와 함께 작성되었습니다. 문의: john.an@3i.ai*
