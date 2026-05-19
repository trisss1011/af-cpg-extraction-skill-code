# cpg-data-extraction / scripts/

> v3.2 전략 C (template + helper scripts) 구성 요소.
> 이 폴더는 `cpg-data-extraction` 스킬 저장 파이프라인을 자동화하는 Python 스크립트 세트입니다.

---

## 개요

| 파일 | 역할 | 호출 주체 |
|---|---|---|
| `save_extract.py` | 1편 논문 JSON → 단독 엑셀(`AF_extract_<번호>_<id>.xlsx`) 저장 | Claude (추출 세션) |
| `migrate_format.py` | v2.x 이전 추출 파일에 v3.0+ 서식 일괄 적용 | 사용자/조정자 (필요 시) |
| `migrate_to_v3.2.py` (v3.2 신규) | 47열(v3.0/v3.1) → 48열(v3.2) 헤더 마이그레이션 (AU=analysis_set 삽입, notes→AV 시프트) | 사용자 (v3.2 업그레이드 시 1회) |
| `version_info.py` | 엑셀 메타데이터에 스킬 버전 기록·조회 | 위 스크립트들 내부 호출 + CLI |

**핵심 원칙**: 이 스크립트들은 **데이터 값을 생성하지 않는다**. Claude가 만든 JSON을 받아 엑셀 서식으로 변환할 뿐이다. 추출 정확성은 Claude + 사용자 검토가 책임진다.

---

## 환경 요구

- **Python**: 3.8 이상 (테스트: 3.12.10)
- **openpyxl**: 3.0 이상 (테스트: 3.1.5)

### 설치 확인

```
python --version            # 3.8 이상 확인
python -m pip show openpyxl # 설치 확인
```

### openpyxl 미설치 시

```
python -m pip install openpyxl
```

(Windows에서 `python` 명령이 Microsoft Store 스텁으로 잡히면 Python 본체 설치 필요. 본 프로젝트 첫설정 가이드 참조.)

---

## 1. save_extract.py — 저장 주력

### 목적

Claude가 추출한 1편 논문 데이터(JSON)를 규격화된 엑셀 파일로 저장.

### 사용법

```
python save_extract.py <input_json_path>
```

예:
```
python 00_skills/cpg-data-extraction/scripts/save_extract.py 99_Scratch/_tmp_2_Zhang_2018.json
```

### 입력 JSON 형식

```json
{
  "기본정보": [
    {
      "번호": 2,
      "study_id": "Zhang_2018",
      "author": "Zhang N et al.",
      "year": 2018,
      "title": "...",
      "journal": "...",
      "exclude": null,
      "...(나머지 42개 선택 필드)": "..."
    }
  ],
  "아웃컴": [
    {
      "번호": 2,
      "study_id": "Zhang_2018",
      "outcome_std": "SR conversion",
      "outcome_original": "conversion rate to sinus rhythm within 24h",
      "data_type": "이분형",
      "E_n": 20,
      "E_val1": 15,
      "E_val2": 20,
      "C_n": 21,
      "C_val1": 17,
      "C_val2": 21,
      "p_value": 0.72,
      "...(기타)": "..."
    }
  ],
  "한의중재_한약": [
    {
      "번호": 2,
      "study_id": "Zhang_2018",
      "formula_name_kr": "온심과립",
      "formula_name_cn": "稳心颗粒",
      "formulation": "과립",
      "composition": "...",
      "administration": "경구, 18g×3회/일",
      "note": "..."
    }
  ]
}
```

**핵심 규칙**:
- 최상위 키 3개(`기본정보`, `아웃컴`, `한의중재_한약`) 모두 리스트여야 함
- `기본정보`, `한의중재_한약`: 정확히 1행
- `아웃컴`: 0~N행 (보통 여러 개)
- 3시트의 `번호`·`study_id`는 **반드시 동일**
- 아웃컴 열 중 `E_val1` 등은 JSON에서 `\n` 없이 단순 키 사용 (엑셀에선 `E_val1\n(Mean/Event)`으로 표시됨)

### 필수 필드

- 기본정보: `번호`(int), `study_id`(str), `author`(str), `year`(int)
- 아웃컴 각 행: `번호`(int), `study_id`(str)
- 한의중재_한약: `번호`(int), `study_id`(str)
- 나머지 누락 OK (빈 셀 처리) — NR 등은 명시 문자열로 기재

### 출력

- **성공**: `90_Output/extracts/AF_extract_<번호>_<study_id>.xlsx` 생성·저장
- **경고 로그**: `90_Output/extracts/_save_warnings.log` (append)
- **임시 JSON**: `99_Scratch/_tmp_*.json`은 성공 시 자동 삭제

### Exit codes

| 코드 | 의미 |
|---|---|
| 0 | 저장 성공 |
| 1 | JSON 읽기/파싱 오류 |
| 2 | 검증 실패 (필수 필드 / 타입 / 구조 / 3시트 번호 불일치 / 템플릿 손상) |
| 3 | 파일 쓰기 오류 (디스크·권한) |
| 4 | 템플릿 미존재 |

### 경고 종류 (저장은 진행, 로그만 기록)

| 코드 | 조건 |
|---|---|
| `OVERWRITE` | 동일 파일명 이미 존재 (덮어썼음) |
| `CELL_TRUNC` | 셀 값 32,767자 초과 (Excel 한도) |
| `CJK_SUSPICIOUS` | 한자·특수문자 렌더링 의심 (PUA 등) |
| `FILE_TOO_SMALL` | 생성된 파일 크기 3KB 미만 |

---

## 2. migrate_format.py — 서식 마이그레이션

### 목적

v2.x 시절 저장된 기존 추출 파일(`AF_extract_*.xlsx`)에 v3.2의 추가 서식(freeze panes, 숫자 서식 등)을 일괄 적용.

### 언제 쓰는가

- v2.x로 뽑은 기존 파일들을 v3.2 규격에 맞추고 싶을 때
- **데이터 값은 절대 건드리지 않음** — 서식만 변경 (유형 2 마이그레이션)

### 사용법

```
python migrate_format.py <file_or_dir> [--dry-run] [--force]
```

**옵션**:
- `--dry-run`: 실제 저장 없이 변경 예정 내용만 표시
- `--force`: 이미 v3.2 태그된 파일도 재적용

**예**:
```
# 폴더 전체 dry-run
python migrate_format.py 90_Output/extracts --dry-run

# 실제 마이그레이션
python migrate_format.py 90_Output/extracts

# 단일 파일
python migrate_format.py 90_Output/extracts/AF_extract_5_Xxx_2024.xlsx
```

### 적용 항목

- `freeze_panes = 'A2'` (3시트 모두)
- 숫자 서식 `'0'`:
  - 기본정보: A(번호), D(year), R(af_type_code), S(af_type_other), U(comorbidity_code)
  - 아웃컴: A(번호)
  - 한의중재_한약: A(번호)
- 기본정보 AJ~AT열 RoB 배경 `#D6EAF8` 재확인

### 안전장치

- **자동 백업**: 마이그레이션 전 `<파일명>.bak.xlsx` 생성
- **멱등성**: 이미 v3.2 태그된 파일은 skip
- **dry-run**: 실제 저장 전 확인 가능
- **데이터 셀 값 불변**: 서식만 조작, `.value`는 읽기만

### 실패 시 복구

```
# .bak.xlsx로 원복
cp AF_extract_5_Xxx_2024.bak.xlsx AF_extract_5_Xxx_2024.xlsx
```

---

## 2.5. migrate_to_v3.2.py — 헤더 마이그레이션 (v3.2 신규)

### 목적

v3.0/v3.1로 만든 47열 추출/마스터 파일을 v3.2의 48열 구조로 변환. AU(notes)를 AV로 시프트하고 AU 자리에 `analysis_set` 신규 열을 삽입한다.

### 언제 쓰는가

- v3.0/v3.1로 추출한 기존 파일을 v3.2 마스터에 머지하기 전에 1회
- 마스터 파일 자체를 v3.2 스키마로 업그레이드할 때

### 사용법

```
# 단일 파일
python migrate_to_v3.2.py "90_Output/AF_CPG_data_extraction_심상송.xlsx"

# 폴더 (재귀 아님)
python migrate_to_v3.2.py 90_Output/extracts/

# 와일드카드
python migrate_to_v3.2.py "90_Output/extracts/*.xlsx"
```

### 처리 결과

| 상태 | 의미 |
|---|---|
| `OK` | 47열 → 48열 변환 완료, `.v3.1.bak` 백업 생성 |
| `SKIP_ALREADY_V3.2` | 이미 48열이므로 변환 불필요 |
| `SKIP_UNEXPECTED_COL_COUNT` | 열 수가 47도 48도 아님 (수동 확인 필요) |
| `SKIP_UNEXPECTED_LAST_HEADER` | 47번째 열이 `notes`가 아님 (변환 거부) |
| `ERROR_LOCKED` | Excel에서 열려있어 변환 불가 |

### 안전장치

- **자동 백업**: 변환 전 `<원본>.v3.1.bak` 생성 (멱등 — 이미 백업 있으면 재생성 안 함)
- **멱등성**: 이미 48열이면 skip (재실행 안전)
- **헤더 정합 확인**: 47열인 경우 마지막 열명이 `notes`가 아니면 거부
- **셀 서식 보존**: 폰트·정렬·배경·테두리·number_format 모두 유지

### 권장 순서 (v3.0/v3.1 → v3.2 업그레이드)

```
1. python migrate_to_v3.2.py "90_Output/AF_CPG_data_extraction_<나>.xlsx"
2. python migrate_to_v3.2.py "90_Output/extracts/*.xlsx"
3. (선택) python migrate_format.py 90_Output/extracts/   # 서식 재정비
```

---

## 3. version_info.py — 버전 태깅

### 목적

엑셀 파일 메타데이터에 스킬 버전(현재 `v3.2`)을 기록·조회. save_extract.py·migrate_format.py가 내부적으로 호출.

### 기록 방식 (자동 선택)

1. **Custom Document Property** (openpyxl 3.1+ 지원): 키 `cpg_skill_version`
2. **Fallback: properties.keywords**: `cpg-skill=v3.2` 태그 삽입

### CLI 사용

```
# 버전 읽기
python version_info.py read <xlsx>

# 버전 설정 (기본 v3.2)
python version_info.py set <xlsx>

# 버전 수동 지정
python version_info.py set <xlsx> v3.1
```

### 라이브러리 사용 (다른 Python 스크립트에서)

```python
from version_info import set_version, get_version, SKILL_VERSION

wb = openpyxl.load_workbook(path)
set_version(wb)          # v3.2 기록
v = get_version(wb)      # 'v3.2' or None
wb.save(path)
```

---

## 공통: Windows 환경 주의사항

### 콘솔 인코딩

Windows cmd는 기본 cp949. 스크립트 출력에 한글·한자가 있으면 깨질 수 있음. 실행 시 다음 환경변수 설정 권장:

```cmd
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
```

PowerShell:
```powershell
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONUTF8 = '1'
```

또는 Python 자체가 3.7+이면 `PYTHONUTF8=1` 한 줄로 UTF-8 모드 강제.

### 파일 경로 공백·한글

스크립트는 pathlib를 사용해 공백·한글 경로 대응. 단 명령줄에서 경로 넘길 때는 큰따옴표로 감싸기:

```
python save_extract.py "99_Scratch/_tmp_2_Zhang_2018.json"
```

### OneDrive 동기화

추출 결과 폴더(`90_Output/extracts/`)가 OneDrive 안에 있으면 저장 시점에 동기화 훅 때문에 수 초 지연 가능. 권장: 작업 폴더는 OneDrive 밖 로컬(`C:\work\`) 사용, 완료 후 OneDrive에 수동 복사.

---

## 트러블슈팅

| 증상 | 원인 | 조치 |
|---|---|---|
| `[ERROR] 템플릿 미존재` | `sample_v2.6.xlsx` 없음 | 스킬 폴더 재배포 확인 |
| `[VALIDATION ERROR] 기본정보 필수 키 누락` | JSON에 `번호`/`study_id`/`author`/`year` 중 누락 | JSON 재생성 |
| `[VALIDATION ERROR] 시트 간 번호 불일치` | 3시트 중 일부 번호가 다름 | Claude 추출 단계 재검토 |
| `[ERROR] 프로젝트 루트(90_Output 포함) 미탐지` | 스크립트가 프로젝트 밖에서 실행됨 | 프로젝트 루트 내부에서 실행 |
| 한글 깨짐 | 콘솔 인코딩 | 위 `PYTHONIOENCODING` 설정 |
| `openpyxl` ImportError | 패키지 미설치 | `pip install openpyxl` |

---

## 파일 레이아웃

```
00_skills/cpg-data-extraction/
├── SKILL.md
├── CHANGELOG.md
├── sample_v2.6.xlsx           ← 템플릿 (save_extract가 복사함, v3.2: 48열)
├── sample_v2.4.xlsx           ← 구버전 호환용 (보관)
├── references/
│   ├── af-outcomes.md
│   ├── korean-medicine-intervention.md
│   └── rob-extraction-fields.md
└── scripts/                   ← 이 폴더
    ├── README.md              ← 이 문서
    ├── save_extract.py
    ├── migrate_format.py
    ├── migrate_to_v3.2.py     ← v3.2 신규 (47→48열 헤더 마이그레이션)
    └── version_info.py
```

---

## 변경 이력

- **v3.2 (2026-05-19)**: cowork v2.6 변경사항 통합 — analysis_set 신열(48열), study_design 분류 개정, AF/RVR 엄격화, HRV 제외, SAE 통합 등. 템플릿 `sample_v2.4.xlsx` → `sample_v2.6.xlsx`. `BASIC_HEADERS` 47→48열. 신규 스크립트 `migrate_to_v3.2.py` (47열→48열 헤더 마이그레이션).
- **v3.0 (2026-04-24)**: 최초 빌드. Phase 1 10개 결정 반영. save_extract·migrate_format·version_info 3종 세트.
