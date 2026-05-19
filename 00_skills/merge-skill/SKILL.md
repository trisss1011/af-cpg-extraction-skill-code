---
name: merge-skill
version: 2.6
updated: 2026-05-19
description: >
  심방세동 한의CPG 데이터 추출 작업에서 세션별로 분리 저장된 추출 파일들(`90_Output/extracts/AF_extract_<번호>_<study_id>.xlsx`)을 마스터 엑셀 `AF_CPG_data_extraction_심상송.xlsx`에 안전하게 병합한다.
  동시 세션 추출의 lost update(쓰기 손실) 경쟁 조건을 회피하기 위해 추출은 단독 파일로 분리하고, 병합은 단일 시점에 원자적으로 처리한다.
  트리거: 머지, 병합, 추출 머지, 추출 병합, 마스터에 합쳐줘, 추출 파일 합쳐줘, AF_extract 합쳐줘,
  세션별 파일 합치기, 추출 결과 통합, merge extracts, consolidate extracts.
---

# CPG 추출 결과 머지 스킬

## 목적

여러 세션에서 병렬 추출된 `90_Output/extracts/AF_extract_<번호>_<study_id>.xlsx` 파일들을 마스터 `90_Output/AF_CPG_data_extraction_심상송.xlsx`에 통합한다. openpyxl은 파일 전체를 메모리 스냅샷으로 로드한 뒤 통째로 덮어쓰는 방식이라 동시 세션 쓰기에서 lost update가 발생한다. 본 스킬은 추출(병렬, 단독 파일)과 병합(단일 시점, 원자적)을 분리하여 이 문제를 근본적으로 차단한다.

## 작업 흐름

```
폴더 스캔 → 마스터/락 확인 → 스키마 검증 → 번호 키 검증 → 중복 검증 →
미리보기(삽입 위치 포함) → 사용자 승인 → 백업 →
row 2 스타일 템플릿 추출 → 번호 오름차순으로 insert_rows + 값/스타일 기록 →
행 수 검증 → 단일 save →
사후 검증(서식 spot check 포함) → 추출 파일 이관 → 로그 기록 → 결과 보고
```

## 파일 구조

```
90_Output/
├── AF_CPG_data_extraction_심상송.xlsx       # 마스터 (고정 파일명)
├── extracts/                               # 세션별 추출 파일
│   ├── AF_extract_30_Li_2025.xlsx
│   ├── AF_extract_31_Wang_2024.xlsx
│   └── merged/
│       └── 20260420_143022/                # 머지 시점별 아카이브
│           ├── AF_extract_30_Li_2025.xlsx
│           └── merge_log.txt
└── backups/
    ├── master_backup_20260420_143022.xlsx
    └── ...                                 # 최근 10개만 유지
```

## 핵심 원칙

1. **원자성**: 모든 시트의 변경분을 메모리에서 검증한 뒤 단일 `save()`로 처리. 시트별 분할 저장 금지.
2. **검증 우선**: pre-flight 검증을 통과하지 못한 파일은 머지에서 제외하고 사유 보고. 강제 진행 옵션 없음.
3. **무손실 보존**: 추출 파일은 머지 후 삭제하지 않고 `merged/<timestamp>/`로 이동.
4. **중복 차단**: 마스터에 이미 존재하는 study_id는 기본 거부. 사용자 명시적 재머지 지시 시만 overwrite.
5. **사용자 승인**: 머지 실행 전 미리보기 → 명시적 승인. 자동 진행 금지.
6. **Excel 잠금 차단**: 마스터에 락 파일이 있으면 즉시 중단. 진행 시 lost update 재발.
7. **번호 기반 정렬 삽입** (v1.1): 신규 행은 A열 `번호`(정수) 오름차순 위치에 삽입. 기존 행의 순서는 건드리지 않음. 같은 번호 내부(아웃컴)는 추출 당시 순서(stable) 유지.
8. **서식 보존** (v2.5.1): 기존 행은 `insert_rows`의 자동 시프트로 셀 서식이 그대로 보존된다. 신규 행은 row 2의 셀 스타일을 col별 템플릿으로 복사하여 시각적 일관성을 확보한다. 데이터 영역에 대한 `delete_rows` + 재기입은 셀 서식을 잃으므로 금지.

---

## 1단계: 마스터 파일 식별

마스터 파일명은 **고정**이다: `90_Output/AF_CPG_data_extraction_심상송.xlsx`

- 파일 존재: 다음 단계
- 파일 없음: 중단, "마스터 파일(`AF_CPG_data_extraction_심상송.xlsx`)을 `90_Output/`에서 찾을 수 없습니다. 파일명·위치를 확인해주세요"
- 사용자가 다른 작업자 파일로 머지를 명시적으로 요청한 경우에만 다른 파일명 허용 (예: "`AF_CPG_data_extraction_김철수.xlsx`로 머지해줘")

## 2단계: 잠금 파일 확인

마스터 파일과 동일 폴더에 `~$<마스터파일명>` 형태의 Excel 락 파일이 있는지 확인.

- 존재: **즉시 중단**, "마스터 파일이 Excel에서 열려 있습니다. 저장하지 말고 닫은 뒤 다시 실행하세요."
- 없음: 다음 단계

## 3단계: 추출 파일 스캔

`90_Output/extracts/` 디렉터리에서 패턴 `AF_extract_*_*.xlsx` 파일 수집.

- `merged/` 하위는 제외
- 임시 락 파일(`~$*.xlsx`) 제외
- 0개: "머지할 추출 파일이 없습니다" 보고 후 정상 종료

## 4단계: 스키마 검증

각 추출 파일에 대해 마스터와 비교한다.

### 4A. 시트 구성

마스터의 시트 집합과 추출의 시트 집합이 정확히 일치해야 통과. 다르면 거부.

### 4B. 헤더 일치

시트별로 1행(헤더)을 비교한다.

| 케이스 | 처리 |
|--------|------|
| 추출 헤더 = 마스터 헤더 | 통과 |
| 추출 헤더가 마스터 헤더의 prefix (마스터에 끝부분 컬럼 추가, 순서 동일) | 통과 + 신규 컬럼은 빈 값으로 패딩, 사용자에게 경고 |
| 컬럼 순서 다름 | 거부 |
| 컬럼명 다름 (오타·표기 불일치 포함) | 거부 |
| 추출에만 있는 컬럼 | 거부 (데이터 손실 위험) |
| 컬럼 수가 동일하나 일부 이름 다름 | 거부 |

거부 사유는 어느 시트의 어느 컬럼이 문제인지 구체적으로 보고.

### 4C. 결과 분류

- **통과**: 다음 단계로
- **거부**: 머지 대상 제외, 원위치 유지, 사유 기록

### 4D. 정렬 키(번호) 검증 (v1.1)

각 추출 파일의 `기본정보` 시트 A열(`번호`)이 정수이고 비어있지 않은지 확인한다.

- 통과: 해당 번호를 해당 추출의 정렬 키로 사용
- 누락/비정수: 해당 파일 거부, 사유 보고

같은 추출 내 세 시트(`기본정보`, `아웃컴`, `한의중재_한약`)의 모든 데이터행 A열 번호는 동일해야 한다 (해당 논문 고유 번호). 불일치 시 거부.

## 5단계: 중복 study_id 검증

각 추출 파일의 `기본정보` 시트에서 study_id를 추출한다 (1차 키).

- 마스터의 `기본정보`에 동일 study_id가 존재하면 충돌
- 충돌 시 기본 거부

### 명시적 재머지 (overwrite)

사용자가 다음과 같이 명시적으로 지시한 경우에만 overwrite 진행:

- "30번 재머지"
- "Li_2025 덮어쓰기"
- "재머지: Li_2025"
- "overwrite 30"

명시되지 않은 study_id는 절대 자동 overwrite 금지.

overwrite 처리 방식:

1. 마스터의 모든 시트에서 해당 study_id가 포함된 행을 모두 삭제 (역순 삭제)
2. 추출 파일의 행을 신규로 append
3. 즉, "기존 완전 삭제 후 신규 삽입". 부분 업데이트 금지

## 6단계: 미리보기 출력

다음 형식으로 사용자에게 표시한다.

```
머지 대상: 3개 파일 중 2개 통과 (번호 오름차순 처리)
─────────────────────────────────────────────────
✓ AF_extract_30_Li_2025.xlsx (번호=30)
   기본정보 +1행   → 삽입 예정: 기본정보 row8
   아웃컴 +29행    → 삽입 예정: 아웃컴 row73
   한의중재_한약 +1행 → 삽입 예정: 한의중재_한약 row8
✓ AF_extract_31_Wang_2024.xlsx (번호=31)
   기본정보 +1행   → 삽입 예정: 기본정보 row9
   아웃컴 +18행    → 삽입 예정: 아웃컴 row102
   한의중재_한약 +1행 → 삽입 예정: 한의중재_한약 row9
✗ AF_extract_32_Chen_2025.xlsx — 거부
   사유: 아웃컴 시트 헤더 불일치 ('effect_se' 컬럼이 추출에만 존재)

마스터 행 수 변화 (예상)
─────────────────────────────────────────────────
기본정보:        4 → 6
아웃컴:         49 → 96
한의중재_한약:    4 → 6

정렬 규칙: A열 '번호' 오름차순. 기존 행 순서는 유지.

백업 위치:
  90_Output/backups/master_backup_20260420_143022.xlsx
머지 후 추출 파일 이관:
  90_Output/extracts/merged/20260420_143022/

진행할까요? (예/아니오)
```

사용자 "예" 답변 후에만 다음 단계 진행. 다른 답은 모두 중단.

## 7단계: 백업

마스터를 `90_Output/backups/master_backup_<YYYYMMDD_HHMMSS>.xlsx`로 복사.

- `shutil.copy2` 사용 (메타데이터 보존)
- `backups/` 폴더 없으면 생성
- 복사 실패 시 머지 중단 (디스크/권한 문제 사용자에게 보고)

## 8단계: 백업 정리 (Rolling 10)

`90_Output/backups/master_backup_*.xlsx` 파일을 mtime 내림차순 정렬.

- 11번째부터 삭제
- 백업 정리 실패는 경고만 출력하고 머지는 계속 진행 (백업 정리 실패가 머지를 막을 이유는 없음)

## 9단계: 직접 시트 머지 (insert_rows + 서식 보존) + 행 수 검증

v2.5.1부터는 **`insert_rows`(ASCENDING) + 템플릿 스타일 복사** 방식으로 처리한다. v1.1까지 사용하던 `delete_rows` + 재기입 방식은 데이터 행의 셀 서식(폰트·정렬·wrap_text 등)을 잃기 때문에 폐기됨. 핵심은 (a) 시프트 자동 처리, (b) 신규 셀에만 템플릿 스타일을 복사한다는 것.

```python
from copy import copy

# 1. 마스터 로드 (read-write)
wb_master = openpyxl.load_workbook(MASTER)

# 2. 머지 전 시트별 데이터행 수 기록 (master_before)
# 데이터행 = 헤더(1행) 제외, 모든 셀이 None/빈 문자열인 행은 제외

# 3. 시트별 row 2 셀 스타일을 col별 템플릿으로 추출
#    (font, fill, border, alignment, number_format, protection)
template = {}  # sheet -> col -> {style attrs}
for s in SHEETS:
    ws = wb_master[s]
    template[s] = {}
    for c in range(1, ws.max_column + 1):
        src = ws.cell(row=2, column=c)
        template[s][c] = {
            "font":          copy(src.font),
            "fill":          copy(src.fill),
            "border":        copy(src.border),
            "alignment":     copy(src.alignment),
            "number_format": src.number_format,
            "protection":    copy(src.protection),
        }

# 4. overwrite 대상(사용자 명시 재머지) 처리:
#    - 마스터에서 해당 study_id 행을 모두 삭제 (역순 ws.delete_rows)
#    - 제거 행 수 기록 (overwrite_removed)
#    NOTE: overwrite는 행 단위 삭제이므로 다른 행 서식 손실은 없음

# 5. 통과된 추출 파일을 번호(기본정보 A열) 오름차순 정렬 → sorted_extracts
#    동률(같은 번호)은 스캔 시점 순서 유지

# 6. sorted_extracts를 순서대로 시트에 직접 삽입 (ASCENDING)
for extract in sorted_extracts:
    n = extract.번호
    for s in ['기본정보', '아웃컴', '한의중재_한약']:
        ws = wb_master[s]
        max_col = ws.max_column
        ext_rows = [list(r) for r in extract[s]]   # 헤더 제외, 빈 행 제거
        if not ext_rows:
            continue
        # 4B 통과 케이스(마스터에 신규 컬럼 추가): 행 끝에 None 패딩
        ext_rows = [r + [None]*(max_col - len(r)) for r in ext_rows]

        # 삽입 위치: 현재 시트 상태에서 "처음으로 번호 > n" 인 xlsx row
        target = None
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v is not None and v > n:
                target = r
                break
        if target is None:
            # 마지막 데이터행 다음 (또는 시트가 비어 있으면 row 2)
            last = max((r for r in range(2, ws.max_row + 1)
                        if any(ws.cell(row=r, column=cc).value not in (None,'')
                               for cc in range(1, max_col+1))),
                       default=1)
            target = last + 1

        # ASCENDING 순서로 insert_rows: row >= target 셀을 자동 시프트(스타일 보존)
        ws.insert_rows(target, amount=len(ext_rows))

        # 신규 셀에 값 + 템플릿 스타일 복사
        for ri, row in enumerate(ext_rows):
            for ci in range(1, max_col + 1):
                cell = ws.cell(row=target + ri, column=ci, value=row[ci-1])
                t = template[s][ci]
                cell.font          = copy(t["font"])
                cell.fill          = copy(t["fill"])
                cell.border        = copy(t["border"])
                cell.alignment     = copy(t["alignment"])
                cell.number_format = t["number_format"]
                cell.protection    = copy(t["protection"])

# 7. 행 수 검증
#    expected = master_before + sum(extract 시트별 행 수) - overwrite_removed
#    actual   = 시트별 비-빈 데이터행 수
#    불일치 시 save 중단, 백업 복구 안내

# 8. 단일 atomic save
wb_master.save(MASTER)
```

**핵심 포인트**

- **ASCENDING 적용**: 추출을 번호 오름차순으로 처리하면서, 각 단계에서 *현재* 시트 상태를 기준으로 target을 다시 계산한다. DESCENDING은 마스터 max_row 이후 위치에 insert_rows를 호출할 때 시프트가 일어나지 않아 빈 행이 발생하므로 금지.
- **insert_rows의 시프트 동작**: `ws.insert_rows(idx, n)` 은 row ≥ idx인 셀을 n만큼 아래로 이동. 이동된 기존 셀의 스타일은 그대로 따라간다. 새로 만들어지는 빈 셀은 기본 스타일이므로 **반드시 템플릿을 명시 복사**해야 한다.
- **헤더(row 1)는 건드리지 않음**: insert_rows는 row 2 이상만 영향. row 1 서식은 자동 보존.
- **delete_rows 금지(데이터 영역)**: 데이터행 일괄 삭제 + 재기입 패턴은 v2.5.1부터 금지. overwrite의 행 단위 delete_rows는 허용(다른 행 영향 없음).

**삽입 위치 산정 예시**

기존 기본정보(번호): `[2, 3, 4, 6, 17, 30, 36]` (xlsx row 2~8), 신규 번호 `n=31`
→ 처음으로 v > 31 인 row = row 8 (값 36)
→ `ws.insert_rows(8, 1)` 후 row 8에 31 + 템플릿 스타일 기록
→ 결과: `[2, 3, 4, 6, 17, 30, 31, 36]` (row 2~9)

**기존이 비정렬인 경우**: linear scan 방식이라 "처음으로 번호가 커지는 위치"에 삽입된다. 기존 정렬이 깨져있어도 이 스킬은 기존 행을 재배치하지 않는다. 원칙 7 참조.

**아웃컴 블록 삽입**: 한 논문의 여러 아웃컴 행은 추출 파일 내 원래 순서대로 마스터의 같은 위치에 통째로 들어간다(`insert_rows(target, n)` 한 번에 n행). 내부 순서를 가르는 추가 정렬은 수행하지 않는다.

## 10단계: 사후 검증

저장된 마스터를 `read_only=True`로 다시 열어 행 수와 서식을 재확인한다.

- **행 수**: 9단계 검증 통과했어도 디스크 저장 후 무결성 재확인. 불일치 시 백업 자동 복구 + 사용자 보고.
- **한자/특수문자**: spot check로 확인 (예: 첫 번째 머지된 행의 `outcome_original` 컬럼 값 비교).
- **서식 보존 spot check** (v2.5.1):
  - 기존 행 (예: row 5 col 3) 의 font_name / font_size / alignment.horizontal / alignment.wrap_text 가 백업과 동일한지 확인.
  - 신규 행 (첫 머지 행) 의 동일 속성이 row 2 템플릿과 동일한지 확인.
  - 둘 중 하나라도 어긋나면 백업 복구 후 사용자 보고.

## 11단계: 추출 파일 이관

`90_Output/extracts/merged/<YYYYMMDD_HHMMSS>/` 폴더 생성.

- 머지 성공한 추출 파일을 모두 이동 (`shutil.move`)
- 거부된 파일은 원위치 유지 (사용자가 수정 후 재시도 가능)
- 이동 실패 시 경고만, 마스터는 이미 안전하게 머지된 상태이므로 머지 자체는 성공으로 처리

## 12단계: 머지 로그

`90_Output/extracts/merged/<YYYYMMDD_HHMMSS>/merge_log.txt`에 기록한다.

```
머지 시각: 2026-04-20 14:30:22
마스터: AF_CPG_data_extraction_심상송.xlsx
백업:   master_backup_20260420_143022.xlsx

머지 성공:
- AF_extract_30_Li_2025.xlsx (study_id=Li_2025)
   기본정보 +1, 아웃컴 +29, 한의중재_한약 +1
- AF_extract_31_Wang_2024.xlsx (study_id=Wang_2024)
   기본정보 +1, 아웃컴 +18, 한의중재_한약 +1

거부 (원위치 유지):
- AF_extract_32_Chen_2025.xlsx
   사유: 아웃컴 헤더 불일치 (추출에만 'effect_se' 존재)

Overwrite 대상:
- 없음

행 수 변화:
- 기본정보:        4 → 6
- 아웃컴:         49 → 96
- 한의중재_한약:    4 → 6
```

## 13단계: 결과 보고

사용자에게 다음 항목을 짧게 요약한다.

- 처리한 파일 수 (성공 / 거부)
- 시트별 추가 행 수
- 거부된 파일과 사유
- 백업 파일 경로
- 로그 파일 경로
- 마스터 파일 컴퓨터 링크

---

## 에러 처리 매트릭스

| 상황 | 대응 |
|------|------|
| 마스터 파일 없음 | 중단, 사용자에게 마스터 위치 질문 |
| 마스터 후보 다수 | 사용자에게 선택 질문 |
| 락 파일 존재 (`~$*.xlsx`) | 즉시 중단, "Excel 닫으세요" |
| 추출 파일 없음 | 정상 종료, "머지할 파일 없음" |
| 시트 구성 불일치 | 해당 파일만 거부, 나머지 진행 |
| 헤더 불일치 (자동 패딩 불가) | 해당 파일만 거부 |
| A열 '번호' 누락 / 비정수 | 해당 파일만 거부, "기본정보 A열 번호 필드가 비어있거나 정수가 아닙니다" |
| 시트 간 번호 불일치 | 해당 파일만 거부, "기본정보/아웃컴/한의중재_한약 시트의 번호가 일치하지 않습니다" |
| 중복 study_id (명시 없음) | 해당 파일만 거부 |
| 중복 study_id (overwrite 명시) | 마스터에서 해당 study_id 전 시트 행 삭제 후 신규 삽입 |
| 백업 실패 | 머지 중단, 사용자 보고 |
| 행 수 검증 실패 (save 전) | save 취소, 사용자 보고 |
| 사후 검증 실패 (save 후, 행 수) | 백업 자동 복구, 사용자 보고 |
| 사후 검증 실패 (save 후, 서식) | 백업 자동 복구, 사용자 보고. 9단계의 템플릿 추출/복사 누락 점검 |
| 추출 파일 이관 실패 | 경고, 머지는 성공 처리 |

---

## 절대 금지

- **사용자 승인 없는 자동 머지**: 미리보기 단계에서 명시적 "예" 없으면 진행 금지
- **백업 없는 마스터 수정**: 백업 실패는 머지 중단 사유
- **시트별 분할 save**: 원자성 위반, 중간 실패 시 데이터 불일치
- **추출 파일 삭제**: 머지 후 이동만, 삭제 금지
- **락 파일 무시 진행**: lost update 재발 보장
- **자동 overwrite**: 사용자 명시 없는 study_id overwrite 절대 금지
- **헤더 불일치 강제 진행**: 데이터 컬럼 시프트 → invisible corruption
- **맹목적 append** (v1.1 금지): 모든 신규 행은 반드시 번호 오름차순 위치에 삽입. 마스터 맨 뒤에 그냥 붙이는 방식 금지
- **기존 행 재배치**: 기존 마스터 데이터의 순서는 머지 스킬이 건드리지 않음 (정렬이 깨져있어도 수정 권한 없음)
- **`delete_rows` + 재기입** (v2.5.1 금지): 데이터 영역을 통째로 삭제 후 다시 쓰는 방식은 셀 서식(폰트·정렬·wrap 등)을 손실시킨다. 반드시 `insert_rows`(ASCENDING) + 템플릿 스타일 복사를 사용한다. 단, overwrite 처리에서의 행 단위 `delete_rows`(특정 study_id 행 제거)는 다른 행에 영향 없으므로 허용.
- **DESCENDING 적용**: 추출을 번호 내림차순으로 처리하거나 plans을 desc 정렬하여 insert_rows를 호출하면, 마스터 max_row 이후 위치에서는 시프트가 일어나지 않아 빈 행이 발생한다. 반드시 ASCENDING.
- **신규 셀 스타일 미복사**: insert_rows로 만들어진 빈 셀에 값만 쓰고 템플릿 스타일을 복사하지 않으면 기본 스타일(맑은 고딕 11pt 등)이 적용되어 기존 행과 시각적으로 어긋난다. 반드시 row 2 셀 스타일을 col별로 복사.

---

## 비고

- 본 스킬은 `cpg-data-extraction` 스킬과 페어링된다. 추출 스킬이 세션별 단독 파일 (`90_Output/extracts/AF_extract_<번호>_<study_id>.xlsx`)로 저장하면, 본 스킬이 일괄 머지한다.
- 머지 도중 lost update가 발생하지 않으려면 머지 시점에 다른 세션이 마스터를 건드리지 않아야 한다. 머지 작업은 단독 세션에서만 실행한다.
- 추출 파일에 형식 오류(예: heredoc 꼬리표 섞임)가 있어도 openpyxl이 정상 로드되면 머지 가능. 다만 데이터 무결성은 추출 단계에서 검증한 것을 신뢰한다.

---

## 변경 이력

### v2.6 (2026-05-19)
- cpg-data-extraction v2.6과 버전 정렬 (스킬 동작 변경 없음, 호환성 메모 추가)
- **호환성 주의**: cpg-data-extraction v2.6은 기본정보 시트가 47열 → 48열로 변경됨 (AU `analysis_set` 신열 + AV `notes`). v2.5.x 추출 파일을 v2.6 마스터에 직접 머지 시 4B 헤더 검증에서 거부된다. 사전에 `00_skills/cpg-data-extraction/scripts/migrate_to_v3.2.py`로 추출 파일을 48열로 마이그레이션한 뒤 머지할 것.
- 마스터·추출 파일 모두 48열로 정렬되어 있으면 머지 알고리즘은 v2.5.1과 동일하게 동작 (insert_rows ASCENDING + 템플릿 스타일 복사).

### v2.5.1 (2026-05-07)
- 9단계 머지 알고리즘을 `delete_rows`(데이터 영역 일괄) + 재기입 → **`insert_rows`(ASCENDING) + 템플릿 스타일 복사** 로 변경
- 데이터 행 셀 서식(Arial 9pt / 좌측·중앙 정렬 / wrap_text 등)이 머지 후 기본 스타일(맑은 고딕 11pt 등)로 바뀌던 결함 해결
- row 2의 셀 스타일을 col별 템플릿(font/fill/border/alignment/number_format/protection)으로 추출하여 신규 셀에 복사 적용
- 기존 행은 `insert_rows`의 자동 시프트로 스타일 자동 보존
- 절대 금지에 "delete_rows + 재기입(데이터 영역)", "DESCENDING 적용", "신규 셀 스타일 미복사" 추가
- 사고 사례: 2026-05-07 머지(`master_backup_20260507_021506.xlsx` 시점) 후 데이터 행 서식 손실 확인 → 백업 복구 후 본 알고리즘으로 재머지하여 정상화. 로그: `90_Output/extracts/merged/20260507_021506/merge_log.txt`

### v1.1 (2026-04-21)
- 머지 방식을 `append`(맨 뒤에 붙이기) → **번호 기반 정렬 삽입**으로 변경
- A열 `번호` 오름차순으로 신규 행을 삽입. 기존 행은 재배치하지 않음
- 같은 번호 내부(아웃컴 다수 행)는 추출 당시 순서(stable) 유지
- 4D 단계 신설: 번호 키 검증 (누락/비정수/시트 간 불일치 거부)
- 미리보기(6단계)에 각 추출의 번호 및 삽입 예상 행 번호 표시
- 에러 매트릭스에 "A열 번호 누락/비정수", "시트 간 번호 불일치" 추가
- 절대 금지 항목에 "맹목적 append", "기존 행 재배치" 추가

### v1.0 (2026-04-20)
- 초기 배포. append 기반 머지, 스키마 검증, 중복 차단, 사용자 승인, 원자 저장, 추출 파일 이관, 롤링 백업 10개.
