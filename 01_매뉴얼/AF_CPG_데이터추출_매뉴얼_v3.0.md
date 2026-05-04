# 심방세동 한의표준임상진료지침
# 데이터 추출 스킬 배포 및 사용 매뉴얼

**매뉴얼 v3.0 · 2026. 04. 27.**
심방세동 한의CPG 개발팀

---

## 스킬 버전 정보

| 스킬 | 버전 | 날짜 |
|---|---|---|
| cpg-data-extraction | v3.0 | 2026-04-24 |
| merge-skill | v1.1 | 2026-04-21 |
| 매뉴얼 | v3.0 | 2026-04-27 |

주요 변경 사항은 각 스킬 폴더의 `CHANGELOG.md`를 참조한다. 스킬 업데이트 시 책임연구자가 공유 폴더에 새 버전을 올리고, 각 작업자는 로컬에 덮어쓰기한다.

---

## 1. 개요

이 매뉴얼은 심방세동 한의CPG 자료추출 작업에 사용되는 두 스킬 — **cpg-data-extraction(v3.0)** 과 **merge-skill(v1.1)** — 의 배포 및 사용 절차를 안내한다. 두 스킬은 **Claude Code(터미널 CLI)** 환경에서 동작한다. 본 매뉴얼은 Claude Code를 처음 쓰는 작업자가 환경 구축부터 추출·머지·검수까지 한 번에 따라할 수 있도록 작성되었다.

### 1-1. 산출물

- **추출 단계**: `90_Output/extracts/AF_extract_<번호>_<study_id>.xlsx` (논문 1건당 1파일)
- **병합 단계**: 마스터 엑셀 `AF_CPG_data_extraction_[작업자 이름].xlsx` (예: `AF_CPG_data_extraction_심상송.xlsx`)
- **CPG 근거표(evidence table)** : 후속 스킬에서 생성

### 1-2. v3.0 주요 변경 (요약)

| 영역 | v2.4 → v3.0 |
|---|---|
| 저장 방식 | Claude가 인라인으로 openpyxl 코드 작성 → **외부 Python 스크립트(`scripts/save_extract.py`) 호출** |
| 환경 요구 | (해당 없음) → **Python 3.8+ / openpyxl 3.0+ 필요** |
| 임시 파일 | 없음 → 추출 데이터를 `99_Scratch/_tmp_<번호>_<study_id>.json`으로 저장 후 스크립트 호출 |
| Z열 코드 | 4종 → **`other: <설명>` 코드 신설** (3군 연구·KM vs KM 비교 등 자유 서술) |
| 엑셀 서식 | (기본) → **freeze panes(헤더 1행 고정)**, **번호·연도 숫자 서식 `'0'`** 추가 |
| merge-skill | v1.0 append → **v1.1 번호 기반 정렬 삽입** (마스터 정렬 유지) |

호환성: v2.4 산출물과 컬럼·데이터 의미가 동일하므로 이전 추출물 재추출은 불필요하다. 서식 차이는 `migrate_format.py` 1회 실행으로 일괄 해소된다.

### 1-3. 배포 파일 목록

| 파일/폴더 | 역할 |
|---|---|
| `00_skills/cpg-data-extraction/SKILL.md` | 추출 스킬 본체 |
| `00_skills/cpg-data-extraction/CHANGELOG.md` | 추출 스킬 버전 이력 |
| `00_skills/cpg-data-extraction/sample_v2.4.xlsx` | 엑셀 출력 형식 샘플 (열 구조·서식 기준 템플릿) |
| `00_skills/cpg-data-extraction/references/` | 추출 상세 기준 (af-outcomes / korean-medicine-intervention / rob-extraction-fields) |
| `00_skills/cpg-data-extraction/scripts/save_extract.py` | **(v3.0 신규)** JSON → 엑셀 저장 스크립트 |
| `00_skills/cpg-data-extraction/scripts/migrate_format.py` | **(v3.0 신규)** v2.x 파일에 v3.0 서식 일괄 적용 |
| `00_skills/cpg-data-extraction/scripts/version_info.py` | **(v3.0 신규)** 엑셀 메타데이터에 스킬 버전 기록 |
| `00_skills/cpg-data-extraction/scripts/README.md` | 스크립트 환경 요구·사용법·트러블슈팅 |
| `00_skills/merge-skill/SKILL.md` | 머지 스킬 본체 |
| `01_매뉴얼/AF_CPG_데이터추출_매뉴얼_v3.0.docx` | 본 매뉴얼 |
| `01_매뉴얼/SESSION_STARTERS.md` | 작업 시작 템플릿 모음 |

### 1-4. 작업 폴더 체계

배포 파일을 압축 해제하면 다음 구조가 만들어진다.

```
AF_CPG_작업폴더/
├── 00_skills/                                    ← 스킬 본체
│   ├── cpg-data-extraction/
│   └── merge-skill/
├── 01_매뉴얼/                                    ← 본 매뉴얼·세션 템플릿
├── 02_papers/                                    ← 추출할 PDF 논문
├── 10_참고자료/                                  ← 선정·배제 관리 엑셀 등
├── 90_Output/
│   ├── AF_CPG_data_extraction_[작업자 이름].xlsx ← 마스터 (자기 것만)
│   ├── extracts/                                 ← 세션별 단독 추출 파일
│   │   └── merged/                               ← 머지 후 아카이브
│   └── backups/                                  ← 마스터 자동 백업 (최근 10개)
└── 99_Scratch/                                   ← 임시 JSON·로그·작업 메모
```

> **주의**: 작업 폴더는 **OneDrive·Dropbox 등 클라우드 동기화 폴더 안에 두지 않는다.** 동기화 중 파일 잠금·충돌이 생기면 추출·머지가 중단된다. 권장 위치: `C:\work\AF_CPG\` 같은 로컬 디스크 경로.

---

## 2. 사전 준비

이 장은 **처음 사용자**를 위한 일회성 환경 구축 절차다. 한 번 셋업하면 다음부터는 3장(작업 흐름)부터 따라가면 된다.

### 2-1. Claude Code 설치 (Windows)

Claude Code는 터미널에서 동작하는 Anthropic 공식 CLI다. Claude Desktop과 별개의 프로그램이다.

#### 설치 절차

1. **Node.js 설치 확인**

   Windows PowerShell을 열고 다음을 입력한다.

   ```powershell
   node --version
   ```

   `v18` 이상이 출력되면 OK. 명령이 인식되지 않으면 [https://nodejs.org/](https://nodejs.org/)에서 LTS 버전을 설치한다.

2. **Claude Code 설치**

   PowerShell에서 다음을 실행한다.

   ```powershell
   npm install -g @anthropic-ai/claude-code
   ```

   설치가 끝나면 다음으로 확인한다.

   ```powershell
   claude --version
   ```

   버전이 출력되면 설치 완료.

3. **첫 실행 및 로그인**

   원하는 빈 폴더에서 `claude`를 실행하고 안내에 따라 Anthropic 계정으로 로그인한다 (브라우저 창이 열린다). 팀 계정 이메일로 로그인할 것 (개인 계정 아님).

   ```powershell
   claude
   ```

#### Anthropic 가이드 확인

설치 단계에서 막히면 Anthropic 공식 가이드(공유 채널에 링크)를 참고한다. Windows에서 `npm` 명령이 잡히지 않으면 PowerShell을 **관리자 권한**으로 다시 열어 시도한다.

### 2-2. Python · openpyxl 설치 (Windows)

v3.0부터 추출 결과를 엑셀로 변환하는 단계가 외부 Python 스크립트로 빠졌다. 작업자 PC에 Python이 없으면 추출이 저장 단계에서 중단된다.

#### 설치 절차

1. **Python 설치 확인**

   PowerShell에서:

   ```powershell
   python --version
   ```

   - `Python 3.8` 이상 출력 → OK, 다음 단계로
   - `Python 3.7` 이하 또는 `Microsoft Store`에서 열라는 메시지 → 새로 설치 필요

2. **Python 신규 설치**

   [https://www.python.org/downloads/windows/](https://www.python.org/downloads/windows/)에서 **Python 3.12** 이상의 Windows installer (64-bit)를 받아 실행한다.

   설치 화면에서 반드시 다음 둘을 체크한다.

   - ☑ **Add python.exe to PATH** (환경변수 등록)
   - ☑ Install for all users (관리자 권한 있을 때)

   설치 후 PowerShell을 새로 열어 `python --version`이 정상 출력되는지 확인한다.

3. **openpyxl 설치**

   ```powershell
   python -m pip install openpyxl
   ```

   설치 후 다음으로 검증한다.

   ```powershell
   python -m pip show openpyxl
   ```

   `Version: 3.x.x` 가 보이면 OK.

4. **(선택) 한글 출력 인코딩 보정**

   Windows PowerShell에서 한글·한자 출력이 깨지는 경우, 다음 환경변수를 설정한다.

   ```powershell
   [Environment]::SetEnvironmentVariable("PYTHONUTF8", "1", "User")
   ```

   설정 후 PowerShell을 새로 열어 적용된다. 한 번만 하면 영구 적용.

#### 설치 검증

작업 폴더에서 다음을 실행해 스크립트가 살아있는지 확인한다.

```powershell
python 00_skills/cpg-data-extraction/scripts/save_extract.py --help
```

도움말 또는 사용법이 출력되면 준비 완료. `ModuleNotFoundError: No module named 'openpyxl'`가 나오면 3단계로 돌아간다.

### 2-3. 작업 폴더 다운로드 및 배치

1. 책임연구자가 공유한 배포 파일(`AF CPG skill 배포용_ver.3.0.zip`)을 다운로드한다.
2. **로컬 디스크의 OneDrive 밖 경로**에 압축 해제한다 (예: `C:\work\AF_CPG\`).
3. 압축 해제 후 1-4의 폴더 구조와 동일한지 확인한다.
4. 자기 마스터 파일이 없다면, `90_Output/AF_CPG_data_extraction_[작업자이름].xlsx` 템플릿을 받아 본인 이름으로 저장한다 (예: `AF_CPG_data_extraction_심상송.xlsx`).

### 2-4. Claude Code 첫 실행 (작업폴더 진입)

1. PowerShell을 연다.
2. 작업 폴더로 이동한다.

   ```powershell
   cd C:\work\AF_CPG
   ```

3. `claude` 실행.

   ```powershell
   claude
   ```

4. 처음 실행하면 폴더 신뢰(trust) 여부를 묻는다. **Yes**(이 폴더 신뢰함)로 답한다.
5. 프롬프트가 나타나면 Claude Code 세션 진입 완료.

> **참고**: 한 번 신뢰한 폴더는 다음부터 자동으로 진입된다. 매번 같은 작업 폴더에서 `claude`를 띄우면 된다.

### 2-5. 스킬 인식 확인

세션을 처음 열면 다음을 입력해 스킬이 인식되는지 확인한다.

```
00_skills 폴더에 있는 스킬 두 개를 확인하고, 각각 어떤 트리거에 반응하는지 알려줘.
```

Claude가 `cpg-data-extraction (v3.0)` 과 `merge-skill (v1.1)` 두 개를 인식하고 트리거 단어 목록을 답변하면 준비 완료. 인식하지 않으면 다음을 점검한다.

- 작업 폴더 위치가 올바른가? (`pwd` 또는 `Get-Location` 으로 확인)
- `00_skills/cpg-data-extraction/SKILL.md` 파일이 존재하는가?
- 매번 같은 작업 폴더에서 `claude`를 실행하고 있는가?

---

## 3. 작업 흐름

매일의 작업 흐름은 다음과 같다. 각 단계는 SESSION_STARTERS.md의 템플릿을 그대로 복붙해 사용해도 된다.

```
[아침] claude 진입 → 진행 상태 확인
   ↓
[추출] 단건 또는 병렬로 PDF 추출 → 채팅 검토 → 엑셀 저장
   ↓
[검토] 90_Output/extracts/ 의 단독 파일을 직접 열어 검수
   ↓
[저녁] merge-skill 로 자기 마스터에 머지
   ↓
[제출] 책임연구자가 지정한 주기로 자기 마스터를 제출
```

### 3-1. 단건 추출 (기본·권장)

처음 몇 편, 복잡한 논문(다군·다중 아웃컴·복잡한 한약 처방)에 권장한다.

#### 트리거 입력

```
02_papers/30_2023_LI_Effect_of_xxx.pdf 논문을
cpg-data-extraction 스킬 v3.0으로 추출해줘.

6단계(채팅 출력 → 연구자 확인)까지 먼저 진행하고,
내가 OK하면 scripts/save_extract.py로 저장해.
```

PDF 경로 대신 번호로 지정해도 된다.

```
02_papers/30번 논문 추출해줘
```

#### Claude의 동작 순서

1. PDF 읽기·연구설계 파악 (RCT/quasi-RCT/non-RCT)
2. 기본정보 47열 추출
3. RoB 2.0 근거 11열(AJ~AT) 수집
4. 한약 중재 추출 (한의중재_한약 시트)
5. 아웃컴 목록화 + 자료유형 판별
6. 아웃컴별 데이터 추출
7. **채팅창에 5블록 + ⚠️ 플래그 출력** → 사용자 OK 대기
8. (사용자 승인 후) JSON 임시 파일 → `save_extract.py` 호출 → 저장
9. 완료 보고

#### 채팅 출력 5블록 (Claude가 보여주는 형태)

| 블록 | 내용 | 원문 대조 위치 | 확인 포인트 |
|---|---|---|---|
| ① | 논문 서지정보 | 표지/초록 | 저자명, 연도, 저널명 |
| ② | 연구 특성 | Methods | 연구설계, 대상자 수, AF 유형 |
| ③ | 중재 정보 | Methods > 치료방법 | 처방명, 용량, 비교군 |
| ④ | 아웃컴 데이터 | Results 표/본문 | Mean, SD, Event, Total |
| ⑤ | 불확실·NR 항목 | — | ⚠️ 표시 항목 판단 |

#### 원문 대조 및 수정 요청

원문과 한 항목씩 비교한다. 수정·보완이 필요하면 자연어로 채팅창에 전달한다.

```
수정사항:
- VR after 값: E군 78.3±8.2 → 78.5±8.2 (표2 기준)
- SD/SE 불확실 항목: 원문에 SD로 명시되어 있음
- follow_up: NR → 치료 종료 후 3개월
```

⚠️ 표시 항목과 별도 "확인 필요" 목록은 모두 답해야 다음 단계로 넘어간다.

#### 엑셀 저장 (v3.0 신규 워크플로우)

작업자가 "저장해줘" 또는 "OK 저장 진행"이라 입력하면 Claude가 다음을 자동 수행한다.

```
1. 추출 데이터를 JSON 구조로 정리 (3시트 = 기본정보·아웃컴·한의중재_한약)
2. 99_Scratch/_tmp_<번호>_<study_id>.json 으로 저장
3. python 00_skills/cpg-data-extraction/scripts/save_extract.py <임시 JSON>
4. 스크립트:
     sample_v2.4.xlsx 복사 → 검증 → 값 채움 →
     서식 적용(freeze panes, 숫자 서식, RoB 음영) → 저장 →
     v3.0 메타데이터 태깅
5. 성공 시 임시 JSON 자동 삭제 (실패 시 디버깅용으로 보존)
```

#### 결과 파일

```
파일명 : AF_extract_<번호>_<study_id>.xlsx
저장   : 90_Output/extracts/
예     : AF_extract_30_Li_2025.xlsx
```

저장 완료 후 채팅창에 "저장 완료" 메시지가 출력되어야 다음 논문으로 넘어갈 수 있다.

#### 저장 실패 시 (Exit code 별 대응)

| 코드 | 의미 | 대응 |
|---|---|---|
| 0 | 성공 | — |
| 1 | JSON 읽기/파싱 오류 | `99_Scratch/_tmp_*.json` 내용 확인 |
| 2 | 검증 실패 (필수 키·시트 번호 불일치 등) | 추출 단계 재검토. 3시트의 `번호`·`study_id`가 일치하는지 확인 |
| 3 | 디스크/권한 오류 | 작업 폴더 권한, 디스크 여유 확인 |
| 4 | 템플릿(`sample_v2.4.xlsx`) 미존재 | `00_skills/cpg-data-extraction/` 안에 샘플이 있는지 확인 |

자세한 트러블슈팅은 `00_skills/cpg-data-extraction/scripts/README.md` 참조.

### 3-2. 병렬 추출 (속도 우선)

비슷한 품질의 RCT 여러 편을 빠르게 처리하고 싶을 때.

#### 트리거 입력

```
02_papers/에서 30~32번 논문(3편)을 cpg-data-extraction 스킬 v3.0으로
병렬 추출해줘.

규칙:
- 서브에이전트 1명 = PDF 1편 (절대 묶지 마)
- 각 서브에이전트는 프롬프트에 PDF 파일명 앞 숫자를 재확인하고 시작
- 추출 결과는 scripts/save_extract.py로 개별 저장
- 메인 세션에는 "완료/실패 상태 + 파일 경로"만 반환 (데이터 값 반환 금지)
- 전편 다 끝나면 ⚠️ 의심 항목만 모아서 한꺼번에 보고

병렬 처리 중 사용자 검토는 엑셀 파일을 직접 열어서 할 거야.
채팅 요약에만 의존하지 말고 엑셀로 저장까지 꼭 완료해줘.
```

#### 병렬 추출 시 혼입 방지 7규칙

| # | 규칙 |
|---|---|
| 1 | 서브에이전트 1명 = PDF 1편 (예외 없음) |
| 2 | 서브에이전트 프롬프트에 번호·파일명 명시 |
| 3 | 서브에이전트 시작 시 PDF 파일명 앞 숫자 재확인 |
| 4 | 서브에이전트는 추출값을 메인에 반환 금지, 파일만 저장 |
| 5 | 사용자 검토 시 엑셀 파일 직접 열기 (채팅 요약 의존 금지) |
| 6 | merge 시 번호/study_id 일치 검증 (스킬 내장) |
| 7 | 동시 서브에이전트 수 상한 = 10 |

#### 단계 확장 권장

처음에는 **3편**부터, 감 잡으면 **5편 → 10편**으로 단계 확장. 절대 "10편 한 서브에이전트에게" 금지 — 데이터 혼입 위험.

### 3-3. 이어서 작업 (재개·새 세션)

이전 세션이 끊겼거나 컨텍스트가 압축됐을 때.

```
99_Scratch/전략C_구축계획.md 또는 SESSION_STARTERS.md 먼저 읽고
현재 진행 상태 파악해.

이제 02_papers/<번호>_<파일명>.pdf 부터 이어서 추출해줘.
cpg-data-extraction 스킬 v3.0 사용.
```

새 Claude Code 창에서도 같은 작업 폴더로 들어가면 `90_Output/extracts/`의 기존 추출 파일이 그대로 남아 있다. 어디까지 했는지 폴더 내용으로 알 수 있다.

### 3-4. 머지 (자기 마스터 통합)

각 작업자는 본인 작업 폴더 내 `90_Output/extracts/`에 쌓인 단독 추출 파일을 자신의 마스터(`AF_CPG_data_extraction_[작업자 이름].xlsx`)에 직접 통합한다. 작업 폴더가 작업자별로 분리되어 있으므로 각자 자신의 마스터만 머지하면 lost update는 발생하지 않는다.

여러 작업자의 마스터를 하나로 합치는 최종 통합(grand merge)은 책임연구자가 별도로 수행한다. 작업자는 자기 마스터를 주기적으로 제출하기만 하면 된다.

#### 권장 머지 주기

| 구분 | 권장값 | 비고 |
|---|---|---|
| 기본 주기 | 하루 작업 종료 시 1회 | 퇴근 전 또는 작업 마무리 시점 |
| 누적 기준 | 5~10건 쌓이면 1회 | 한 task에서 5건 추출했다면 task 종료와 함께 머지 |
| 상한선 | 20건 쌓기 전 반드시 머지 | 거부되는 파일이 있을 때 원인 추적이 어려워짐 |
| 비권장 | 추출 1건마다 즉시 머지 | 백업·검증 오버헤드 누적 |

#### STEP 1. 머지 전 준비

- 그날 추출한 모든 논문에 대해 "저장 완료" 메시지를 받았는지 확인.
- `90_Output/extracts/` 폴더에 단독 파일들이 들어 있는지 직접 확인.
- 자신의 마스터(`AF_CPG_data_extraction_[작업자 이름].xlsx`)가 Excel에서 열려 있다면 **반드시 닫는다.** 열려 있으면 머지가 거부된다.
- Claude Code 세션이 2개 떠 있다면, 머지를 실행할 세션만 남기고 다른 세션에서는 같은 마스터에 머지하지 않는다.

#### STEP 2. 머지 트리거 입력

Claude Code 채팅창에 다음 중 한 문장을 입력한다.

```
오늘 추출한 파일들을 merge-skill로
AF_CPG_data_extraction_심상송.xlsx 에 병합해줘.

- 90_Output/extracts/의 AF_extract_*.xlsx 전부 스캔
- 미리보기 보여주고 내 승인받은 뒤 진행
- 번호 오름차순 정렬 삽입
- 백업·로그 생성
```

또는 짧게:

```
머지해줘
```

merge-skill이 자동 수행하는 것:

- `extracts/` 폴더 스캔 → 추출 파일 목록 확인
- 마스터/락 파일 확인 (Excel 열려있으면 즉시 중단)
- 스키마(시트 구성·헤더) 검증
- 번호 키 검증 (정수·시트 간 동일)
- study_id 중복 검증
- 미리보기 출력 후 사용자 승인 대기

#### STEP 3. 미리보기 확인 후 승인

Claude가 다음 형식으로 머지 계획을 보여준다.

```
머지 대상: 3개 파일 중 2개 통과 (번호 오름차순 처리)
─────────────────────────────────────────────
✓ AF_extract_30_Li_2025.xlsx (번호=30)
   기본정보 +1행, 아웃컴 +29행, 한의중재_한약 +1행
   삽입 예정: 기본정보 row8, 아웃컴 row73, 한의중재_한약 row8
✓ AF_extract_31_Wang_2024.xlsx (번호=31)
   기본정보 +1행, 아웃컴 +18행, 한의중재_한약 +1행
✗ AF_extract_32_Chen_2025.xlsx — 거부
   사유: 아웃컴 시트 헤더 불일치

마스터 행 수 변화 (예상)
─────────────────────────────────────────────
기본정보:        4 → 6
아웃컴:         49 → 96
한의중재_한약:    4 → 6

진행할까요? (예/아니오)
```

행 수 변화가 예상과 맞으면 **"예"** 라고 답한다. 다른 답은 모두 중단으로 처리된다.

"예" 입력 시 자동 수행되는 것:

- 마스터 자동 백업 → `90_Output/backups/master_backup_<시각>.xlsx` (최근 10개 보관)
- 메모리 머지 → 행 수 검증 → 단일 저장 (atomic save)
- 저장 후 재검증
- 머지 성공한 추출 파일을 `90_Output/extracts/merged/<타임스탬프>/` 로 자동 이동
- `merge_log.txt` 기록

#### STEP 4. 머지 후 확인

- Claude의 완료 보고에서 "기본정보 N → M행" 변화가 예상치와 일치하는지 확인.
- `90_Output/extracts/` 폴더가 비어 있고, 머지된 파일이 `extracts/merged/<타임스탬프>/`에 들어 있는지 확인.
- 필요 시 마스터를 Excel로 열어 새로 추가된 행이 정상인지 빠르게 스캔.
- 문제 발견 시 `90_Output/backups/`의 직전 백업으로 즉시 복구한다.

#### 재머지 (이미 마스터에 있는 논문 덮어쓰기)

이미 마스터에 있는 논문을 재추출했다면, 단순 머지로는 "study_id 중복" 사유로 거부된다. 명시적으로 재머지를 지시한다.

```
30번 재머지해줘
```

또는

```
Li_2025 덮어쓰기
```

merge-skill은 해당 study_id의 마스터 내 전체 시트 행을 삭제한 뒤 신규 데이터를 삽입한다 (부분 업데이트가 아니라 "기존 완전 삭제 후 신규 삽입").

#### 머지 거부 상황과 대응

| 상황 | 대응 |
|---|---|
| 시트 구성 불일치 | 추출 파일 수정 후 다음 머지에 다시 포함 |
| 헤더 불일치 (컬럼명·순서 다름) | `sample_v2.4.xlsx` 기준으로 재추출 |
| A열 `번호` 누락/비정수 | 재추출 (스크립트가 자동으로 채우므로 보통 추출 단계에서 누락 사고) |
| 시트 간 번호 불일치 | 추출 단계 재검토 (3시트 동일 번호 원칙) |
| study_id 중복 (이미 마스터 존재) | 재추출 의도면 "30번 재머지"로 명시 |
| 마스터 파일이 Excel에서 열려 있음 | Excel 닫고 머지 재시도 |
| 같은 마스터 동시 머지 | 한쪽 세션 머지가 끝날 때까지 대기 |

#### 자기 마스터 → 책임연구자 제출

자신의 마스터(`AF_CPG_data_extraction_[작업자 이름].xlsx`)는 머지 완료 시점에 책임연구자가 지정한 공유 경로(이메일/공유 드라이브)로 제출한다. 제출 주기는 책임연구자 안내에 따른다(예: 주 1회). 제출 후에도 자신의 로컬 마스터는 계속 누적해서 사용한다 (초기화하지 않음).

### 3-5. 작업 효율 권장사항

- **세션당 5건 권장**: 누적 컨텍스트가 길어질수록 응답 속도·정확도가 떨어진다. 5건 이상이면 새 세션을 연다.
- **2개 세션 병렬**: 한 세션이 응답을 생성하는 동안 다른 세션에서 다음 논문을 검토하면 처리량이 늘어난다. 단, 같은 마스터에 동시 머지는 금지(5장 참조).
- **세션 시작 상태 확인**:
  ```
  99_Scratch/전략C_구축계획.md 읽고 90_Output/extracts/ 폴더 상태 보여줘.
  ```
- **⚠️만 집중 확인**: 6단계 출력이 길면
  ```
  ⚠️ 플래그 달린 항목만 먼저 보여줘. 나머지는 엑셀에서 직접 볼게.
  ```
- **이상 느껴지면 즉시 중단**: `취소` 또는 `cancel` 입력. 다시 처음부터.
- **교차 검증**: 원문 대조가 어려운 논문은 Gemini 등 다른 LLM에 동일한 SKILL.md를 제공해 별도로 추출시킨 뒤 두 결과를 비교하는 방식도 유효하다.

---

## 4. 시트별 열 설명

추출 엑셀은 3개 시트로 구성된다. 각 시트의 `번호`(A열)·`study_id`(B열)는 동일해야 한다.

### 4-1. 기본정보 시트 (47열)

논문 1건당 1행. 서지·연구특성·중재·RoB 근거가 모두 한 행에 들어간다.

#### 서지정보 (A~I)

| 열 | 이름 | 내용 |
|---|---|---|
| A | 번호 | PDF 파일명 맨 앞 숫자 (예: 30) |
| B | study_id | 1저자_연도 (예: Li_2025) |
| C | author | 1저자명 (영문) |
| D | year | 출판 연도 |
| E | title | 논문 제목 |
| F | journal | 저널명 |
| G | vol_issue | 권/호 |
| H | pages | 쪽수 범위 |
| I | doi_pmid | DOI 또는 PMID |

#### 선정·배제 (J~K)

| 열 | 이름 | 내용 |
|---|---|---|
| J | exclude | Y/공백. Y이면 메타분석에서 배제. Y행 전체는 연분홍(`FADBD8`)으로 표시된다 |
| K | exclude_reason | exclude=Y인 경우 배제 사유 (PRISMA flow용) |

#### 연구 특성 (L~Y)

| 열 | 이름 | 내용 |
|---|---|---|
| L | country | 국가명만 (도시·지역 제외). 예: `中国` 또는 `China` |
| M | setting | 단일/다기관 + 입원/외래만. 예: `단일기관, 입원` |
| N | study_design | RCT / quasi-RCT / non-RCT / 관찰연구 |
| O | sample_size | 초기 E:C → 최종 분석 E:C |
| P | age_E / age_C | 평균±SD 또는 Median(IQR). E 먼저: `E: 값 / C: 값` |
| Q | sex_E / sex_C | `E: <남>M/<여>F / C: <남>M/<여>F` |
| R | af_type_code | AF 유형 코드 (부록 A-3) — **숫자만** |
| S | af_type_other | 특수 AF 유형 코드 (부록 A-4) — **숫자만** |
| T | af_type_text | AF 유형 원문 텍스트 |
| U | comorbidity_code | 동반질환 코드 (부록 A-5) — **숫자만**, 일반 AF=`NA` |
| V | comorbidity_text | 동반질환 원문. NA면 `NA` |
| W | tcm_pattern | 한의 변증 (예: 心氣虛, 痰瘀互結). 없으면 `NR` |
| X | disease_duration | 이환 기간. 시간 단위 영문약어 (`d`/`wk`/`mo`/`yr`) |
| Y | baseline_comparable | 기저치 동질성. 소문자 `p`. 예: `Yes (p>0.05)` |

#### 중재 정보 (Z~AG)

| 열 | 이름 | 내용 |
|---|---|---|
| Z | comparison_type | 비교 구조 코드 (부록 A-2) — **v3.0: `other:` 신설** |
| AA | intervention_E | 한약 처방명. 형식: `漢字 (romanization)` |
| AB | intervention_C | 대조군 중재 이름만 |
| AC | intervention_C_method | 대조군 복용/시행 방법. 형식: `투여경로, 용량×횟수/일` |
| AD | co_intervention | **차별 부가중재만**. 공통 배경치료 제외 |
| AE | treatment_period | 치료 기간 — 영문약어 (`4wk`, `3mo`) |
| AF | treatment_sessions | 총 치료 횟수 |
| AG | follow_up | 치료 종료 후 관찰 기간. 없으면 `NR` |

> **co_intervention(AD) 기재 예**:
> - 차별 부가중재 없음 → `NR`
> - 치료군만 부가 → `아미오다론 0.2g×1회/일 (치료군)`
> - 양쪽 모두 차별 → `아미오다론 (치료군); 생활습관 교육 (대조군)`

#### 기타·아웃컴 목록 (AH~AI)

| 열 | 이름 | 내용 |
|---|---|---|
| AH | funding | 연구비 출처 |
| AI | outcomes_reported | 논문에서 보고한 **모든** 아웃컴(표준 외 포함). **세미콜론** 구분 |

#### RoB 2.0 근거 (AJ~AT) — 연하늘(`D6EAF8`) 음영

RoB 2.0 평가의 근거가 되는 원문 발췌·요약을 기록한다. 평가 자체가 아니라 **근거 인용**이라는 점에 주의한다. 11열 모두 추출(누락 금지). 미보고는 `NR`.

| 열 | 이름 | 내용 |
|---|---|---|
| AJ | rob_d1_sequence | 무작위 배정 순서 생성 방법 (원문) |
| AK | rob_d1_concealment | 배정 순서 은폐 수단 (원문) |
| AL | rob_d1_baseline | 기저치 유의차 변수 존재 여부 |
| AM | rob_d2_blinding | 환자·시술자 눈가림 방식 |
| AN | rob_d2_analysis | 분석 집단 기준 (`ITT`/`PP`/`mITT`/NR) |
| AO | rob_d3_dropouts | 군별 초기→최종 인원 또는 탈락자 수 |
| AP | rob_d3_reasons | 탈락·결측 사유 |
| AQ | rob_d4_assessor | 결과 평가자 눈가림 |
| AR | rob_d5_protocol | 임상시험 사전 등록 번호 (예: ChiCTR…) |
| AS | rob_d5_outcome | Methods 계획 vs Results 보고 일치 여부 |
| AT | rob_other_coi | 연구비·COI 선언 문구 |

도메인 분류·예시는 `00_skills/cpg-data-extraction/references/rob-extraction-fields.md` 참조.

#### 작업자 메모 (AU) — v2.4 신규

| 열 | 이름 | 내용 |
|---|---|---|
| AU | notes | 작업자가 직접 기록하는 논문별 특이사항 (원문 수치 의심·중복 게재 의심·재검토 요청 등). **Claude는 추출 시 이 열을 공란으로 둔다 (NR도 기재 금지)** |

머지 시 이 열 내용은 그대로 마스터에 반영된다(작업자 메모 보존).

### 4-2. 아웃컴 시트 (22열)

논문에 보고된 **모든 아웃컴**을 대상으로 한다(표준 + 비표준). 한 논문에서 여러 아웃컴·시점을 보고하면 그만큼 행이 늘어난다.

#### 식별·분류 (A~G)

| 열 | 이름 | 내용 |
|---|---|---|
| A | 번호 | 기본정보의 번호와 동일 |
| B | study_id | 기본정보의 study_id와 동일 |
| C | outcome_std | 표준 아웃컴이면 `af-outcomes.md` 매핑표 코드. 비표준이면 약어/짧은 영문명 |
| D | outcome_original | 논문 원문 용어 (표준·비표준 모두 원문 그대로) |
| E | data_type | 연속형 / 이분형 |
| F | importance | 표준이면 `Critical`/`Important`. 비표준이면 **공란** |
| G | time_point | 측정 시점. `Baseline` 또는 시간값(`4wk`). 치료 종료 후 추적은 `FU_` 접두 (`FU_1yr`) |

#### 원시데이터 (H~M)

E군(치료군)·C군(대조군) 각각의 원시 수치. **가능하면 이 6열을 완전히 채우는 것이 가장 좋다.** 메타분석 소프트웨어가 이로부터 효과크기와 95% CI를 자동 계산한다.

| 열 | 이름 | 내용 |
|---|---|---|
| H | E_n | 치료군 분석 인원수 |
| I | E_val1 (Mean/Event) | 연속형: 평균 / 이분형: 사건 수 |
| J | E_val2 (SD/Total) | 연속형: SD / 이분형: 전체 수 |
| K | C_n | 대조군 분석 인원수 |
| L | C_val1 (Mean/Event) | 연속형: 평균 / 이분형: 사건 수 |
| M | C_val2 (SD/Total) | 연속형: SD / 이분형: 전체 수 |

#### 효과크기 (N~R)

N~R 5열은 "논문이 직접 계산해서 보고한 효과크기와 신뢰구간"을 옮기는 자리다. 원시데이터(H~M)가 모두 있으면 비워둬도 된다.

- 원시 완전 → N~Q 빈칸 허용. R(p_value)도 원문 명시 시만 기록
- 원시 없음/불완전 → 보고된 효과크기·CI·p값을 **반드시 추출**

| 열 | 이름 | 내용 |
|---|---|---|
| N | effect_type | OR / RR / HR / MD / SMD. 검정통계량(t·χ²·F·Z)은 N열 금지 — U(notes)에 |
| O | effect_value | 효과크기 점추정치 |
| P | ci_lower | 95% CI 하한 |
| Q | ci_upper | 95% CI 상한 |
| R | p_value | 원문 명시 p값. 부등호 그대로 (`<0.001`). 미보고 `NR` |

#### 방향성 (S~T)

| 열 | 이름 | 내용 |
|---|---|---|
| S | direction | 정량 데이터(H~Q)가 충분하면 빈칸 허용. 서술형 결과만 있으면 "치료군 유리/대조군 유리/차이없음" 중 하나 |
| T | event_direction | 이분형 행만. event 수가 많을수록 좋은지(부록 A-6). 연속형은 `N/A` |

#### 보조정보 (U~V)

| 열 | 이름 | 내용 |
|---|---|---|
| U | notes | Claude가 자동 기재. 데이터 변환 내역(예: `Mean±SE → SD 변환`), 검정통계량(t·χ²·F·Z), 세부 이상반응 |
| V | val2_type | 연속형 행 기본값 `SD` (또는 `SE`). **이분형 행은 공란** (val2가 Total이므로 SD/SE 개념 부적용) |

> **이상반응 기재**: 아웃컴 시트에 `AE total` 1행만 추출. 세부 유형은 U(notes)에 텍스트로 기록.

### 4-3. 한의중재_한약 시트 (8열)

한약 처방을 사용하는 논문 1건당 1행. 주처방만 기록, 부처방은 H(note)에 서술.

| 열 | 이름 | 내용 |
|---|---|---|
| A | 번호 | 기본정보의 번호와 동일 |
| B | study_id | 기본정보의 study_id와 동일 |
| C | formula_name_kr | 처방명 (한국어) |
| D | formula_name_cn | 처방명 (한자) |
| E | formulation | 탕약/과립/캡슐/환제/주사제/`NR` |
| F | composition | 약재+용량 원문 (쉼표 구분, 원문 순서 보존) |
| G | administration | 투여경로·용량·횟수. 예: `경구, 탕약 수전 400mL, 200mL×2회/일` |
| H | note | 비고 (부처방·NR 사유 등) |

> **F(composition) 기재 핵심 규칙**:
> 1. 형식: `약재명 + 공백 + 용량원문단위` (예: `麻黃 9g`)
> 2. 구분자 쉼표(`,`)+공백 통일. `、` `·` `；` 도 쉼표로 변환
> 3. **번역·단위환산 금지**. 원문 언어·단위 그대로
> 4. **법제 표기 유지**: `生·炒·炙·蜜炙·制·焦·煨` 등 포제 접두어는 원문 그대로
> 5. `各9g` 같은 공통 표기는 풀어서 개별 기록 (note에 명시)
> 6. 약재 전체 용량 미기재 → F 약재명만 나열, H에 `용량 미기재` 명시
> 7. 처방 전체 구성 비공개 → F `NR`, H에 사유

자세한 기준은 `references/korean-medicine-intervention.md` 참조.

### 4-4. 엑셀 서식 규칙 (v3.0 추가)

scripts/save_extract.py 가 자동 적용한다. 작업자가 직접 수정할 일은 없다.

#### 셀 서식

| 항목 | 헤더행 | 데이터행 |
|---|---|---|
| 폰트 | Arial 9, Bold | Arial 9 일반 |
| 정렬 | 가운데/가운데/줄바꿈 켜짐 | 왼쪽/가운데/줄바꿈 켜짐 |
| 테두리 | 없음 | 없음 |

#### 배경색 (우선순위: 헤더 > exclude=Y > RoB 열 > 일반)

| 대상 | 색상 |
|---|---|
| 모든 시트 헤더행 | `2F4F4F` (다크 틸) |
| exclude=Y 행 전체 | `FADBD8` (연분홍) |
| 기본정보 AJ~AT (RoB 근거) | `D6EAF8` (연하늘) |
| 그 외 데이터행 | 색 없음 |

#### 숫자 서식 `'0'` (v3.0 확장)

- 기본정보: A(번호), **D(year) (v3.0 신규)**, R(af_type_code), S(af_type_other), U(comorbidity_code)
- 아웃컴: **A(번호) (v3.0 신규)**
- 한의중재_한약: **A(번호) (v3.0 신규)**

#### freeze panes (v3.0 신규)

모든 시트의 헤더 1행 고정 (`freeze_panes='A2'`). 스크롤 시 헤더 상시 노출.

---

## 5. 주의사항

### 5-1. 왜 추출은 단독 파일에 쓰고, 머지는 merge-skill로만 하는가

openpyxl은 파일 전체를 메모리에 로드한 뒤 통째로 덮어쓰는 방식으로 동작한다. 같은 마스터 파일을 동시에 열고 쓰는 두 세션이 있으면 다음 시나리오가 발생한다.

```
T1  세션A: 마스터 로드 → 메모리 [Zhang, Zhou, Yang] 4행
T2  세션B: 마스터 로드 → 메모리 [Zhang, Zhou, Yang] 4행 (동일)
T3  세션A: Li_2025 append → 메모리 [Zhang, Zhou, Yang, Li] 5행
T4  세션A: save → 디스크 [Zhang, Zhou, Yang, Li] 5행 ✓
T5  세션B: Wang_2024 append → 메모리 [Zhang, Zhou, Yang, Wang] 5행
            (B는 Li_2025 추가 사실을 모름)
T6  세션B: save → 디스크 [Zhang, Zhou, Yang, Wang] 5행
            ↳ Li_2025 영구 소실
```

이 현상을 **lost update(쓰기 손실)** 이라고 한다. v2.3부터 작업 폴더를 작업자별로 분리하여 마스터를 각자 소유하게 했고, 추출은 논문별 단독 파일에만 저장하고 마스터로의 통합은 merge-skill이 원자적으로 수행하게 함으로써 이 문제를 해결했다.

단, 같은 작업자가 자기 마스터에 대해 두 세션에서 머지를 동시에 실행하면 여전히 lost update가 난다. 그래서 **"한 마스터 동시 1세션 머지" 규칙**을 준수해야 한다.

### 5-2. 절대 금지

- 여러 논문 동시 추출 요청 (데이터 혼입 위험)
- Claude 출력을 검증 없이 그대로 사용
- 원문에 없는 수치를 추측해서 채우기
- 추출 결과를 `extracts/` 단독 파일이 아닌 마스터에 직접 수기 입력·복사붙여넣기 (lost update 및 서식 변형 원인)
- 같은 마스터 파일에 대해 2개 이상 세션에서 머지 동시 실행
- 머지 중에 마스터 파일을 Excel로 열어두기
- 다른 작업자의 마스터(`AF_CPG_data_extraction_[다른 사람].xlsx`)에 머지하기
- **`sample_v2.4.xlsx` 직접 수정** — 스크립트가 매번 이 파일을 복사하므로 원본 손상 금지
- **`scripts/*.py` 직접 수정** — 스킬 업데이트 시에만 일괄 교체
- **`90_Output/extracts/`의 `.bak.xlsx` 삭제** — 마이그레이션 복구용
- **OneDrive 동기화 중에 저장 시도** — 락 충돌 가능, 잠시 기다린 뒤 시도

### 5-3. 반드시 준수

- 1논문 1건 → 채팅 확인 → 단독 파일 저장 순서를 고수
- ⚠️ 표시 항목은 반드시 원문 확인 후 판단
- NR(Not Reported)은 정말로 보고되지 않은 경우에만 유지 (추측 금지)
- 저장된 추출 파일은 머지 전에 한 번 더 엑셀에서 직접 검수
- 머지는 하루 1회 또는 5~10건 누적 시 수행. 20건 이상 쌓기 전에는 반드시 실행
- 머지는 **자기 마스터에만**, 한 시점에 **한 세션에서만** 실행
- v3.0의 `other:` 코드는 반드시 **소문자 `other:` + 공백** 으로 시작 (예: `other: 3-arm (KM / WM / KM+WM 세 군 비교)`)

### 5-4. 혼입 방지 7규칙 (병렬 추출 시 필수)

| # | 규칙 |
|---|---|
| 1 | 서브에이전트 1명 = PDF 1편 (예외 없음) |
| 2 | 서브에이전트 프롬프트에 번호·파일명 명시 |
| 3 | 서브에이전트 시작 시 PDF 파일명 앞 숫자 재확인 |
| 4 | 서브에이전트는 추출값을 메인에 반환 금지, 파일만 저장 |
| 5 | 사용자 검토 시 엑셀 파일 직접 열기 (채팅 요약 의존 금지) |
| 6 | merge 시 번호/study_id 일치 검증 (스킬 내장) |
| 7 | 동시 서브에이전트 수 상한 = 10 |

---

## 6. 트러블슈팅 / FAQ

### 6-1. 환경·설치

**Q. `claude --version`이 명령을 찾을 수 없다고 한다.**

→ Claude Code가 설치되지 않았거나 PATH에 잡히지 않았다. PowerShell을 새로 열어 `npm list -g @anthropic-ai/claude-code`로 설치 여부를 확인하고, 미설치면 2-1로 돌아간다.

**Q. `python` 명령을 입력하면 Microsoft Store가 열린다.**

→ Windows 10/11의 기본 동작. 실제 Python을 설치하지 않고 스텁만 동작 중이다. [python.org](https://www.python.org/downloads/windows/)에서 정식 설치하고 "Add python.exe to PATH" 체크. 설치 후 PowerShell을 새로 열어야 적용된다.

**Q. `ModuleNotFoundError: No module named 'openpyxl'`**

→ openpyxl이 설치되지 않았다. `python -m pip install openpyxl` 실행. 설치 후에도 같은 에러면 다른 버전의 Python이 우선 인식되고 있을 가능성. `where python` (PowerShell)으로 어떤 python이 사용 중인지 확인하고, `pip` 대신 같은 python으로 `python -m pip install openpyxl`.

**Q. PowerShell에서 한글·한자가 깨진다.**

→ 환경변수 `PYTHONUTF8=1` 영구 설정. 2-2 4단계 참조. 일시적으로는 `$env:PYTHONUTF8="1"` 만 입력해도 그 세션 동안 적용된다.

### 6-2. 추출

**Q. Claude가 스킬을 인식하지 못한다.**

→ 작업 폴더 위치 확인 (`pwd` / `Get-Location`). `00_skills/cpg-data-extraction/SKILL.md`와 `00_skills/merge-skill/SKILL.md` 경로가 정확한지 확인. 폴더에서 `claude` 다시 실행.

**Q. PDF를 읽지 못한다.**

→ 스캔 PDF(이미지 기반)일 수 있다. OCR 처리된 텍스트 PDF인지 확인. 텍스트 PDF인데도 안 읽히면 `02_papers/` 경로가 정확한지, 파일명에 특수문자가 없는지 확인.

**Q. 추출 도중 "임시 JSON 파일을 저장할 수 없다"는 에러.**

→ `99_Scratch/` 폴더가 없거나 권한이 없다. 폴더 직접 생성하거나 권한 확인.

**Q. `save_extract.py`가 exit code 2(검증 실패)로 거부했다.**

→ 임시 JSON이 `99_Scratch/_tmp_*.json`에 보존되어 있다. 이를 Claude에게 전달:
```
방금 추출한 논문의 임시 JSON이 99_Scratch/_tmp_*.json에 남아있어.
save_extract.py가 왜 거부했는지 확인하고 고쳐줘.
```
보통은 (a) 3시트의 `번호`/`study_id` 불일치, (b) 필수 키 누락, (c) 타입 오류 중 하나다.

**Q. 아웃컴 시트에 없는 아웃컴이 논문에 있다.**

→ v3.0은 **모든 아웃컴**을 추출한다 (표준 + 비표준). 비표준 아웃컴은 F(importance)를 공란으로 두고 C(outcome_std)에 약어/짧은 영문명을 기재한다. 기본정보 AI(outcomes_reported)에는 모든 아웃컴이 세미콜론으로 나열된다.

**Q. 세션이 끊어졌는데 이어서 작업할 수 있나?**

→ 새 세션에서 같은 폴더를 연결하면 `extracts/` 폴더의 기존 추출 파일이 그대로 남아있다. 이어서 다음 논문 추출을 요청하면 된다. 이전 채팅 맥락은 유지되지 않으므로 어떤 논문까지 완료했는지 알려준다 (3-3 템플릿).

### 6-3. 머지

**Q. 추출 파일을 마스터에 직접 복사-붙여넣기 해도 되나?**

→ 권장하지 않는다. 수동 복사는 행 순서·서식·특수문자 변형 위험. 반드시 merge-skill을 통해 병합한다. merge-skill은 백업·스키마 검증·행 수 검증을 자동 수행한다.

**Q. 동일 논문을 두 번 추출했다.**

→ 같은 study_id로 두 개의 추출 파일이 생긴다. 머지 시 먼저 처리된 파일만 통합되고 나중 파일은 "study_id 중복"으로 거부된다. 재추출이 의도라면 "30번 재머지" 또는 "Li_2025 덮어쓰기"로 명시한다.

**Q. 머지 중에 에러가 났다.**

→ `90_Output/backups/` 폴더의 최신 백업으로 마스터를 복구할 수 있다. merge-skill은 머지 직전 자동 백업을 만들며 최근 10개를 보관한다.

**Q. 내 추출 파일이 머지에서 거부됐다.**

→ 거부 사유는 미리보기에 표시된다. 대부분 헤더/시트 구성 불일치 또는 `sample_v2.4.xlsx` 기준 미준수가 원인. 사유에 맞춰 재추출 후 다음 머지에 다시 포함한다. v2.x 파일이라면 `migrate_format.py`로 서식 마이그레이션 시도:

```
90_Output/extracts/ 폴더의 v2.x 파일들을 v3.0 서식으로 일괄 마이그레이션해줘.
migrate_format.py 사용. dry-run 먼저 보여주고 승인받은 뒤 실행.
```

**Q. merge-skill이 "마스터 파일이 Excel에서 열려 있습니다"라며 거부한다.**

→ 마스터 파일이 Excel에서 열려 있거나, 비정상 종료로 락 파일(`~$AF_CPG_data_extraction_*.xlsx`)이 남아있다. Excel 닫고, 폴더에 `~$` 로 시작하는 임시 파일이 남아있으면 삭제 후 재시도.

### 6-4. 데이터·서식

**Q. 엑셀 서식이 다른 작업자와 다르다.**

→ `sample_v2.4.xlsx`가 `00_skills/cpg-data-extraction/`에 있는지 확인. v3.0 스크립트가 이 샘플을 매번 복사하므로 누락되면 저장 자체가 거부된다 (exit code 4).

**Q. 추출 결과 파일명은 어떻게 되나?**

→ `AF_extract_<번호>_<study_id>.xlsx`. 예: `AF_extract_30_Li_2025.xlsx`. `<번호>`는 PDF 파일명 맨 앞 숫자, `<study_id>`는 기본정보 시트의 B열 값.

**Q. v2.x로 추출한 기존 파일을 v3.0으로 일괄 변환하려면?**

→ `migrate_format.py`. 데이터 값은 건드리지 않고 서식만 바꾼다. dry-run으로 먼저 확인:
```powershell
python 00_skills/cpg-data-extraction/scripts/migrate_format.py 90_Output/extracts --dry-run
```
문제없으면 실제 실행:
```powershell
python 00_skills/cpg-data-extraction/scripts/migrate_format.py 90_Output/extracts
```
`.bak.xlsx` 백업이 자동 생성된다.

**Q. v3.0의 `comparison_type: other:` 는 언제 쓰나?**

→ 기존 4종 코드(`KM_alone_vs_placebo`, `KM_alone_vs_WM`, `KM+WM_vs_WM`, `KM+WM_vs_KM`)에 해당하지 않을 때. 3군 이상 연구나 KM vs KM 비교 등. 반드시 `other: ` (소문자+콜론+공백) 으로 시작한다.

예시:
- `other: 3-arm (KM / WM / KM+WM 세 군 비교)`
- `other: KM vs 다른 한약 (Shensong Yangxin vs Wenxin Keli)`
- `other: 위약 대조 add-on (placebo + WM)`

---

## 7. 부록 — 색상 범례 및 코드표

추출 엑셀에서 사용하는 색상과 코드 체계. 코드 체계 변경 시에만 업데이트한다.

### A-1. 색상 범례

| 색상 | 의미 |
|---|---|
| `2F4F4F` (다크 틸) | 모든 시트 헤더행 |
| `FADBD8` (연분홍) | exclude=Y 행 (배제 연구, 메타분석 제외) |
| `D6EAF8` (연하늘) | 기본정보 AJ~AT열 (RoB 근거) |

### A-2. comparison_type (Z열)

| 코드 | 의미 |
|---|---|
| `KM_alone_vs_placebo` | 한방 단독 vs 위약/무처치 |
| `KM_alone_vs_WM` | 한방 단독 vs 양방 단독 |
| `KM+WM_vs_WM` | 한방+양방 병행 vs 양방 단독 |
| `KM+WM_vs_KM` | 한방+양방 병행 vs 한방 단독 |
| `other: <설명>` **(v3.0 신규)** | 위 4개 외 — 자유 서술. `3-arm` 태그 권장 |

### A-3. af_type_code (R열)

| 코드 | 의미 |
|---|---|
| 1 | 발작성 AF |
| 2 | 지속성 AF |
| 3 | 영구성 AF |
| 4 | 혼합/미특정 |

### A-4. af_type_other (S열)

| 코드 | 의미 |
|---|---|
| 5 | AF with RVR |
| 6 | NVAF — 논문에 NVAF 명시된 경우만 |
| 7 | 기타 (post-RFCA, 판막성 AF 등) |

### A-5. comorbidity_code (U열)

| 코드 | 의미 |
|---|---|
| `NA` | 일반 AF (동반질환 포함 요건 아님) |
| 1 | RFCA (af_type_other와 중복이어도 반드시 1로 기재) |
| 2 | 심부전 |
| 3 | 심근경색/관상동맥질환 |
| 4 | 갑상선/기타 |
| 5 | 고혈압 |

복수 동반질환은 쉼표로 구분 (예: `2,5`).

### A-6. event_direction (T열, 아웃컴)

| 값 | 용도 |
|---|---|
| `Event=유효` | 총유효율, 동율동전환율 등 긍정적 결과가 Event |
| `Event=재발(불량결과)` | AF 재발률 등 부정적 결과가 Event |
| `Event=이상반응(불량결과)` | 이상반응 발생 |
| `N/A` | 연속형 아웃컴 |

위 4가지 외 값 사용 금지.

### A-7. RoB 2.0 근거열 (AJ~AT, 기본정보)

| 열 | 내용 |
|---|---|
| `rob_d1_sequence` | 무작위 배정 순서 생성 방법 (원문) |
| `rob_d1_concealment` | 배정 순서 은폐 수단 (원문) |
| `rob_d1_baseline` | 기저치 유의차 변수 존재 여부 |
| `rob_d2_blinding` | 환자·시술자 눈가림 방식 |
| `rob_d2_analysis` | 분석 집단 기준 (`ITT`/`PP`/`mITT`) |
| `rob_d3_dropouts` | 군별 초기→최종 인원 또는 탈락자 수 |
| `rob_d3_reasons` | 탈락·결측 사유 |
| `rob_d4_assessor` | 결과 평가자 눈가림 |
| `rob_d5_protocol` | 임상시험 사전 등록 번호 |
| `rob_d5_outcome` | Methods 계획 vs Results 보고 일치 여부 |
| `rob_other_coi` | 연구비·COI 선언 문구 |

### A-8. notes (AU열, 기본정보) — v2.4

기본정보 시트의 마지막 열. 작업자가 추출 후 해당 논문의 특이사항을 자유 기술하는 용도. **Claude는 추출 시 이 열을 공란으로 둔다 (NR도 기재 금지).** 논문별 메모(원문 수치 의심·중복 게재 의심·재검토 요청 등)는 작업자가 추출 파일을 검토한 뒤 직접 입력한다.

아웃컴 시트의 notes(U열)와 이름은 같지만 용도가 다르다. 아웃컴 notes는 데이터 변환 내역·세부 이상반응 기록용이며 Claude가 자동 기재한다.

### A-9. val2_type (V열, 아웃컴)

- 연속형 행 기본값: `SD` (또는 `SE`로 명시)
- 이분형 행: 공란 (val2가 Total이므로 SD/SE 개념 부적용)
- SD/SE 불명확 시: `SD`로 기재 + U(notes)에 `"SD/SE 불명확"` 기록

### A-10. 트리거 단어 치트시트

#### cpg-data-extraction 호출

```
추출 / 추출해줘 / 추출시행 / extraction / 논문 데이터 추출 /
자료추출 / data extraction / 메타분석 데이터 / 아웃컴 추출 /
한약 추출 / 처방 구성 / Mean SD / event total
```

#### merge-skill 호출

```
머지 / 병합 / 추출 머지 / 추출 병합 / 마스터에 합쳐줘 /
추출 파일 합쳐줘 / AF_extract 합쳐줘 / 세션별 파일 합치기 /
추출 결과 통합 / merge extracts / consolidate extracts
```

### A-11. SESSION_STARTERS — 빠른 시작 템플릿

자세한 본문은 `01_매뉴얼/SESSION_STARTERS.md`. 여기서는 4가지 핵심 템플릿만 발췌.

#### 단건 추출

```
02_papers/<파일명>.pdf 논문을
cpg-data-extraction 스킬 v3.0으로 추출해줘.

6단계(채팅 출력 → 연구자 확인)까지 먼저 진행하고,
내가 OK하면 scripts/save_extract.py로 저장해.
```

#### 병렬 추출

```
02_papers/에서 <N>~<M>번 논문(<총편수>편)을 cpg-data-extraction 스킬 v3.0으로
병렬 추출해줘.

규칙:
- 서브에이전트 1명 = PDF 1편 (절대 묶지 마)
- 각 서브에이전트는 프롬프트에 PDF 파일명 앞 숫자를 재확인하고 시작
- 추출 결과는 scripts/save_extract.py로 개별 저장
- 메인 세션에는 "완료/실패 상태 + 파일 경로"만 반환 (데이터 값 반환 금지)
- 전편 다 끝나면 ⚠️ 의심 항목만 모아서 한꺼번에 보고
```

#### 이어서 작업

```
99_Scratch/전략C_구축계획.md 또는 SESSION_STARTERS.md 먼저 읽고
현재 진행 상태 파악해.

이제 02_papers/<번호>_<파일명>.pdf 부터 이어서 추출해줘.
cpg-data-extraction 스킬 v3.0 사용.
```

#### 머지

```
오늘 추출한 파일들을 merge-skill로
AF_CPG_data_extraction_심상송.xlsx에 병합해줘.

- 90_Output/extracts/의 AF_extract_*.xlsx 전부 스캔
- 미리보기 보여주고 내 승인받은 뒤 진행
- 번호 오름차순 정렬 삽입
- 백업·로그 생성
```

---

## 문의

- 스킬 자체의 버그·이상 동작: 책임연구자에게 연락
- 추출 데이터 판단 회의(원문 수치 해석·SD/SE 모호 등): 책임연구자 검토 요청
- Claude Code 환경 문제: Anthropic 공식 가이드 또는 책임연구자 공유 채널

---

*본 매뉴얼은 cpg-data-extraction v3.0 / merge-skill v1.1 기준이다. 스킬이 갱신되면 매뉴얼도 함께 개정한다. 매뉴얼 버전과 스킬 버전이 어긋나면 SKILL.md·CHANGELOG.md를 우선한다.*
