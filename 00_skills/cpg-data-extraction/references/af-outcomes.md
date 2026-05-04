# 심방세동(Atrial Fibrillation) CPG 표준 아웃컴 목록

> 최종 확정: 2026-03-26 연구팀 합의
> GRADE 방법론에 따라 Critical / Important로 중요도 분류

---

## 1. 확정 아웃컴 (Primary extraction targets)

### 1.1 심장리듬 관련

| 표준명 (국문) | 표준명 (영문) | 자료 유형 | 중요도 | 측정/비고 |
|-------------|-------------|---------|--------|---------|
| 심실박동수 | Ventricular rate (VR) / Heart rate (HR) | 연속형 (bpm) | **Critical** | ECG, Holter |
| AF 부담 | AF burden | 연속형 (%) | Important | Holter monitoring |
| 발작 빈도 | AF episode frequency | 연속형 (회/기간) | Important | Holter monitoring |
| 발작 지속시간 | AF episode duration | 연속형 (분, 시간) | Important | Holter monitoring |
| 동율동전환율 | Conversion rate to sinus rhythm | 이분형 (%) | Important | ECG |
| 동율동유지율 | Sinus rhythm maintenance rate | 이분형 (%) | Important | ECG at follow-up |
| AF 재발률 | AF recurrence rate | 이분형 (%) | Important | ECG, Holter |

> **주의**: VR과 HR은 논문마다 혼용되므로 원문 용어를 비고란에 기록하고, 표준명은 "VR/HR"로 통일.

### 1.2 심기능지표

| 표준명 (국문) | 표준명 (영문) | 자료 유형 | 중요도 | 측정/비고 |
|-------------|-------------|---------|--------|---------|
| 좌심실박출률 | Left ventricular ejection fraction (LVEF) | 연속형 (%) | Important | Echocardiography |
| 좌심방 직경 | Left atrial diameter (LAD) | 연속형 (mm) | Important | Echocardiography |

### 1.3 임상증상

| 표준명 (국문) | 표준명 (영문) | 자료 유형 | 중요도 | 척도 설명 |
|-------------|-------------|---------|--------|---------|
| 총유효율 (TER) | Total effective rate (TER) | 이분형 (%) | Important | 중국 TCM RCT에서 주로 사용. 순서형(현효·유효·무효) 보고 시 (현효+유효)=Event, 전체=Total로 이분형 변환 후 추출. 변환 시 비고란에 명시 |
| EHRA 점수 | EHRA symptom score | 연속형(평균) **및** 순서형(등급) 모두 추출 | Important | 1~4등급: 1=무증상, 2=경증, 3=중증, 4=중증장애. 높을수록 불량. 보고 형태(평균 또는 등급별 인원수)를 비고란에 기록 |

### 1.4 삶의 질

| 표준명 (국문) | 표준명 (영문) | 자료 유형 | 중요도 | 척도 설명 |
|-------------|-------------|---------|--------|---------|
| AFEQT 점수 | Atrial Fibrillation Effect on QualiTy of life | 연속형 (0~100) | Important | 높을수록 양호. 총점 및 하위 도메인(증상/일상활동/치료 관심도) 각각 추출 |

### 1.5 운동능력

| 표준명 (국문) | 표준명 (영문) | 자료 유형 | 중요도 | 비고 |
|-------------|-------------|---------|--------|------|
| 6분 보행검사 | 6-minute walk test (6MWT) | 연속형 (m) | Important | 운동능력 대리지표 |

### 1.6 안전성

| 표준명 (국문) | 표준명 (영문) | 자료 유형 | 중요도 | 비고 |
|-------------|-------------|---------|--------|------|
| 이상반응 전체 | Adverse events (total) | 이분형 (%) | Important | 종류, 건수, 중증도 기재. "없음" 또는 "보고 없음"도 기록 |
| 중대한 이상반응 | Serious adverse events (SAE) | 이분형 (%) | Important | 사망, 입원, 영구적 장애 포함 |
| 치료 중단율 | Discontinuation rate | 이분형 (%) | Important | 이상반응으로 인한 탈락자 수 / 전체 수 |

---

## 2. 아웃컴 명칭 표준화 규칙

동일 아웃컴에 대해 논문마다 다른 용어 사용 시 아래 표준명으로 통일.

| 논문에서 쓰는 용어 | 표준명 |
|-----------------|--------|
| HR, heart rate, 心率, 心室率 | VR/HR |
| AF burden, AF负荷, 房颤负荷 | AF 부담 |
| 发作频率, 发作次数, episode frequency | 발작 빈도 |
| 发作持续时间, episode duration | 발작 지속시간 |
| 转复率, 复律率, conversion rate, cardioversion rate, 窦性心律转复率 | 동율동전환율 |
| 维持率, sinus rhythm maintenance, 窦性心律维持率 | 동율동유지율 |
| 复发率, recurrence rate, 房颤复发率 | AF 재발률 |
| LVEF, 射血分数, ejection fraction | 좌심실박출률 (LVEF) |
| LAD, 左房内径, left atrial dimension, left atrial diameter | 좌심방 직경 (LAD) |
| adverse event, 不良反应, side effect, 副作用 | 이상반응 |
| 总有效率, 有效率, total effective rate, effective rate, response rate | 총유효율 (TER) |

---

## 3. 아웃컴 시트 C열 (outcome_std) 매핑표 (v1.13)

추출 엑셀 아웃컴 시트 C열에는 아래 **코드**를 기재한다. 약어 우선, 없으면 짧은 영문명.

### 3.1 표준 아웃컴 매핑

| 국문 표준명 (§1) | C열 기재값 | 중요도 |
|---|---|---|
| 심실박동수 | `VR/HR` | Critical |
| AF 부담 | `AF burden` | Important |
| 발작 빈도 | `AF episode frequency` | Important |
| 발작 지속시간 | `AF episode duration` | Important |
| 동율동전환율 | `SR conversion` | Important |
| 동율동유지율 | `SR maintenance` | Important |
| AF 재발률 | `AF recurrence` | Important |
| 좌심실박출률 | `LVEF` | Important |
| 좌심방 직경 | `LAD` | Important |
| 총유효율 (TER) | `TER` | Important |
| EHRA 점수 | `EHRA` | Important |
| AFEQT 점수 | `AFEQT` | Important |
| 6분 보행검사 | `6MWT` | Important |
| 이상반응 전체 | `AE total` | Important |
| 중대한 이상반응 | `SAE` | Important |
| 치료 중단율 | `Discontinuation` | Important |

### 3.2 비표준 아웃컴 규칙

표준 목록(§1)에 없는 아웃컴은 다음 규칙으로 C열 값을 결정한다:

1. **통용 약어** 있으면 약어 사용: `LVEDD`, `LVESD`, `BNP`, `Ang-II`, `NE`, `E/A ratio`
2. **약어 없으면** 짧은 영문명: `TCM symptom score`, `Platelet aggregation rate`
3. F열(importance)은 **공란**으로 둔다.
