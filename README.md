# AF CPG Extraction Skill (Claude Code 버전)

심방세동(Atrial Fibrillation, AF) 한의표준임상진료지침(CPG) 개발 작업에서 **논문 데이터 추출**을 자동화하는 Claude **Code** 환경용 스킬입니다.

> Cowork 환경용 버전은 [af-cpg-extraction-skill-cowork](https://github.com/trisss1011/af-cpg-extraction-skill-cowork) 저장소에서 관리됩니다.

## 폴더 구조

| 폴더 | 설명 |
|------|------|
| `00_skills/` | Claude Code 스킬 본체 |
| `01_매뉴얼/` | 작업자용 사용 매뉴얼 |
| `02_papers/` | 추출 대상 논문 예시 |
| `10_참고자료/` | 논문 선정/배제 기록 등 |
| `90_Output/` | 추출 결과 템플릿 및 예시 |
| `START_HERE.txt` | 시작 안내 |

## 버전 이력

- **v3.0** — Claude Code 환경용 최초 배포 버전
- **v3.1** — 토의목록 append 기능 추가
- **v3.2** — cowork v2.6 변경사항 통합: study_design 분류 개정 (随机 표현만으로 RCT 인정), analysis_set 신규 AU열 (ITT/PP/NR, 48열), AF with RVR 코드 5 엄격화 (HR≥110), HRV 파생 지표 아웃컴 완전 제외, SAE 통합 추출, comorbidity 코드 6(高栓塞·高出血) 백포트, merge-skill 서식 보존 알고리즘 통합, scripts/save_extract.py가 sample_v2.6.xlsx(48열) 참조 (현재 버전)

각 버전은 Git 태그로 관리되며 [Releases](https://github.com/trisss1011/af-cpg-extraction-skill-code/releases)에서 zip으로 다운로드할 수 있습니다.

## 라이선스 / 이용

본 저장소는 심방세동 CPG 개발 프로젝트 내부 사용을 목적으로 하는 **Private 저장소**입니다.
