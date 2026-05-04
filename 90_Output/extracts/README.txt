이 폴더에는 cpg-data-extraction 스킬이 생성한 논문별 단독 추출 파일이 쌓입니다.

파일명 규칙
  AF_extract_<번호>_<study_id>.xlsx
  예) AF_extract_30_Li_2025.xlsx

처리 흐름
  1) Claude가 추출 후 save_extract.py로 이 폴더에 저장
  2) merge-skill이 이 폴더를 스캔해서 마스터에 병합
  3) 머지 성공한 파일은 자동으로 merged/<타임스탬프>/ 로 이동

주의
  - 직접 파일을 편집하지 마세요. 머지 단계에서 검증 거부될 수 있습니다.
  - 머지 직전에 파일이 모두 닫혀 있어야 합니다 (Excel에서 열어두지 말 것).
  - 본 README는 압축·배포용 placeholder 입니다. 삭제해도 무방합니다.
