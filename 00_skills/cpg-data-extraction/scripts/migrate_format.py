#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
migrate_format.py — v3.2 (cpg-data-extraction 서식 마이그레이션)

v2.4 이전 추출 파일(`AF_extract_*.xlsx`)에 v3.0+ 추가 서식을 일괄 적용.
유형 2(서식만 변경) 변경에 한정 — 데이터 값은 절대 건드리지 않음.

v3.2 주의: 헤더가 47열 → 48열로 변경됨 (AU=analysis_set, AV=notes). 47열 추출 파일은
먼저 `migrate_to_v3.2.py`로 48열 마이그레이션 후 본 스크립트 적용.

적용 항목 (⑤ B+ 신규 서식):
  - freeze_panes = 'A2' (3시트 모두)
  - 번호 열(A) 숫자 서식 '0' (3시트 모두)
  - 기본정보 D열(year), R·S·U열 숫자 서식 '0'
  - 아웃컴·한의중재_한약 A열 숫자 서식 '0'
  - 기본정보 AJ~AT열 (RoB) 배경 `D6EAF8` 재확인
  - version_info 태깅

안전장치:
  - 작업 전 .bak 백업 자동 생성
  - 이미 v3.0 태그된 파일은 skip (재실행 안전)
  - --dry-run 지원
  - 데이터 셀 값 ABSOLUTELY 건드리지 않음

사용법:
    python migrate_format.py <file_or_dir> [--dry-run] [--force]

    <file_or_dir>  : 단일 xlsx 파일 경로 또는 폴더 (폴더면 AF_extract_*.xlsx 전체 스캔)
    --dry-run      : 실제 저장 없이 변경 예정 내용만 보고
    --force        : 이미 v3.0 태그된 파일도 재적용

Exit codes:
    0  성공 (변경된 파일 수 표시)
    1  인자 오류
    2  대상 파일 없음
    3  파일 쓰기 오류
"""

import sys
import shutil
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# 같은 scripts/ 폴더의 version_info 모듈
try:
    from version_info import set_version, get_version, SKILL_VERSION
except ImportError:
    # 폴백: 직접 상수
    SKILL_VERSION = 'v3.2'
    def set_version(wb, version=SKILL_VERSION):
        wb.properties.keywords = (wb.properties.keywords or '') + f' cpg-skill={version}'
    def get_version(wb):
        kws = wb.properties.keywords or ''
        for tok in kws.split(','):
            if 'cpg-skill=' in tok.strip():
                return tok.strip().split('=', 1)[1]
        return None


# ============================================================
# 설정
# ============================================================

ROB_COL_START = 36   # AJ
ROB_COL_END   = 46   # AT (inclusive)
ROB_FILL_RGB  = '00D6EAF8'

# 숫자 서식 '0' 대상 (시트별 1-indexed 컬럼)
NUM_FMT_ZERO = {
    '기본정보': [1, 4, 18, 19, 21],   # A, D, R, S, U
    '아웃컴': [1],                     # A
    '한의중재_한약': [1],              # A
}


# ============================================================
# 유틸
# ============================================================

def log(msg):
    try:
        print(msg)
    except UnicodeEncodeError:
        sys.stdout.buffer.write((msg + '\n').encode('utf-8'))
        sys.stdout.buffer.flush()


def stdout_utf8():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


def collect_target_files(path: Path):
    """단일 파일 or 폴더에서 AF_extract_*.xlsx 수집."""
    if path.is_file():
        return [path]
    if path.is_dir():
        files = sorted(path.glob('AF_extract_*.xlsx'))
        # 임시 파일·백업 제외
        files = [f for f in files if not f.name.startswith('~$') and not f.name.endswith('.bak.xlsx')]
        return files
    return []


# ============================================================
# 마이그레이션 로직
# ============================================================

def has_ghost_rows(wb):
    """어느 한 시트라도 A열(번호)이 None인 행이 있는지."""
    for sn in ('기본정보', '아웃컴', '한의중재_한약'):
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value is None:
                return True
    return False


def needs_migration(wb, force):
    """이 워크북에 마이그레이션이 필요한지 판단.
    - force → 항상 True
    - 유령 빈 행 있음 → True (버전 무관하게 정리)
    - 버전 태그 없음/다름 → True
    - v3.0 태그 + 유령 행 없음 → False (skip)."""
    if force:
        return True, 'force 옵션'
    if has_ghost_rows(wb):
        return True, '유령 빈 행 감지 (정리 필요)'
    current = get_version(wb)
    if current == SKILL_VERSION:
        return False, f'이미 {SKILL_VERSION} 태그됨'
    if current is None:
        return True, '버전 태그 없음 (v2.x 추정)'
    return True, f'이전 버전: {current}'


def apply_v3_formatting(wb):
    """v3.0 서식 적용 + 유령 빈 행(A열 '번호'가 None) 정리.
    데이터 값은 건드리지 않음 (번호가 있는 행은 모두 보존).
    반환: 변경 요약 dict."""
    changes = {
        'ghost_rows_removed': {},   # {sheet_name: count}
        'freeze_panes': [],
        'number_format': [],
        'rob_fill': 0,
    }

    for sn in wb.sheetnames:
        if sn not in NUM_FMT_ZERO:
            continue
        ws = wb[sn]

        # Step 1: 유령 빈 행 정리 (A열 '번호'가 None인 행)
        # v2.x 저장 방식이 템플릿 샘플 데이터를 덮어쓰되 여분 행을 지우지 않아
        # 셀 값은 None이지만 셀 자체는 남는 경우가 있음. 이를 물리적으로 제거.
        ghost_rows = []
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value is None:
                ghost_rows.append(r)
        if ghost_rows:
            # 역순 삭제로 인덱스 보존
            for r in sorted(ghost_rows, reverse=True):
                ws.delete_rows(r, 1)
            changes['ghost_rows_removed'][sn] = len(ghost_rows)

        # Step 2: freeze_panes
        if ws.freeze_panes != 'A2':
            ws.freeze_panes = 'A2'
            changes['freeze_panes'].append(sn)

        # Step 3: 숫자 서식 '0' — 데이터 행 전체 (정리 후의 행 기준)
        target_cols = NUM_FMT_ZERO[sn]
        for col_idx in target_cols:
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.number_format != '0':
                    cell.number_format = '0'
                    changes['number_format'].append(f'{sn} {get_column_letter(col_idx)}{row_idx}')

        # Step 4: 기본정보 RoB 열(AJ~AT) 배경 재확인 (정리 후 실제 데이터 행만)
        if sn == '기본정보':
            fill = PatternFill(start_color=ROB_FILL_RGB, end_color=ROB_FILL_RGB, fill_type='solid')
            for row_idx in range(2, ws.max_row + 1):
                for col_idx in range(ROB_COL_START, ROB_COL_END + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    existing = cell.fill.fgColor.rgb if cell.fill.fgColor else None
                    if existing != ROB_FILL_RGB:
                        cell.fill = fill
                        changes['rob_fill'] += 1

    return changes


def migrate_file(file_path: Path, dry_run=False, force=False):
    """단일 파일 마이그레이션.
    반환: (status, detail) — status in {'migrated', 'skipped', 'error'}."""
    try:
        wb = openpyxl.load_workbook(file_path)
    except Exception as e:
        return ('error', f'워크북 열기 실패: {e}')

    # 필요 여부 판단
    need, reason = needs_migration(wb, force)
    if not need:
        wb.close()
        return ('skipped', reason)

    # 변경 예정 내용 분석
    changes = apply_v3_formatting(wb)

    summary_parts = []
    if changes.get('ghost_rows_removed'):
        grp = changes['ghost_rows_removed']
        parts = [f'{sn}:{n}' for sn, n in grp.items()]
        summary_parts.append(f'유령행제거({", ".join(parts)})')
    if changes['freeze_panes']:
        summary_parts.append(f'freeze {len(changes["freeze_panes"])}시트')
    if changes['number_format']:
        summary_parts.append(f'숫자서식 {len(changes["number_format"])}셀')
    if changes['rob_fill']:
        summary_parts.append(f'RoB배경 {changes["rob_fill"]}셀')
    summary = ', '.join(summary_parts) if summary_parts else '변경 없음'

    if dry_run:
        wb.close()
        return ('dry-run', f'예정: {summary} ({reason})')

    # 백업
    bak_path = file_path.with_suffix('.bak.xlsx')
    try:
        shutil.copy2(file_path, bak_path)
    except Exception as e:
        wb.close()
        return ('error', f'백업 실패: {e}')

    # 버전 태그 업데이트
    try:
        set_version(wb, SKILL_VERSION)
    except Exception as e:
        log(f'  (버전 태깅 실패, 무시): {e}')

    # 저장
    try:
        wb.save(file_path)
    except Exception as e:
        wb.close()
        # 백업에서 복구
        try:
            shutil.copy2(bak_path, file_path)
        except Exception:
            pass
        return ('error', f'저장 실패: {e}')

    return ('migrated', f'{summary} ({reason}); 백업: {bak_path.name}')


# ============================================================
# 메인
# ============================================================

def _usage():
    log('사용법: python migrate_format.py <file_or_dir> [--dry-run] [--force]')
    sys.exit(1)


def main():
    stdout_utf8()

    if len(sys.argv) < 2:
        _usage()

    target_path = Path(sys.argv[1]).resolve()
    dry_run = '--dry-run' in sys.argv[2:]
    force = '--force' in sys.argv[2:]

    files = collect_target_files(target_path)
    if not files:
        log(f'[ERROR] 대상 파일 없음: {target_path}')
        sys.exit(2)

    log(f'대상 {len(files)}개 파일 (dry_run={dry_run}, force={force})')
    log('')

    counts = {'migrated': 0, 'skipped': 0, 'error': 0, 'dry-run': 0}
    for f in files:
        status, detail = migrate_file(f, dry_run=dry_run, force=force)
        counts[status] = counts.get(status, 0) + 1
        icon = {'migrated': '✓', 'skipped': '-', 'error': '✗', 'dry-run': '→'}.get(status, '?')
        log(f'  {icon} [{status}] {f.name}: {detail}')

    log('')
    log(f'완료 (migrated={counts["migrated"]}, skipped={counts["skipped"]}, error={counts.get("error", 0)}, dry-run={counts.get("dry-run", 0)})')

    sys.exit(0 if counts.get('error', 0) == 0 else 3)


if __name__ == '__main__':
    main()
