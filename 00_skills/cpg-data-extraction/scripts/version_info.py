#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
version_info.py — v3.0 (cpg-data-extraction 버전 태깅 모듈)

엑셀 파일에 스킬 버전을 메타데이터로 기록·조회한다.
save_extract.py 및 migrate_format.py가 import하여 사용.

기록 방식 (우선순위):
  1) Custom Document Property (키 'cpg_skill_version')  — openpyxl 3.1+ 지원
  2) 표준 properties.keywords에 'cpg-skill=<버전>' 태그 추가  — fallback

CLI 사용:
    python version_info.py read <xlsx>                 # 버전 읽기
    python version_info.py set <xlsx> [version]        # 버전 설정 (인자 없으면 'v3.0')

라이브러리 사용:
    from version_info import set_version, get_version, SKILL_VERSION
    set_version(wb)                 # wb는 openpyxl Workbook
    ver = get_version(wb)           # 'v3.0' or None
"""

import sys
from pathlib import Path

import openpyxl

SKILL_VERSION = 'v3.0'
SKILL_NAME    = 'cpg-data-extraction'
PROP_KEY      = 'cpg_skill_version'   # custom property key

# Fallback keyword prefix (표준 properties.keywords에 삽입 시 사용)
KEYWORD_TAG_PREFIX = 'cpg-skill='


# ============================================================
# Custom Property 접근 (openpyxl 3.1+)
# ============================================================

def _custom_props_supported(wb):
    """wb.custom_doc_props 지원 여부 감지."""
    try:
        _ = wb.custom_doc_props
        return True
    except AttributeError:
        return False


def _set_via_custom_prop(wb, version):
    """Custom property로 버전 기록. 기존 동일 키 있으면 교체."""
    from openpyxl.packaging.custom import CustomDocumentProperty
    cdpl = wb.custom_doc_props
    # 기존 제거 (중복 방지)
    existing = [p for p in list(cdpl.props) if p.name == PROP_KEY]
    for p in existing:
        try:
            cdpl.props.remove(p)
        except Exception:
            pass
    # 새로 추가
    new_prop = CustomDocumentProperty(name=PROP_KEY, value=version)
    cdpl.props.append(new_prop)


def _get_via_custom_prop(wb):
    """Custom property에서 버전 읽기. 없으면 None."""
    try:
        cdpl = wb.custom_doc_props
        for p in cdpl.props:
            if p.name == PROP_KEY:
                return str(p.value)
    except Exception:
        pass
    return None


# ============================================================
# Keyword Fallback (openpyxl 구버전)
# ============================================================

def _set_via_keywords(wb, version):
    kws = wb.properties.keywords or ''
    # 기존 cpg-skill= 태그 제거
    tokens = [t.strip() for t in kws.split(',') if t.strip()]
    tokens = [t for t in tokens if not t.startswith(KEYWORD_TAG_PREFIX)]
    tokens.append(f'{KEYWORD_TAG_PREFIX}{version}')
    wb.properties.keywords = ', '.join(tokens)


def _get_via_keywords(wb):
    kws = wb.properties.keywords or ''
    for tok in kws.split(','):
        tok = tok.strip()
        if tok.startswith(KEYWORD_TAG_PREFIX):
            return tok[len(KEYWORD_TAG_PREFIX):]
    return None


# ============================================================
# 공개 API
# ============================================================

def set_version(wb, version=SKILL_VERSION):
    """Workbook에 스킬 버전 기록. Custom property 우선, 실패 시 keywords fallback."""
    try:
        if _custom_props_supported(wb):
            _set_via_custom_prop(wb, version)
            return 'custom_prop'
    except Exception:
        pass
    _set_via_keywords(wb, version)
    return 'keywords'


def get_version(wb):
    """Workbook에서 스킬 버전 읽기. 없으면 None."""
    v = _get_via_custom_prop(wb)
    if v:
        return v
    return _get_via_keywords(wb)


# ============================================================
# CLI
# ============================================================

def _usage():
    print('사용법:')
    print('  python version_info.py read <xlsx>')
    print('  python version_info.py set <xlsx> [version]   # 기본 v3.0')
    sys.exit(1)


def _main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    if len(sys.argv) < 3:
        _usage()

    cmd = sys.argv[1]
    path = Path(sys.argv[2]).resolve()
    if not path.exists():
        print(f'[ERROR] 파일 없음: {path}')
        sys.exit(1)

    if cmd == 'read':
        wb = openpyxl.load_workbook(path, read_only=False)
        v = get_version(wb)
        wb.close()
        if v:
            print(v)
            sys.exit(0)
        else:
            print('(버전 정보 없음)')
            sys.exit(2)

    elif cmd == 'set':
        version = sys.argv[3] if len(sys.argv) >= 4 else SKILL_VERSION
        wb = openpyxl.load_workbook(path)
        how = set_version(wb, version)
        wb.save(path)
        print(f'✓ {path.name}에 version={version} 기록 (방식: {how})')
        sys.exit(0)

    else:
        _usage()


if __name__ == '__main__':
    _main()
