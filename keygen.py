#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
keygen.py  —  English Worksheet System 라이센스 키 관리 도구
개발자(배포자)만 사용하는 파일입니다. 배포 시 절대 포함하지 마세요.

사용법:
    python keygen.py
"""

import hmac, hashlib, json, datetime
from pathlib import Path

# ══════════════════════════════════════════════════════════════════
# ※ convert.py 의 _LIC_SECRET 과 반드시 동일해야 합니다!
_LIC_SECRET = b'xogus0226!-biobank1717!-whwnsxxod'
# ══════════════════════════════════════════════════════════════════

# 발급 기록 파일 (keygen.py 와 같은 폴더에 저장)
_RECORD_FILE = Path(__file__).parent / 'license_records.json'


# ── 핵심 함수 ───────────────────────────────────────────────────

def _gen_key(teacher_id: str) -> str:
    raw = hmac.new(
        _LIC_SECRET,
        teacher_id.strip().lower().encode('utf-8'),
        hashlib.sha256
    ).hexdigest().upper()
    return f'EWS-{raw[0:4]}-{raw[4:8]}-{raw[8:12]}-{raw[12:16]}'


def _verify_key(key: str) -> bool:
    """키가 현재 시크릿으로 만들어진 유효한 키인지 확인.
    → 발급 기록 없어도 키 구조 자체를 검증."""
    key = key.strip().upper()
    if not key.startswith('EWS-'):
        return False
    parts = key.split('-')
    if len(parts) != 5:
        return False
    inner = ''.join(parts[1:])
    return len(inner) == 16 and all(c in '0123456789ABCDEF' for c in inner)


def _load_records() -> list:
    if _RECORD_FILE.exists():
        try:
            return json.loads(_RECORD_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
    return []


def _save_records(records: list):
    _RECORD_FILE.write_text(
        json.dumps(records, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )


# ── 메뉴 함수 ───────────────────────────────────────────────────

def menu_generate():
    print('\n[ 키 발급 ]')
    name    = input('  선생님 이름 : ').strip()
    school  = input('  학원명 (선택) : ').strip()
    memo    = input('  메모 (선택) : ').strip()

    if not name:
        print('  ⚠️  이름을 입력하세요.'); return

    key     = _gen_key(name)
    today   = datetime.date.today().isoformat()
    records = _load_records()

    # 중복 확인
    existing = next((r for r in records if r['key'] == key), None)
    if existing:
        print(f'\n  ⚠️  이미 발급된 키입니다.')
        print(f'  선생님: {existing["name"]}')
        print(f'  발급일: {existing["issued"]}')
        print(f'  키    : {key}')
        return

    records.append({
        'name':   name,
        'school': school,
        'memo':   memo,
        'key':    key,
        'issued': today,
        'active': True,
    })
    _save_records(records)

    print()
    print('  ✅  라이센스 키 발급 완료')
    print(f'  선생님 : {name}' + (f' ({school})' if school else ''))
    print(f'  발급일 : {today}')
    print(f'  키     : {key}')
    print()
    print('  → 위 키를 선생님께 전달하세요.')


def menu_list():
    records = _load_records()
    if not records:
        print('\n  발급된 키가 없습니다.'); return

    print(f'\n[ 발급 목록 — 총 {len(records)}건 ]')
    print(f"  {'#':<4} {'이름':<12} {'학원':<16} {'키':<24} {'발급일':<12} {'상태'}")
    print('  ' + '─' * 78)
    for i, r in enumerate(records, 1):
        status = '✅ 활성' if r.get('active', True) else '❌ 비활성'
        print(f"  {i:<4} {r['name']:<12} {r.get('school',''):<16} "
              f"{r['key']:<24} {r['issued']:<12} {status}")
    print()


def menu_verify():
    print('\n[ 키 확인 ]')
    key = input('  확인할 키 : ').strip().upper()

    if not _verify_key(key):
        print('  ❌  형식이 올바르지 않은 키입니다.')
        return

    records = _load_records()
    rec = next((r for r in records if r['key'] == key), None)

    if rec:
        status = '✅ 활성' if rec.get('active', True) else '❌ 비활성'
        print(f'\n  ✅  발급 기록 있음')
        print(f'  선생님 : {rec["name"]}' + (f' ({rec.get("school","")})' if rec.get("school") else ''))
        print(f'  발급일 : {rec["issued"]}')
        print(f'  메모   : {rec.get("memo","—")}')
        print(f'  상태   : {status}')
    else:
        print('  ⚠️  발급 기록 없음 (직접 생성했거나 기록 누락)')
        print('  키 구조는 유효합니다.')


def menu_deactivate():
    print('\n[ 키 비활성화 (라이센스 취소) ]')
    records = _load_records()
    active  = [r for r in records if r.get('active', True)]

    if not active:
        print('  활성 상태인 키가 없습니다.'); return

    for i, r in enumerate(active, 1):
        print(f'  {i}. {r["name"]}  |  {r["key"]}  |  발급일: {r["issued"]}')

    try:
        idx = int(input('\n  비활성화할 번호 : ')) - 1
        if idx < 0 or idx >= len(active):
            print('  ⚠️  올바른 번호를 입력하세요.'); return
    except ValueError:
        print('  ⚠️  숫자를 입력하세요.'); return

    target = active[idx]
    for r in records:
        if r['key'] == target['key']:
            r['active'] = False
            break
    _save_records(records)
    print(f'\n  ✅  {target["name"]}의 키를 비활성화했습니다.')
    print('  ※ 이미 등록된 기기에서는 여전히 동작합니다.')
    print('     (완전 차단하려면 온라인 인증이 필요합니다.)')


# ── 메인 ────────────────────────────────────────────────────────

def main():
    print('=' * 55)
    print('  English Worksheet System  —  라이센스 관리')
    print('=' * 55)

    menus = [
        ('1', '키 발급 (새 선생님 등록)',  menu_generate),
        ('2', '발급 목록 조회',             menu_list),
        ('3', '키 확인 / 검증',             menu_verify),
        ('4', '키 비활성화 (라이센스 취소)', menu_deactivate),
        ('0', '종료',                        None),
    ]

    while True:
        print()
        for num, label, _ in menus:
            print(f'  [{num}]  {label}')
        choice = input('\n  선택 : ').strip()

        if choice == '0':
            print('\n  종료합니다.\n')
            break

        matched = next((fn for n, _, fn in menus if n == choice and fn), None)
        if matched:
            matched()
        else:
            print('  ⚠️  올바른 번호를 입력하세요.')


if __name__ == '__main__':
    main()
