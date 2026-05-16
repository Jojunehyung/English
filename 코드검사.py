#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
코드검사.py — Python 파일 죽은코드 / 중복 / 인라인 import 검사 도구
더블클릭 또는 터미널에서 직접 실행 가능
"""

import ast
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog
from pathlib import Path

# ── 기본 대상 파일 ───────────────────────────────────────────────
DEFAULT_TARGET = Path(__file__).parent / 'convert.py'

# ── 색상 ────────────────────────────────────────────────────────
BG    = '#f0f2f7'
CARD  = '#ffffff'
ITEM  = '#f7f8fc'
BDR   = '#d0d5e8'
GOLD  = '#b8932a'
GOLDB = '#d4a730'
TEXT  = '#1a1d2e'
SUB   = '#5c6585'
DIM   = '#a8b0cc'
OK    = '#16803c'
WARN  = '#b45309'
ERR   = '#dc2626'
SEL   = '#dce6ff'
KO    = '맑은 고딕'
MONO  = 'Consolas'


# ════════════════════════════════════════════════════════════════
# 분석 엔진
# ════════════════════════════════════════════════════════════════

def analyze(source: str) -> dict:
    try:
        tree = ast.parse(source)
    except SyntaxError as e:
        return {'error': str(e), 'dead': [], 'duplicates': [], 'inline': []}

    # ── 모듈 레벨 정보 수집 ──────────────────────────────────────
    module_import_linenos: set[int] = set()
    module_top_packages:   set[str] = set()
    module_private_funcs:  dict[str, int] = {}

    for node in ast.iter_child_nodes(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            module_import_linenos.add(node.lineno)
            if isinstance(node, ast.Import):
                for alias in node.names:
                    module_top_packages.add(alias.name.split('.')[0])
            elif node.module:
                module_top_packages.add(node.module.split('.')[0])
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            name = node.name
            if name.startswith('_') and not name.startswith('__'):
                module_private_funcs[name] = node.lineno

    # ── 단일 순회: 이름 참조 / 인라인 import / 중복 함수 ────────
    referenced_names: set[str] = set()
    inline_imports:   list[tuple[int, str]] = []
    scope_func_names: dict[int, dict[str, list[int]]] = {}

    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and node.id.startswith('_'):
            referenced_names.add(node.id)
        elif isinstance(node, ast.Call):
            if isinstance(node.func, ast.Attribute):
                referenced_names.add(node.func.attr)

        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            sid = id(node)
            scope_func_names[sid] = {}
            for child in ast.iter_child_nodes(node):
                if isinstance(child, (ast.Import, ast.ImportFrom)):
                    if child.lineno not in module_import_linenos:
                        pkg = ''
                        if isinstance(child, ast.Import):
                            pkg = child.names[0].name.split('.')[0]
                        elif child.module:
                            pkg = child.module.split('.')[0]
                        if pkg and pkg in module_top_packages:
                            inline_imports.append((child.lineno, ast.unparse(child)))
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    scope_func_names[sid].setdefault(child.name, []).append(child.lineno)

    # 모듈 레벨 중복 함수
    module_scope: dict[str, list[int]] = {}
    for node in ast.iter_child_nodes(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            module_scope.setdefault(node.name, []).append(node.lineno)
    scope_func_names[id(tree)] = module_scope

    # ── 결과 정리 ────────────────────────────────────────────────
    dead = [
        (name, lineno)
        for name, lineno in module_private_funcs.items()
        if name not in referenced_names
    ]
    dead.sort(key=lambda x: x[1])

    duplicates = []
    for scope_defs in scope_func_names.values():
        for name, lines in scope_defs.items():
            if len(lines) > 1:
                duplicates.append((name, lines))
    duplicates.sort(key=lambda x: x[1][0])

    inline_imports.sort(key=lambda x: x[0])

    return {
        'error':      None,
        'dead':       dead,
        'duplicates': duplicates,
        'inline':     inline_imports,
    }


# ════════════════════════════════════════════════════════════════
# GUI
# ════════════════════════════════════════════════════════════════

class CheckerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('코드 검사 도구')
        self.geometry('820x640')
        self.minsize(640, 480)
        self.configure(bg=BG)

        self._current_file: Path | None = None
        self._setup_style()
        self._build_ui()

        # 기본 파일 자동 로드
        if DEFAULT_TARGET.exists():
            self._load_file(DEFAULT_TARGET)

    # ── TTK 스타일 ──────────────────────────────────────────────

    def _setup_style(self):
        s = ttk.Style(self)
        s.theme_use('clam')
        s.configure('Vertical.TScrollbar',
                    background=BDR, troughcolor=ITEM,
                    arrowcolor=SUB, relief='flat')
        s.configure('Issue.Treeview',
                    font=(KO, 9), rowheight=24,
                    background=CARD, fieldbackground=CARD,
                    foreground=TEXT, borderwidth=0)
        s.configure('Issue.Treeview.Heading',
                    font=(KO, 9, 'bold'),
                    background=ITEM, foreground=TEXT)
        s.map('Issue.Treeview', background=[('selected', SEL)])

    # ── UI 구성 ──────────────────────────────────────────────────

    def _build_ui(self):
        # ── 헤더 ──
        hdr = tk.Frame(self, bg='#1a1d2e')
        hdr.pack(fill='x')
        ih = tk.Frame(hdr, bg='#1a1d2e')
        ih.pack(fill='x', padx=20, pady=14)
        tk.Label(ih, text='🔍  코드 검사 도구',
                 font=(KO, 14, 'bold'), bg='#1a1d2e', fg='#ffffff').pack(side='left')
        tk.Label(ih, text='죽은코드  ·  중복함수  ·  인라인 import',
                 font=(KO, 8), bg='#1a1d2e', fg=DIM).pack(side='left', padx=(12, 0), pady=(2, 0))

        # ── 파일 선택 바 ──
        fb = tk.Frame(self, bg=BDR)
        fb.pack(fill='x')
        fi = tk.Frame(fb, bg=CARD)
        fi.pack(fill='x', padx=1, pady=1)
        frow = tk.Frame(fi, bg=CARD)
        frow.pack(fill='x', padx=16, pady=10)

        tk.Label(frow, text='대상 파일', font=(KO, 8), bg=CARD, fg=SUB,
                 width=7, anchor='w').pack(side='left')
        self._file_lbl = tk.Label(frow, text='파일을 선택하세요',
                                  font=(MONO, 9), bg=ITEM, fg=TEXT,
                                  anchor='w', padx=8, pady=4, relief='flat')
        self._file_lbl.pack(side='left', fill='x', expand=True, padx=(0, 8))
        tk.Button(frow, text='📂  파일 선택',
                  font=(KO, 9), bg=GOLD, fg='#ffffff',
                  relief='flat', cursor='hand2', bd=0,
                  padx=12, pady=5,
                  command=self._pick_file).pack(side='left', padx=(0, 6))
        tk.Button(frow, text='🔄  다시 검사',
                  font=(KO, 9), bg=ITEM, fg=TEXT,
                  relief='flat', cursor='hand2', bd=0,
                  padx=12, pady=5,
                  command=self._rescan).pack(side='left')

        # ── 요약 배지 행 ──
        self._badge_row = tk.Frame(self, bg=BG)
        self._badge_row.pack(fill='x', padx=16, pady=(12, 0))

        # ── 탭 영역 ──
        nb_frame = tk.Frame(self, bg=BG)
        nb_frame.pack(fill='both', expand=True, padx=16, pady=12)

        self._nb = ttk.Notebook(nb_frame)
        self._nb.pack(fill='both', expand=True)

        self._tab_dead  = self._make_tab('💀  죽은코드',   ['함수명', '정의 라인', '설명'])
        self._tab_dup   = self._make_tab('♻️  중복함수',   ['함수명', '라인 목록', '설명'])
        self._tab_inline= self._make_tab('📦  인라인 import', ['라인', 'import 구문', '비고'])
        self._tab_all   = self._make_tab('📋  전체 목록',  ['종류', '이름/구문', '라인', '내용'])

        # ── 상태바 ──
        sb = tk.Frame(self, bg='#1a1d2e', height=30)
        sb.pack(fill='x', side='bottom')
        self._status_lbl = tk.Label(sb, text='파일을 선택하세요',
                                    font=(KO, 8), bg='#1a1d2e', fg=DIM)
        self._status_lbl.pack(side='left', padx=16, pady=6)
        self._lines_lbl = tk.Label(sb, text='',
                                   font=(MONO, 8), bg='#1a1d2e', fg=DIM)
        self._lines_lbl.pack(side='right', padx=16, pady=6)

    def _make_tab(self, title, columns):
        frame = tk.Frame(self._nb, bg=BG)
        self._nb.add(frame, text=f'  {title}  ')

        wrap = tk.Frame(frame, bg=BDR, bd=1, relief='flat')
        wrap.pack(fill='both', expand=True, pady=(8, 0))

        tv = ttk.Treeview(wrap, columns=columns, show='headings',
                          style='Issue.Treeview')
        sb = ttk.Scrollbar(wrap, orient='vertical', command=tv.yview,
                           style='Vertical.TScrollbar')
        tv.configure(yscrollcommand=sb.set)

        col_widths = {
            '함수명': 160, '정의 라인': 80, '설명': 320,
            '라인 목록': 160, 'import 구문': 340, '비고': 120,
            '라인': 60, '종류': 80, '이름/구문': 260, '내용': 200,
        }
        for col in columns:
            w = col_widths.get(col, 120)
            tv.heading(col, text=col)
            tv.column(col, width=w, anchor='w', minwidth=50)

        sb.pack(side='right', fill='y')
        tv.pack(fill='both', expand=True, padx=1, pady=1)

        # 태그 색상
        tv.tag_configure('dead',   foreground=ERR)
        tv.tag_configure('dup',    foreground=WARN)
        tv.tag_configure('inline', foreground='#1a4a8a')
        tv.tag_configure('ok',     foreground=OK)

        return tv

    # ── 파일 처리 ────────────────────────────────────────────────

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title='검사할 Python 파일 선택',
            filetypes=[('Python 파일', '*.py'), ('모든 파일', '*.*')],
            initialdir=str(Path(__file__).parent))
        if path:
            self._load_file(Path(path))

    def _rescan(self):
        if self._current_file:
            self._load_file(self._current_file)

    def _load_file(self, path: Path):
        self._current_file = path
        self._file_lbl.configure(text=str(path))
        self._status_lbl.configure(text='검사 중...', fg=WARN)
        self.update_idletasks()

        try:
            source = path.read_text(encoding='utf-8')
        except Exception as e:
            self._status_lbl.configure(text=f'❌  읽기 실패: {e}', fg=ERR)
            return

        lines = source.splitlines()
        result = analyze(source)
        self._render(result, len(lines), path.name)

    # ── 결과 렌더링 ──────────────────────────────────────────────

    def _render(self, result: dict, total_lines: int, filename: str):
        # 기존 내용 지우기
        for tv in (self._tab_dead, self._tab_dup, self._tab_inline, self._tab_all):
            tv.delete(*tv.get_children())

        # 배지 행 초기화
        for w in self._badge_row.winfo_children():
            w.destroy()

        if result.get('error'):
            self._status_lbl.configure(
                text=f'❌  파싱 오류: {result["error"]}', fg=ERR)
            self._badge(self._badge_row, '❌ 파싱 오류', result['error'][:60], ERR)
            return

        dead    = result['dead']
        dups    = result['duplicates']
        inlines = result['inline']
        total   = len(dead) + len(dups) + len(inlines)

        # ── 배지 ──
        def badge(icon, label, count, color):
            f = tk.Frame(self._badge_row, bg=color if count else ITEM,
                         padx=10, pady=6)
            f.pack(side='left', padx=(0, 8))
            fg = '#ffffff' if count else DIM
            tk.Label(f, text=f'{icon}  {label}',
                     font=(KO, 8, 'bold'), bg=f['bg'], fg=fg).pack(side='left')
            tk.Label(f, text=f'  {count}건',
                     font=(MONO, 10, 'bold'), bg=f['bg'], fg=fg).pack(side='left')

        badge('💀', '죽은코드',     len(dead),    ERR  if dead    else ITEM)
        badge('♻️', '중복함수',     len(dups),    WARN if dups    else ITEM)
        badge('📦', '인라인 import', len(inlines), '#1a4a8a' if inlines else ITEM)

        result_f = tk.Frame(self._badge_row, bg=OK if total == 0 else ITEM, padx=10, pady=6)
        result_f.pack(side='right')
        res_fg = '#ffffff' if total == 0 else DIM
        res_txt = '✅  이슈 없음' if total == 0 else f'⚠️  총 {total}건'
        tk.Label(result_f, text=res_txt, font=(KO, 9, 'bold'),
                 bg=result_f['bg'], fg=res_fg).pack()

        # ── 죽은코드 탭 ──
        for name, lineno in dead:
            self._tab_dead.insert('', 'end', tags=('dead',),
                values=(name, f'{lineno}번', '호출되지 않는 모듈 레벨 비공개 함수'))

        if not dead:
            self._tab_dead.insert('', 'end', tags=('ok',),
                values=('—', '—', '죽은코드 없음 ✓'))

        # ── 중복함수 탭 ──
        for name, lines in dups:
            self._tab_dup.insert('', 'end', tags=('dup',),
                values=(name, ', '.join(f'{l}번' for l in lines),
                        '같은 스코프 내 동일 이름 함수 정의'))

        if not dups:
            self._tab_dup.insert('', 'end', tags=('ok',),
                values=('—', '—', '중복 없음 ✓'))

        # ── 인라인 import 탭 ──
        for lineno, stmt in inlines:
            self._tab_inline.insert('', 'end', tags=('inline',),
                values=(f'{lineno}번', stmt,
                        '모듈 레벨에 이미 있는 패키지 — 상단으로 이동 권장'))

        if not inlines:
            self._tab_inline.insert('', 'end', tags=('ok',),
                values=('—', '—', '인라인 import 없음 ✓'))

        # ── 전체 목록 탭 ──
        for name, lineno in dead:
            self._tab_all.insert('', 'end', tags=('dead',),
                values=('💀 죽은코드', name, f'{lineno}번', '호출 없음'))
        for name, lines in dups:
            self._tab_all.insert('', 'end', tags=('dup',),
                values=('♻️ 중복함수', name,
                        ', '.join(f'{l}번' for l in lines), '동일 이름 중복'))
        for lineno, stmt in inlines:
            self._tab_all.insert('', 'end', tags=('inline',),
                values=('📦 인라인 import', stmt, f'{lineno}번', '상단 이동 권장'))

        if total == 0:
            self._tab_all.insert('', 'end', tags=('ok',),
                values=('✅ 이상 없음', '—', '—', '모든 검사 통과'))

        # ── 상태바 ──
        if total == 0:
            self._status_lbl.configure(text='✅  이슈 없음 — 깔끔합니다', fg=OK)
        else:
            self._status_lbl.configure(
                text=f'⚠️  {total}건 발견  (죽은코드 {len(dead)} / 중복 {len(dups)} / 인라인import {len(inlines)})',
                fg=WARN)
        self._lines_lbl.configure(text=f'{filename}  |  {total_lines:,}줄')

    def _badge(self, parent, label, text, color):
        f = tk.Frame(parent, bg=color, padx=10, pady=6)
        f.pack(side='left', padx=(0, 8))
        tk.Label(f, text=f'{label}: {text}',
                 font=(KO, 8), bg=color, fg='#ffffff').pack()


# ════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    app = CheckerApp()
    app.mainloop()
