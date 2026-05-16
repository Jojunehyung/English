#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cleanup_hook.py — PostToolUse hook for convert.py
Reads stdin JSON from Claude Code, checks if a .py file was edited,
then reports dead code and duplicate/inline imports.
"""

import sys
import json
import ast
import os


def get_file_path() -> str:
    try:
        raw = sys.stdin.buffer.read().lstrip(b"\xef\xbb\xbf")  # strip UTF-8 BOM if present
        data = json.loads(raw.decode("utf-8"))
        fp = data.get("tool_input", {}).get("file_path", "")
        if not fp:
            fp = data.get("tool_response", {}).get("filePath", "")
        return fp
    except Exception:
        return ""


def analyze(source: str) -> dict:
    try:
        tree = ast.parse(source)
    except SyntaxError as e:
        return {"error": str(e)}

    # ── Single pass: collect everything ──────────────────────────────────────

    # Module-level function names (private: starts with _)
    module_private_funcs: dict[str, int] = {}
    # Module-level import lines
    module_import_linenos: set[int] = set()
    # Inline imports (inside func/class bodies): (lineno, unparsed)
    inline_imports: list[tuple[int, str]] = []
    # Duplicate function names per scope
    scope_func_names: dict[int, dict[str, list[int]]] = {}
    # All call targets (Names called anywhere in file)
    called_names: set[str] = set()

    # First: collect module-level imports and the top-level package names
    module_top_packages: set[str] = set()
    for node in ast.iter_child_nodes(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            module_import_linenos.add(node.lineno)
            if isinstance(node, ast.Import):
                for alias in node.names:
                    module_top_packages.add(alias.name.split(".")[0])
            else:
                if node.module:
                    module_top_packages.add(node.module.split(".")[0])
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if node.name.startswith("_") and not node.name.startswith("__"):
                module_private_funcs[node.name] = node.lineno

    # Walk entire tree once
    for node in ast.walk(tree):
        # Collect all Name references (calls, assignments, default args, etc.)
        if isinstance(node, ast.Name) and node.id.startswith("_"):
            called_names.add(node.id)
        elif isinstance(node, ast.Call):
            if isinstance(node.func, ast.Attribute):
                called_names.add(node.func.attr)

        # Inline imports: flag only if the package is already at module level
        # (truly redundant inline imports), not optional dependencies
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            scope_id = id(node)
            scope_func_names[scope_id] = {}
            for child in ast.iter_child_nodes(node):
                if isinstance(child, (ast.Import, ast.ImportFrom)):
                    if child.lineno not in module_import_linenos:
                        # Only flag if the package root is already at module level
                        if isinstance(child, ast.Import):
                            pkg = child.names[0].name.split(".")[0]
                        else:
                            pkg = (child.module or "").split(".")[0]
                        if pkg and pkg in module_top_packages:
                            inline_imports.append((child.lineno, ast.unparse(child)))
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    scope_func_names[scope_id].setdefault(child.name, []).append(child.lineno)

    # Also check module-level for duplicate names
    module_scope: dict[str, list[int]] = {}
    for node in ast.iter_child_nodes(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            module_scope.setdefault(node.name, []).append(node.lineno)
    scope_func_names[id(tree)] = module_scope

    # ── Build results ─────────────────────────────────────────────────────────

    dead = [
        (name, lineno)
        for name, lineno in module_private_funcs.items()
        if name not in called_names
    ]

    duplicates = []
    for scope_defs in scope_func_names.values():
        for name, lines in scope_defs.items():
            if len(lines) > 1:
                duplicates.append((name, lines))

    return {
        "dead": dead,
        "duplicates": duplicates,
        "inline_imports": inline_imports,
    }


def report(path: str):
    with open(path, encoding="utf-8") as f:
        source = f.read()

    result = analyze(source)

    if "error" in result:
        print(f"[cleanup_hook] SyntaxError: {result['error']}")
        return

    messages = []
    for name, lineno in result["dead"]:
        messages.append(f"  [dead] '{name}' line {lineno}: never called")
    for name, lines in result["duplicates"]:
        messages.append(f"  [dup] '{name}' lines {lines}")
    for lineno, stmt in result["inline_imports"]:
        messages.append(f"  [inline-import] line {lineno}: {stmt}")

    fname = os.path.basename(path)
    if messages:
        print(f"\n[cleanup_hook] {fname}: {len(messages)} issue(s) found")
        for m in messages:
            print(m)
        summary = "; ".join(messages[:3])
        sys.stdout.flush()
        print(json.dumps({"systemMessage": f"[cleanup_hook] {fname}: {summary}"}, ensure_ascii=False))
    else:
        print(f"[cleanup_hook] {fname}: OK")


def main():
    # Direct argument takes priority; stdin only when no argument given
    if len(sys.argv) > 1:
        fp = sys.argv[1]
    elif not sys.stdin.isatty():
        fp = get_file_path()
    else:
        fp = ""

    if not fp or not fp.endswith(".py"):
        return
    if not os.path.isfile(fp):
        print(f"[cleanup_hook] 파일 없음: {fp}")
        return

    report(fp)


if __name__ == "__main__":
    main()
