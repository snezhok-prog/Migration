from __future__ import annotations

import glob
import os
from pathlib import Path
from typing import List

from _utils import parse_path_list


def _resolve_candidate(script_dir: str, token: str) -> str:
    p = Path(str(token or "").strip())
    if p.is_absolute():
        return str(p.resolve())

    candidates = [
        (Path.cwd() / p).resolve(),
        (Path(script_dir) / p).resolve(),
        (Path(script_dir).parent / p).resolve(),
    ]
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return str(candidate)
    return str(candidates[0])


def discover_excel_files(script_dir: str, explicit_files: str = "", pattern: str = "*.xlsm") -> List[str]:
    explicit = parse_path_list(explicit_files)
    if explicit:
        out = []
        for raw in explicit:
            candidate = _resolve_candidate(script_dir, raw)
            if os.path.isfile(candidate) and not os.path.basename(candidate).startswith("~$"):
                out.append(os.path.abspath(candidate))
        return sorted(list(dict.fromkeys(out)), key=lambda p: os.path.basename(p).lower())

    files = [
        os.path.abspath(p)
        for p in glob.glob(os.path.join(script_dir, pattern))
        if os.path.isfile(p) and not os.path.basename(p).startswith("~$")
    ]
    files.sort(key=lambda p: os.path.basename(p).lower())
    return files
