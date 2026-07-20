"""Shared filesystem helpers: default output location, safe filenames."""
from pathlib import Path


def default_output_dir() -> Path:
    d = Path.home() / "Documents" / "PJS Output"
    d.mkdir(parents=True, exist_ok=True)
    return d


def unique_path(path: str) -> str:
    """Append ' (1)', ' (2)', ... if path already exists, so we never overwrite."""
    p = Path(path)
    if not p.exists():
        return str(p)
    n = 1
    while True:
        candidate = p.with_name(f"{p.stem} ({n}){p.suffix}")
        if not candidate.exists():
            return str(candidate)
        n += 1
