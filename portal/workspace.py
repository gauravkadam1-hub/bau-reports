from __future__ import annotations

from pathlib import Path


def save_uploads(uploaded_files, folder: Path) -> dict[str, Path]:
    folder.mkdir(parents=True, exist_ok=True)
    saved: dict[str, Path] = {}
    for uploaded in uploaded_files:
        target = folder / Path(uploaded.name).name
        target.write_bytes(uploaded.getvalue())
        saved[uploaded.name] = target
    return saved
