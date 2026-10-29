"""Lossless XLSX ZIP compression. No workbook XML is modified or reserialized."""
from __future__ import annotations

from pathlib import Path
from shutil import copyfileobj
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZIP64_LIMIT, ZipFile


def compact_workbook(path: Path) -> dict[str, int]:
    """Recompress with standard DEFLATE, retaining the original if not smaller.

    Streams parts in 1 MiB chunks. All part names, order and uncompressed bytes
    are preserved. An exception never replaces the source. XLSX ZIP-entry
    timestamps/attributes are not workbook data and are newly generated.
    The caller validates the result before publishing its final report.
    """
    path = Path(path)
    original_size = path.stat().st_size
    candidate = None
    try:
        with NamedTemporaryFile(prefix=".onboarding-compact-", suffix=".xlsx",
                                dir=path.parent, delete=False) as handle:
            candidate = Path(handle.name)
        with ZipFile(path, "r") as source, ZipFile(
                candidate, "w", compression=ZIP_DEFLATED, compresslevel=9,
                allowZip64=True) as target:
            infos = source.infolist()
            if len(infos) != len({info.filename for info in infos}):
                raise ValueError("Cannot compact a workbook with duplicate ZIP parts.")
            target.comment = source.comment
            for info in infos:
                with source.open(info, "r") as reader, target.open(
                        info.filename, "w", force_zip64=info.file_size >= ZIP64_LIMIT) as writer:
                    copyfileobj(reader, writer, length=1024 * 1024)
                # CRC and size are independently computed while writing.
                copied = target.getinfo(info.filename)
                if (copied.CRC, copied.file_size) != (info.CRC, info.file_size):
                    raise ValueError(f"Workbook compression integrity check failed for {info.filename}.")
        compact_size = candidate.stat().st_size
        if compact_size < original_size:
            candidate.replace(path)
        else:
            compact_size = original_size
        return {"before_bytes": original_size, "after_bytes": compact_size,
                "saved_bytes": original_size - compact_size}
    finally:
        if candidate is not None and candidate.exists():
            candidate.unlink()
