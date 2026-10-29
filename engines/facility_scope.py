"""Shared DC/XD source-row scope, matching Flex's existing site eligibility rule.

This module never opens or edits source files. Normalization is used only for
matching; retained field values, column order and row order are left unchanged.
Each engine supplies its own already-resolved facility column. A different site
column must never silently substitute for the field used by report calculations.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import pandas as pd
    import polars as pl


def is_dcxd_site(value: object) -> bool:
    """Match a space + DC/XD suffix, case-insensitively, after space cleanup."""
    if value is None:
        return False
    normalized = re.sub(r"\s+", " ", str(value).strip()).upper()
    return normalized.endswith((" DC", " XD"))


def filter_pandas_dcxd(frame: "pd.DataFrame", facility_column: str) -> "pd.DataFrame":
    """Keep complete matching rows in source order, with a contiguous index."""
    if facility_column not in frame.columns:
        raise ValueError(f"DC/XD filtering requires facility column: {facility_column}")
    site = (frame[facility_column].astype("string").fillna("")
            .str.strip().str.replace(r"\s+", " ", regex=True).str.upper())
    keep = site.str.endswith((" DC", " XD"), na=False)
    return frame.loc[keep].copy().reset_index(drop=True)


def filter_polars_dcxd(frame: "pl.DataFrame", facility_column: str) -> "pl.DataFrame":
    """Use a native Polars expression, not Python per-row callbacks."""
    import polars as pl

    if facility_column not in frame.columns:
        raise ValueError(f"DC/XD filtering requires facility column: {facility_column}")
    site = (pl.col(facility_column).cast(pl.String, strict=False).fill_null("")
            .str.strip_chars().str.replace_all(r"\s+", " ").str.to_uppercase())
    return frame.filter(site.str.ends_with(" DC") | site.str.ends_with(" XD"))
