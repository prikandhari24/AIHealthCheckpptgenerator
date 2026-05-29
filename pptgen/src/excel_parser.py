from __future__ import annotations

from dataclasses import dataclass
import re

import pandas as pd


LEVEL_COLUMN_SUFFIX = " assessed level"
DEFAULT_SHEET_NAME = "AI Maturity Matrix"
LABEL_ALIASES = {
    "ai individual productivity tools": "individual productivity ai tools",
}

LEVEL_ALIASES = {
    "lagging": "behind",
    "behind": "behind",
    "at par": "at_par",
    "atpar": "at_par",
    "on par": "at_par",
    "par": "at_par",
    "leading": "ahead",
    "ahead": "ahead",
}


def _strip_column_suffix(column_name: str, suffix: str) -> str | None:
    pattern = re.compile(rf"{re.escape(suffix)}$", re.IGNORECASE)
    if not pattern.search(column_name):
        return None
    return pattern.sub("", column_name).strip()


def normalize_label(value: str) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("&", " and ")
    text = re.sub(r"^\d+\.\s*", "", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    normalized = " ".join(text.split())
    return LABEL_ALIASES.get(normalized, normalized)


def normalize_level(value: str) -> str:
    normalized = normalize_label(value).replace(" ", "")
    if normalized in LEVEL_ALIASES:
        return LEVEL_ALIASES[normalized]

    normalized = normalize_label(value)
    if normalized in LEVEL_ALIASES:
        return LEVEL_ALIASES[normalized]

    raise ValueError(f"Unsupported assessed level: {value}")


@dataclass(frozen=True)
class IndicatorAssessment:
    dimension: str
    indicator: str
    company_levels: dict[str, str]


@dataclass
class ParsedExcel:
    raw_df: pd.DataFrame
    companies: list[str]
    level_columns: dict[str, str]
    summary_columns: dict[str, str]
    rows: list[IndicatorAssessment]

    def __post_init__(self):
        self._indicator_lookup = {
            normalize_label(row.indicator): row
            for row in self.rows
        }

    @property
    def preview_df(self) -> pd.DataFrame:
        ordered_columns = ["Dimension", "Indicator", *self.level_columns.values()]
        return self.raw_df[ordered_columns].copy()

    def get_indicator_assessment(self, indicator: str) -> IndicatorAssessment | None:
        return self._indicator_lookup.get(normalize_label(indicator))


def _get_sheet_name(file) -> str | int:
    if hasattr(file, "seek"):
        file.seek(0)
    workbook = pd.ExcelFile(file)
    if DEFAULT_SHEET_NAME in workbook.sheet_names:
        return DEFAULT_SHEET_NAME
    return workbook.sheet_names[0]


def parse_excel(file) -> ParsedExcel:
    sheet_name = _get_sheet_name(file)
    if hasattr(file, "seek"):
        file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name)

    required_columns = ["Dimension", "Indicator"]
    missing = [column for column in required_columns if column not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in Excel: {missing}")

    level_columns = {}
    for column in df.columns:
        company = _strip_column_suffix(str(column), LEVEL_COLUMN_SUFFIX)
        if company is not None:
            level_columns[company] = column

    summary_candidates: dict[str, str] = {}
    for column in df.columns:
        company = _strip_column_suffix(str(column), " summarized assessment")
        if company is not None:
            summary_candidates[normalize_label(company)] = column

    summary_columns = {
        company: summary_candidates[normalize_label(company)]
        for company in level_columns
        if normalize_label(company) in summary_candidates
    }
    if not level_columns:
        raise ValueError(
            "No assessed level columns found. Expected columns ending with "
            f"'{LEVEL_COLUMN_SUFFIX}'."
        )

    working_df = df.dropna(subset=["Indicator"]).copy()
    rows: list[IndicatorAssessment] = []

    for _, row in working_df.iterrows():
        company_levels = {
            company: normalize_level(row[column])
            for company, column in level_columns.items()
            if pd.notna(row[column])
        }

        rows.append(
            IndicatorAssessment(
                dimension=str(row["Dimension"]).strip(),
                indicator=str(row["Indicator"]).strip(),
                company_levels=company_levels,
            )
        )

    return ParsedExcel(
        raw_df=working_df,
        companies=list(level_columns.keys()),
        level_columns=level_columns,
        summary_columns=summary_columns,
        rows=rows,
    )
