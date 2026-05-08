from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable, List, Optional
from zoneinfo import ZoneInfo

DATA_DIR = Path("data")
PERIODS_PATH = DATA_DIR / "reporting_periods.json"
DEFAULT_TZ = os.environ.get("REPORT_TIMEZONE", "America/Argentina/Buenos_Aires")

SPANISH_MONTH_SHORT = {
    1: "ene",
    2: "feb",
    3: "mar",
    4: "abr",
    5: "may",
    6: "jun",
    7: "jul",
    8: "ago",
    9: "sep",
    10: "oct",
    11: "nov",
    12: "dic",
}

QUARTER_TO_MONTHS = {
    1: [1, 2, 3],
    2: [4, 5, 6],
    3: [7, 8, 9],
    4: [10, 11, 12],
}


@dataclass
class ReportingPeriod:
    kind: str
    year: int
    quarter: Optional[int]
    months: List[str]
    start_date: str
    end_date_exclusive: str
    label: str
    slug: str
    email_subject: str
    title: str
    subtitle: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class ReportingSchedule:
    timezone: str
    reference_date: str
    periods: List[ReportingPeriod]

    def to_dict(self) -> dict:
        return {
            "timezone": self.timezone,
            "reference_date": self.reference_date,
            "periods": [period.to_dict() for period in self.periods],
        }


def _env_value(name: str) -> str | None:
    value = os.environ.get(name)
    if value is None:
        return None

    value = value.strip()
    return value or None


def _parse_reference_date(raw: Optional[str], tz_name: str) -> date:
    if raw:
        return date.fromisoformat(raw)

    return datetime.now(ZoneInfo(tz_name)).date()


def _month_slug(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def _quarter_for_month(month: int) -> int:
    return ((month - 1) // 3) + 1


def _previous_closed_quarter(reference_date: date) -> tuple[int, int]:
    current_quarter = _quarter_for_month(reference_date.month)

    if current_quarter == 1:
        return reference_date.year - 1, 4

    return reference_date.year, current_quarter - 1


def _previous_closed_year(reference_date: date) -> int:
    return reference_date.year - 1


def _quarter_label(year: int, quarter: int) -> str:
    months = QUARTER_TO_MONTHS[quarter]
    month_span = f"{SPANISH_MONTH_SHORT[months[0]]}-{SPANISH_MONTH_SHORT[months[-1]]}"
    return f"Q{quarter} {year} ({month_span})"


def build_quarter_period(year: int, quarter: int) -> ReportingPeriod:
    if quarter not in QUARTER_TO_MONTHS:
        raise ValueError("REPORT_QUARTER debe ser 1, 2, 3 o 4")

    months = QUARTER_TO_MONTHS[quarter]
    month_slugs = [_month_slug(year, month) for month in months]
    start_date = date(year, months[0], 1)

    if quarter == 4:
        end_date_exclusive = date(year + 1, 1, 1)
    else:
        end_date_exclusive = date(year, months[-1] + 1, 1)

    label = _quarter_label(year, quarter)

    return ReportingPeriod(
        kind="quarter",
        year=year,
        quarter=quarter,
        months=month_slugs,
        start_date=start_date.isoformat(),
        end_date_exclusive=end_date_exclusive.isoformat(),
        label=label,
        slug=f"quarter_{year}_Q{quarter}",
        email_subject=f"Informe trimestral CI | Q{quarter} {year}",
        title=f"Informe trimestral de Comunicaciones Internas - Q{quarter} {year}",
        subtitle=f"Período {SPANISH_MONTH_SHORT[months[0]]}-{SPANISH_MONTH_SHORT[months[-1]]} {year}",
    )


def build_year_period(year: int) -> ReportingPeriod:
    month_slugs = [_month_slug(year, month) for month in range(1, 13)]

    return ReportingPeriod(
        kind="year",
        year=year,
        quarter=None,
        months=month_slugs,
        start_date=date(year, 1, 1).isoformat(),
        end_date_exclusive=date(year + 1, 1, 1).isoformat(),
        label=f"Año {year}",
        slug=f"year_{year}",
        email_subject=f"Informe anual CI | {year}",
        title=f"Informe anual de Comunicaciones Internas - {year}",
        subtitle=f"Período ene-dic {year}",
    )


def _append_unique(periods: List[ReportingPeriod], period: ReportingPeriod) -> None:
    if all(existing.slug != period.slug for existing in periods):
        periods.append(period)


def _resolve_quarter_inputs(reference_date: date) -> tuple[int, int]:
    year_raw = _env_value("REPORT_YEAR")
    quarter_raw = _env_value("REPORT_QUARTER")

    if year_raw and quarter_raw:
        year = int(year_raw)
        quarter = int(quarter_raw)

        if quarter not in {1, 2, 3, 4}:
            raise ValueError("REPORT_QUARTER debe ser 1, 2, 3 o 4")

        return year, quarter

    if year_raw or quarter_raw:
        raise ValueError(
            "REPORT_MODE=quarter requiere REPORT_YEAR y REPORT_QUARTER completos, "
            "o ambos vacíos para resolver automáticamente el trimestre cerrado anterior."
        )

    return _previous_closed_quarter(reference_date)


def _resolve_year_input(reference_date: date) -> int:
    year_raw = _env_value("REPORT_YEAR")

    if year_raw:
        return int(year_raw)

    return _previous_closed_year(reference_date)


def resolve_schedule_from_env() -> ReportingSchedule:
    """Resuelve períodos ejecutivos trimestrales/anuales.

    El flujo ya no genera informes mensuales independientes. Los dashboards de
    entrada vienen filtrados desde Looker Studio por trimestre o año, y luego se
    procesan por scope: argentina, holding y combined.
    """
    tz_name = _env_value("REPORT_TIMEZONE") or DEFAULT_TZ
    reference_date = _parse_reference_date(_env_value("REPORT_REFERENCE_DATE"), tz_name)
    report_mode = (_env_value("REPORT_MODE") or "auto").lower()

    periods: List[ReportingPeriod] = []

    previous_month_last_day = reference_date.replace(day=1) - timedelta(days=1)
    previous_month_year = previous_month_last_day.year
    previous_month = previous_month_last_day.month

    if report_mode == "auto":
        if previous_month in {3, 6, 9}:
            _append_unique(
                periods,
                build_quarter_period(
                    previous_month_year,
                    _quarter_for_month(previous_month),
                ),
            )
        elif previous_month == 12:
            _append_unique(periods, build_quarter_period(previous_month_year, 4))
            _append_unique(periods, build_year_period(previous_month_year))

    elif report_mode == "quarter":
        year, quarter = _resolve_quarter_inputs(reference_date)
        periods = [build_quarter_period(year, quarter)]

    elif report_mode == "year":
        year = _resolve_year_input(reference_date)
        periods = [build_year_period(year)]

    elif report_mode == "quarter_and_year":
        year_raw = _env_value("REPORT_YEAR")
        quarter_raw = _env_value("REPORT_QUARTER")

        if year_raw:
            year = int(year_raw)
        else:
            year, _ = _previous_closed_quarter(reference_date)

        if quarter_raw:
            quarter = int(quarter_raw)
        else:
            quarter = 4

        if quarter not in {1, 2, 3, 4}:
            raise ValueError("REPORT_QUARTER debe ser 1, 2, 3 o 4")

        periods = [
            build_quarter_period(year, quarter),
            build_year_period(year),
        ]

    else:
        raise ValueError("REPORT_MODE inválido. Usá: auto, quarter, year o quarter_and_year")

    return ReportingSchedule(
        timezone=tz_name,
        reference_date=reference_date.isoformat(),
        periods=periods,
    )


def save_schedule(schedule: ReportingSchedule) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    PERIODS_PATH.write_text(
        json.dumps(schedule.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_schedule() -> ReportingSchedule:
    if not PERIODS_PATH.exists():
        schedule = resolve_schedule_from_env()
        save_schedule(schedule)
        return schedule

    raw = json.loads(PERIODS_PATH.read_text(encoding="utf-8"))
    periods = [ReportingPeriod(**item) for item in raw.get("periods", [])]

    return ReportingSchedule(
        timezone=raw.get("timezone", DEFAULT_TZ),
        reference_date=raw.get("reference_date", datetime.now().date().isoformat()),
        periods=periods,
    )


def unique_months_from_periods(periods: Iterable[ReportingPeriod]) -> List[str]:
    return sorted({month for period in periods for month in period.months})


def main() -> None:
    schedule = resolve_schedule_from_env()
    save_schedule(schedule)
    print(json.dumps(schedule.to_dict(), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()