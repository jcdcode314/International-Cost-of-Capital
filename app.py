"""International Cost of Capital Calculator; Python 3.12.

Run: streamlit run app.py
Verify offline: python app.py --self-test
All internal rates are decimals. Source quotes and UI percentage inputs are /100.
"""

from __future__ import annotations

import csv
import io
import json
import math
import os
import re
import struct
import sys
import time
import unicodedata
import unittest
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, replace
from datetime import datetime, timedelta, timezone
from typing import Any
from unittest.mock import patch

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import xlrd
from bs4 import BeautifulSoup


# ---------------------------------------------------------------------------
# Configuration and country/currency identities (not financial observations).
# ---------------------------------------------------------------------------
BETA_XLS = "https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls"
BETA_HTML = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/Betas.html"
RATINGS_URL = "https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/ratings.html"
OECD_PAGE = "https://www.oecd.org/en/data/datasets/corporate-income-tax-rates-database.html"
OECD_FLOW = "OECD.CTP.TPS,DSD_TAX_CIT@DF_CIT,2.0"
OECD_API = f"https://sdmx.oecd.org/public/rest/v1/data/{OECD_FLOW}/.A.CIT_C.ST..S13..."
BONDS_URL = "https://tradingeconomics.com/bonds"
DEFAULT_INDUSTRY = "Software (System & Application)"
USER_AGENT = "InternationalCostOfCapital/1.0 (educational public-data reader)"
APP_REVISION = "2026-11-16.3"

# ISO 3166-1 alpha-3 / ISO 4217. A missing public quote never removes a country.
# Bulgaria adopted EUR in 2026. Euro-area yields remain country-specific.
COUNTRY_TEXT = """
Albania|ALB|ALL|albania
Algeria|DZA|DZD|algeria
Angola|AGO|AOA|angola
Argentina|ARG|ARS|argentina
Armenia|ARM|AMD|armenia
Australia|AUS|AUD|australia
Austria|AUT|EUR|austria
Azerbaijan|AZE|AZN|azerbaijan
Bahamas|BHS|BSD|bahamas
Bahrain|BHR|BHD|bahrain
Bangladesh|BGD|BDT|bangladesh
Barbados|BRB|BBD|barbados
Belgium|BEL|EUR|belgium
Belize|BLZ|BZD|belize
Benin|BEN|XOF|benin
Bermuda|BMU|BMD|bermuda
Bolivia|BOL|BOB|bolivia
Bosnia and Herzegovina|BIH|BAM|bosnia-and-herzegovina
Botswana|BWA|BWP|botswana
Brazil|BRA|BRL|brazil
Bulgaria|BGR|EUR|bulgaria
Cambodia|KHM|KHR|cambodia
Cameroon|CMR|XAF|cameroon
Canada|CAN|CAD|canada
Cayman Islands|CYM|KYD|cayman-islands
Chile|CHL|CLP|chile
China|CHN|CNY|china
Colombia|COL|COP|colombia
Costa Rica|CRI|CRC|costa-rica
Croatia|HRV|EUR|croatia
Cyprus|CYP|EUR|cyprus
Czechia|CZE|CZK|czech-republic
Denmark|DNK|DKK|denmark
Dominican Republic|DOM|DOP|dominican-republic
Ecuador|ECU|USD|ecuador
Egypt|EGY|EGP|egypt
El Salvador|SLV|USD|el-salvador
Estonia|EST|EUR|estonia
Ethiopia|ETH|ETB|ethiopia
Finland|FIN|EUR|finland
France|FRA|EUR|france
Georgia|GEO|GEL|georgia
Germany|DEU|EUR|germany
Ghana|GHA|GHS|ghana
Greece|GRC|EUR|greece
Guatemala|GTM|GTQ|guatemala
Honduras|HND|HNL|honduras
Hong Kong|HKG|HKD|hong-kong
Hungary|HUN|HUF|hungary
Iceland|ISL|ISK|iceland
India|IND|INR|india
Indonesia|IDN|IDR|indonesia
Ireland|IRL|EUR|ireland
Israel|ISR|ILS|israel
Italy|ITA|EUR|italy
Jamaica|JAM|JMD|jamaica
Japan|JPN|JPY|japan
Jordan|JOR|JOD|jordan
Kazakhstan|KAZ|KZT|kazakhstan
Kenya|KEN|KES|kenya
Kuwait|KWT|KWD|kuwait
Latvia|LVA|EUR|latvia
Lebanon|LBN|LBP|lebanon
Lithuania|LTU|EUR|lithuania
Luxembourg|LUX|EUR|luxembourg
Malaysia|MYS|MYR|malaysia
Malta|MLT|EUR|malta
Mauritius|MUS|MUR|mauritius
Mexico|MEX|MXN|mexico
Moldova|MDA|MDL|moldova
Mongolia|MNG|MNT|mongolia
Montenegro|MNE|EUR|montenegro
Morocco|MAR|MAD|morocco
Namibia|NAM|NAD|namibia
Nepal|NPL|NPR|nepal
Netherlands|NLD|EUR|netherlands
New Zealand|NZL|NZD|new-zealand
Nigeria|NGA|NGN|nigeria
North Macedonia|MKD|MKD|macedonia
Norway|NOR|NOK|norway
Oman|OMN|OMR|oman
Pakistan|PAK|PKR|pakistan
Panama|PAN|USD|panama
Paraguay|PRY|PYG|paraguay
Peru|PER|PEN|peru
Philippines|PHL|PHP|philippines
Poland|POL|PLN|poland
Portugal|PRT|EUR|portugal
Qatar|QAT|QAR|qatar
Romania|ROU|RON|romania
Rwanda|RWA|RWF|rwanda
Saudi Arabia|SAU|SAR|saudi-arabia
Senegal|SEN|XOF|senegal
Serbia|SRB|RSD|serbia
Singapore|SGP|SGD|singapore
Slovakia|SVK|EUR|slovakia
Slovenia|SVN|EUR|slovenia
South Africa|ZAF|ZAR|south-africa
South Korea|KOR|KRW|south-korea
Spain|ESP|EUR|spain
Sri Lanka|LKA|LKR|sri-lanka
Sweden|SWE|SEK|sweden
Switzerland|CHE|CHF|switzerland
Taiwan|TWN|TWD|taiwan
Tanzania|TZA|TZS|tanzania
Thailand|THA|THB|thailand
Trinidad and Tobago|TTO|TTD|trinidad-and-tobago
Tunisia|TUN|TND|tunisia
Türkiye|TUR|TRY|turkey
Uganda|UGA|UGX|uganda
Ukraine|UKR|UAH|ukraine
United Arab Emirates|ARE|AED|united-arab-emirates
United Kingdom|GBR|GBP|united-kingdom
United States|USA|USD|united-states
Uruguay|URY|UYU|uruguay
Uzbekistan|UZB|UZS|uzbekistan
Vietnam|VNM|VND|vietnam
Zambia|ZMB|ZMW|zambia
Zimbabwe|ZWE|ZWG|zimbabwe
"""
COUNTRIES = {
    name: {"iso": iso, "currency": currency, "slug": slug}
    for name, iso, currency, slug in
    (line.split("|") for line in COUNTRY_TEXT.strip().splitlines())
}
ALIASES = {
    # Trading Economics titles and instrument labels use US / US 10Y.
    "United States": ("US", "U.S."),
    "United Kingdom": ("UK", "U.K."),
    "Czechia": ("Czech Republic",), "Türkiye": ("Turkey",),
    "South Korea": ("Korea", "Korea, Republic of"),
    "China": ("China (People's Republic of)",),
    "Hong Kong": ("Hong Kong (China)",), "Vietnam": ("Viet Nam",),
    "North Macedonia": ("Macedonia",),
}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def normalize(text: Any) -> str:
    return " ".join(unicodedata.normalize("NFKC", str(text)).split()).casefold()


def percent(value: float | None, places: int = 4) -> str:
    return "Unavailable" if value is None else f"{value * 100:.{places}f}%"


class DataUnavailable(RuntimeError):
    """An identified source is inaccessible or its schema is not recognized."""


class InputError(ValueError):
    """Missing, non-finite, or mathematically invalid project assumptions."""


@dataclass(frozen=True)
class Observation:
    value: float
    reference_date: str
    note: str = ""
    firm_count: int | None = None


@dataclass
class Dataset:
    source: str
    url: str
    retrieved_at: str
    observations: dict[str, Observation]
    note: str = ""


@dataclass(frozen=True)
class FetchFailure:
    source: str
    url: str
    attempted_at: str
    error: str


@dataclass
class SourceState:
    source: str
    url: str
    dataset: Dataset | None
    status: str
    attempted_at: str
    error: str = ""


@dataclass(frozen=True)
class Inputs:
    beta_u: float
    debt_share: float
    tax_rate: float
    erp: float
    rf_us: float
    rf_foreign: float
    spread_us: float
    spread_foreign_override: float | None = None


@dataclass(frozen=True)
class Results:
    debt_to_equity: float
    beta_e: float
    factor: float
    re_us: float
    rd_us: float
    rd_aftertax_us: float
    wacc_us: float
    spread_foreign_auto: float
    spread_foreign: float
    re_foreign: float
    rd_foreign: float
    rd_aftertax_foreign: float
    wacc_foreign: float
    equity_contribution_us: float
    debt_contribution_us: float
    equity_contribution_foreign: float
    debt_contribution_foreign: float


# ---------------------------------------------------------------------------
# Validation and pure financial calculations: independent of Streamlit/network.
# ---------------------------------------------------------------------------
def validate_inputs(p: Inputs) -> None:
    for name, value in asdict(p).items():
        if name == "spread_foreign_override" and value is None:
            continue
        if value is None:
            raise InputError(f"{name} is missing; retrieve it or enable its manual override.")
        try:
            finite = not isinstance(value, bool) and isinstance(value, (int, float)) and math.isfinite(value)
        except OverflowError:
            finite = False
        if not finite:
            raise InputError(f"{name} must be a finite number.")
    if not 0 <= p.debt_share < 1:
        raise InputError("Target debt share must be at least 0% and strictly below 100%.")
    if not 0 <= p.tax_rate <= 1:
        raise InputError("The statutory corporate tax rate must be between 0% and 100%.")
    if p.beta_u < 0 or p.erp < 0 or p.spread_us < 0:
        raise InputError("This model requires nonnegative beta, ERP, and credit spreads.")
    if p.spread_foreign_override is not None and p.spread_foreign_override < 0:
        raise InputError("The foreign credit spread must be nonnegative.")
    if p.rf_us <= -1 or p.rf_foreign <= -1:
        raise InputError("Both government yields must exceed -100%; gross-return factors must be positive.")


def calculate(p: Inputs) -> Results:
    validate_inputs(p)
    de = p.debt_share / (1 - p.debt_share)
    beta_e = p.beta_u * (1 + (1 - p.tax_rate) * de)
    re_us = p.rf_us + beta_e * p.erp
    rd_us = p.rf_us + p.spread_us
    after_us = rd_us * (1 - p.tax_rate)
    equity_us = (1 - p.debt_share) * re_us
    debt_us = p.debt_share * after_us
    factor = (1 + p.rf_foreign) / (1 + p.rf_us)
    re_foreign = (1 + re_us) * factor - 1
    spread_auto = p.spread_us * factor
    spread_foreign = (spread_auto if p.spread_foreign_override is None
                      else p.spread_foreign_override)
    rd_foreign = p.rf_foreign + spread_foreign
    after_foreign = rd_foreign * (1 - p.tax_rate)
    equity_foreign = (1 - p.debt_share) * re_foreign
    debt_foreign = p.debt_share * after_foreign
    # Convert equity and debt separately; do not convert the completed USD WACC.
    result = Results(de, beta_e, factor, re_us, rd_us, after_us,
                     equity_us + debt_us, spread_auto, spread_foreign,
                     re_foreign, rd_foreign, after_foreign,
                     equity_foreign + debt_foreign, equity_us, debt_us,
                     equity_foreign, debt_foreign)
    if not all(math.isfinite(v) for v in asdict(result).values()):
        raise InputError("These assumptions overflow numerical precision; reduce their magnitude.")
    return result


def comparison_yield(country: str, us: float | None, foreign: float | None) -> float | None:
    """The U.S. comparison is linked even when the U.S. yield is overridden."""
    return us if country == "United States" else foreign


def industry_comparison(betas: dict[str, Observation], p: Inputs, selected: str,
                        beta_override: bool) -> pd.DataFrame:
    rows = []
    names = set(betas)
    if beta_override:
        names.add(selected)
    for industry in sorted(names):
        if is_aggregate(industry):
            continue
        observation = betas.get(industry)
        overridden = industry == selected and beta_override
        beta_used = p.beta_u if overridden else observation.value
        result = calculate(replace(p, beta_u=beta_used))
        rows.append({
            "Industry": industry, "Selected": industry == selected,
            "Beta overridden": overridden, "Source beta": observation.value if observation else None,
            "U.S. firms": getattr(observation, "firm_count", None),
            "Beta used": beta_used, "Equity beta": result.beta_e,
            "USD equity (decimal)": result.re_us,
            "Foreign equity (decimal)": result.re_foreign,
            "USD pretax debt (decimal)": result.rd_us,
            "Foreign pretax debt (decimal)": result.rd_foreign,
            "USD after-tax debt (decimal)": result.rd_aftertax_us,
            "Foreign after-tax debt (decimal)": result.rd_aftertax_foreign,
            "USD WACC (decimal)": result.wacc_us,
            "Foreign WACC (decimal)": result.wacc_foreign,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# HTTP retrieval and source-specific parsers. Never substitute sample values.
# ---------------------------------------------------------------------------
def http_get(url: str, accept: str = "text/html", params: dict | None = None) -> requests.Response:
    """Two attempts for temporary failures; connect/read timeouts on each."""
    last_error = ""
    with requests.Session() as session:
        session.headers.update({"User-Agent": USER_AGENT, "Accept": accept})
        for attempt in range(2):
            try:
                response = session.get(url, params=params, timeout=(4, 12))
            except (requests.Timeout, requests.ConnectionError) as exc:
                last_error = f"{type(exc).__name__}: {exc}"
                if attempt == 0:
                    time.sleep(0.4)
                    continue
                raise DataUnavailable(f"Request failed for {url}: {last_error}") from exc
            except requests.RequestException as exc:
                raise DataUnavailable(f"Request failed for {url}: {exc}") from exc
            if response.status_code in (429, 500, 502, 503, 504) and attempt == 0:
                time.sleep(0.5)
                continue
            if not response.ok:
                raise DataUnavailable(f"HTTP {response.status_code} from {url}.")
            return response
    raise DataUnavailable(f"No response from {url}: {last_error}")


def parse_number(value: Any) -> float | None:
    """Return source units: '1.5%' becomes 1.5, NOT 0.015. Callers /100."""
    if isinstance(value, bool):
        raise DataUnavailable("A Boolean is not a valid financial observation.")
    if isinstance(value, (float, int)):
        try:
            number = float(value)
        except OverflowError as exc:
            raise DataUnavailable("Numeric source field exceeds supported precision.") from exc
        return number if math.isfinite(number) else None
    text = str(value).strip().replace("\u2212", "-").replace("\xa0", "")
    if normalize(text) in ("", "na", "n/a", "nan", "none", "..", "-", "—"):
        return None
    if not re.fullmatch(r"[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?|\.\d+)%?", text):
        raise DataUnavailable(f"Unrecognized numeric field: {text[:80]!r}.")
    number = float(text.rstrip("%").replace(",", ""))
    return number if math.isfinite(number) else None


def table_grid(table: Any) -> list[list[str]]:
    """Expand HTML rowspan/colspan so rating groups retain their column identity."""
    cells: dict[tuple[int, int], str] = {}
    trs = table.find_all("tr")
    width = 0
    for row_number, tr in enumerate(trs):
        column = 0
        for cell in tr.find_all(["td", "th"], recursive=False):
            while (row_number, column) in cells:
                column += 1
            try:
                rowspan = max(1, min(int(cell.get("rowspan", 1)), 100))
                colspan = max(1, min(int(cell.get("colspan", 1)), 100))
            except ValueError as exc:
                raise DataUnavailable("Malformed HTML table spans.") from exc
            # Excel HTML sometimes splits a word across overflow spans. Inline
            # fragments must join directly; explicit line breaks still separate words.
            for br in cell.find_all("br"):
                br.replace_with(" ")
            cell_text = " ".join(cell.get_text("", strip=False).split())
            for rr in range(rowspan):
                for cc in range(colspan):
                    cells[row_number + rr, column + cc] = cell_text
            column += colspan
            width = max(width, column)
    return [[cells.get((r, c), "") for c in range(width)] for r in range(len(trs))]


def published_date(text: str) -> str:
    match = re.search(
        r"(?:Date of Analysis|Date updated)\s*:\s*(?:Data used is as of\s*)?"
        r"([A-Za-z]+\s+\d{4}|\d{4}-\d{2}-\d{2})", text, re.I)
    return match.group(1) if match else "Not published / not identified"


def is_aggregate(industry: str) -> bool:
    text = normalize(industry)
    return text.startswith("total market") or text in ("market", "total", "all industries")


def beta_rows(grid: list[list[Any]]) -> dict[str, float]:
    for r, row in enumerate(grid):
        headers = [normalize(value) for value in row]
        if "industry name" not in headers or "unlevered beta" not in headers:
            continue
        industry_column = headers.index("industry name")
        # Exact match intentionally excludes 'Beta' and cash-corrected beta.
        beta_column = headers.index("unlevered beta")
        values = {}
        for data_row in grid[r + 1:]:
            if max(industry_column, beta_column) >= len(data_row):
                continue
            industry = " ".join(str(data_row[industry_column]).split())
            if not industry or is_aggregate(industry):
                continue
            beta = parse_number(data_row[beta_column])
            if beta is None:
                continue
            if beta < 0:
                raise DataUnavailable(f"Negative source beta for {industry}; this model requires nonnegative betas.")
            if industry in values:
                raise DataUnavailable(f"Duplicate industry identity: {industry}.")
            values[industry] = beta
        if values:
            return values
    raise DataUnavailable("Could not identify 'Industry Name' and exact 'Unlevered beta' columns.")


def beta_firm_counts(grid: list[list[Any]]) -> dict[str, int]:
    """Optional source metadata; an unreadable count never replaces a missing beta."""
    for r, row in enumerate(grid):
        headers = [normalize(value) for value in row]
        if "industry name" not in headers or "number of firms" not in headers:
            continue
        name_col, count_col = headers.index("industry name"), headers.index("number of firms")
        counts = {}
        for data_row in grid[r + 1:]:
            if max(name_col, count_col) >= len(data_row):
                continue
            name = " ".join(str(data_row[name_col]).split())
            if not name or is_aggregate(name):
                continue
            try:
                count = parse_number(data_row[count_col])
            except DataUnavailable:
                continue
            if count is not None and count >= 0 and count.is_integer():
                counts[name] = int(count)
        return counts
    return {}


def parse_beta_xls(content: bytes, firm_counts: dict[str, int] | None = None) -> tuple[dict[str, float], str]:
    try:
        book = xlrd.open_workbook(file_contents=content)
    except (xlrd.XLRDError, xlrd.compdoc.CompDocError, ValueError, IndexError, struct.error, EOFError) as exc:
        raise DataUnavailable(f"Damodaran XLS could not be decoded: {exc}") from exc
    # The first sheet is a glossary. Inspect every sheet for the actual header.
    for sheet in book.sheets():
        grid = [sheet.row_values(r) for r in range(sheet.nrows)]
        if not any("industry name" in [normalize(v) for v in row]
                   and "unlevered beta" in [normalize(v) for v in row] for row in grid[:30]):
            continue
        date = "Not published / not identified"
        for r in range(min(30, sheet.nrows)):
            for c in range(sheet.ncols - 1):
                if "date updated" in normalize(sheet.cell_value(r, c)):
                    cell = sheet.cell(r, c + 1)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            date = xlrd.xldate_as_datetime(cell.value, book.datemode).date().isoformat()
                        except (xlrd.xldate.XLDateError, ValueError, OverflowError) as exc:
                            raise DataUnavailable("Damodaran workbook's publication date is invalid.") from exc
                    else:
                        date = published_date(f"Date updated: {cell.value}")
        if firm_counts is not None:
            firm_counts.update(beta_firm_counts(grid))
        return beta_rows(grid), date
    raise DataUnavailable("Damodaran workbook has no recognized U.S. industry beta sheet.")


def parse_beta_html(text: str, firm_counts: dict[str, int] | None = None) -> tuple[dict[str, float], str]:
    soup = BeautifulSoup(text, "html.parser")
    for table in soup.find_all("table"):
        grid = table_grid(table)
        if any("industry name" in [normalize(v) for v in row]
               and "unlevered beta" in [normalize(v) for v in row] for row in grid):
            if firm_counts is not None:
                firm_counts.update(beta_firm_counts(grid))
            return beta_rows(grid), published_date(soup.get_text(" ", strip=True))
    raise DataUnavailable("Official beta HTML has no recognized unlevered-beta table.")


def fetch_betas() -> Dataset:
    errors = []
    for url, parser in ((BETA_XLS, parse_beta_xls), (BETA_HTML, parse_beta_html)):
        try:
            response = http_get(url, "*/*")
            counts = {}
            betas, reference = parser(response.content if url == BETA_XLS else response.text, counts)
            if len(betas) < 20:
                raise DataUnavailable("Fewer than 20 valid industries; refusing a likely incomplete table.")
            note = "Exact Unlevered beta column; U.S. industries; aggregate market rows excluded."
            if errors:
                note += " Official HTML fallback used, at its published rounded precision. " + " | ".join(errors)
            return Dataset("Damodaran U.S. industry betas", url, utc_now(),
                           {name: Observation(beta, reference,
                                              f"U.S. firms: {counts[name]}" if name in counts else "Firm count unavailable.",
                                              counts.get(name)) for name, beta in betas.items()}, note)
        except DataUnavailable as exc:
            errors.append(str(exc))
    raise DataUnavailable("Beta retrieval failed: " + " | ".join(errors))


def parse_bbb_spread(text: str) -> tuple[float, str]:
    soup = BeautifulSoup(text, "html.parser")
    for table in soup.find_all("table"):
        grid = table_grid(table)
        for marker_index, row in enumerate(grid):
            markers = [c for c, value in enumerate(row)
                       if "large non-financial service firms" in normalize(value)]
            if not markers:
                continue
            left = min(markers)
            right = next((c for c in range(max(markers) + 1, len(row))
                          if "financial service firms" in normalize(row[c])
                          and "large non-financial" not in normalize(row[c])), len(row))
            for header_index in range(marker_index + 1, min(marker_index + 7, len(grid))):
                header = grid[header_index]
                rating_columns = [c for c in range(left, right)
                                  if normalize(header[c]).startswith("rating")]
                for rating_column in rating_columns:
                    spread_column = next((c for c in range(rating_column + 1, right)
                                          if normalize(header[c]).startswith("spread")), None)
                    if spread_column is None:
                        continue
                    for data_row in grid[header_index + 1:]:
                        if normalize(data_row[rating_column]).replace(" ", "") == "baa2/bbb":
                            spread = parse_number(data_row[spread_column])
                            if spread is None or spread < 0:
                                raise DataUnavailable("BBB spread is missing or negative.")
                            # Source is in percentage points: 1.11% -> 0.0111.
                            return spread / 100, published_date(soup.get_text(" ", strip=True))
    raise DataUnavailable("Baa2/BBB spread for large non-financial service firms was not identified.")


def fetch_spread() -> Dataset:
    response = http_get(RATINGS_URL)
    spread, date = parse_bbb_spread(response.text)
    return Dataset("Damodaran Baa2/BBB credit spread", RATINGS_URL, utc_now(),
                   {"BBB": Observation(spread, date)},
                   "Large non-financial service firms; spread, not coverage ratio.")


def parse_oecd_csv(text: str, current_year: int) -> dict[str, Observation]:
    try:
        return _parse_oecd_rows(csv.DictReader(io.StringIO(text.lstrip("\ufeff"))), current_year)
    except csv.Error as exc:
        raise DataUnavailable(f"OECD response could not be parsed as CSV: {exc}") from exc


def _parse_oecd_rows(reader: csv.DictReader, current_year: int) -> dict[str, Observation]:
    required = {"REF_AREA", "FREQ", "MEASURE", "TARGETING", "UNIT_MEASURE", "SECTOR",
                "DIV_TREATMENT", "RATE_STRUCTURE", "TAX_BASE", "TIME_PERIOD", "OBS_VALUE", "UNIT_MULT"}
    if not required.issubset(reader.fieldnames or []):
        raise DataUnavailable("OECD CSV schema does not contain the required statutory-rate dimensions.")
    if len(reader.fieldnames) != len(set(reader.fieldnames)):
        raise DataUnavailable("OECD CSV contains duplicate column names.")
    latest: dict[str, Observation] = {}
    exact = {"FREQ": "A", "MEASURE": "CIT_C", "TARGETING": "ST",
             "UNIT_MEASURE": "PT_INC_TAX", "SECTOR": "S13",
             "DIV_TREATMENT": "_Z", "RATE_STRUCTURE": "_Z", "TAX_BASE": "_Z"}
    for row in reader:
        if None in row or any(row.get(field) is None for field in required):
            raise DataUnavailable("OECD CSV contains a truncated or misaligned observation row.")
        if any(row.get(k) != v for k, v in exact.items()):
            continue
        if not re.fullmatch(r"\d{4}", row["TIME_PERIOD"]):
            raise DataUnavailable("OECD annual reference year is not recognized.")
        year = int(row["TIME_PERIOD"])
        if year > current_year:
            continue
        if parse_number(row["UNIT_MULT"]) != 0:
            raise DataUnavailable("OECD unit multiplier changed; refusing an ambiguous tax-rate scale.")
        quoted = parse_number(row["OBS_VALUE"])
        if quoted is None:
            continue
        rate = quoted / 100
        if not 0 <= rate <= 1:
            raise DataUnavailable("OECD combined statutory rate is outside 0–100%.")
        code = row["REF_AREA"]
        old = latest.get(code)
        observation = Observation(rate, str(year), "OECD observation status: " + row.get("OBS_STATUS", "not supplied"))
        if old and old.reference_date == str(year) and old.value != rate:
            raise DataUnavailable(f"Ambiguous OECD observations for {code}, {year}.")
        if old is None or year > int(old.reference_date):
            latest[code] = observation
    if not latest:
        raise DataUnavailable("OECD returned no combined general-government statutory rates.")
    return latest


def fetch_taxes() -> Dataset:
    year = datetime.now(timezone.utc).year
    errors = []
    # These are two documented routes to the same public SDMX service, not paid APIs.
    for base in (OECD_API, OECD_API.replace("/rest/v1/", "/rest/")):
        try:
            response = http_get(base, "text/csv;version=2.0",
                                {"startPeriod": "2000", "dimensionAtObservation": "AllDimensions"})
            observations = parse_oecd_csv(response.text, year)
            return Dataset("OECD combined statutory corporate income tax", response.url,
                           utc_now(), observations,
                           "CIT_C / ST / S13 / PT_INC_TAX; latest available year per jurisdiction, "
                           "up to the current calendar year. Reference year is not retrieval date.")
        except DataUnavailable as exc:
            errors.append(str(exc))
    raise DataUnavailable("OECD statutory tax retrieval failed: " + " | ".join(errors))


def country_names(country: str) -> tuple[str, ...]:
    return (country,) + ALIASES.get(country, ())


def name_matches(text: str, country: str) -> bool:
    """Exact country identity after removing an explicit 10-year maturity suffix."""
    identity = re.sub(r"\s+(?:10\s*[- ]?\s*(?:y|yr|year|years)(?:\b|$)|government bond).*$", "", normalize(text))
    return identity in {normalize(name) for name in country_names(country)}


def has_ten_year(text: str) -> bool:
    return bool(re.search(r"\b10\s*[- ]?\s*(?:y\b|yr\b|year\b|years\b)", text, re.I))


def inflation_linked(text: str) -> bool:
    return bool(re.search(r"\b(?:tips|real|inflation[-\s]*(?:linked|indexed|protected))\b", text, re.I))


def nominal_ten_year(text: str) -> bool:
    return has_ten_year(text) and not inflation_linked(text)


def nominal_symbol_matches(symbol: Any, country: str) -> bool:
    """Reject contradictory maturity identifiers, including the known U.S. nominal series."""
    if symbol is None or symbol == "":
        return True
    if not isinstance(symbol, str):
        return False
    symbol = symbol.upper().strip()
    if country == "United States":
        return symbol in ("USGG10YR:IND", "USGG10YR")
    maturity = re.search(r"(\d+)(?:YR|Y)(?=:|$)", symbol)
    return not maturity or int(maturity.group(1)) == 10


def observation_date(soup: BeautifulSoup, row: Any = None, country: str | None = None) -> str:
    # Dataset temporal coverage identifies its latest observation, not page refresh.
    if country:
        for script in soup.find_all("script", attrs={"type": "application/ld+json"}):
            try:
                structured = json.loads(script.get_text())
            except json.JSONDecodeError:
                continue
            objects = structured if isinstance(structured, list) else [structured]
            for obj in objects:
                if not isinstance(obj, dict):
                    continue
                graph = obj.get("@graph", [])
                if not isinstance(graph, list):
                    continue
                for dataset in [obj] + graph:
                    if not isinstance(dataset, dict) or dataset.get("@type") != "Dataset":
                        continue
                    name = str(dataset.get("name", ""))
                    if not (nominal_ten_year(name) and name_matches(name, country)):
                        continue
                    end = str(dataset.get("temporalCoverage", "")).split("/")[-1]
                    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", end):
                        try:
                            return datetime.strptime(end, "%Y-%m-%d").date().isoformat()
                        except ValueError:
                            continue
    # Inspect only a current summary/description, never dated news or forecasts.
    summary = soup.select_one("#market-summary, #historical-desc #description") if country else None
    description = soup.find("meta", attrs={"name": "description"}) if country else None
    text = summary.get_text(" ", strip=True) if summary else description.get("content", "") if description else ""
    match = re.search(r"\bon\s+([A-Z][a-z]+\s+\d{1,2},\s+\d{4})\b", text)
    if match:
        try:
            return datetime.strptime(match.group(1), "%B %d, %Y").date().isoformat()
        except ValueError:
            pass
    if row is not None:
        node = row.select_one('[id="date"], time, .market-date')
        if node and node.get_text(strip=True):
            value = node.get("datetime") or node.get_text(" ", strip=True)
            return value + (" (source omits year/time zone)" if not re.search(r"\b\d{4}\b", value) else "")
    return "Not provided / not identified"


def parse_yield_html(text: str, country: str, board: bool = False) -> tuple[float, str, str]:
    soup = BeautifulSoup(text, "html.parser")
    country_url = f"/{COUNTRIES[country]['slug']}/government-bond-yield"
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    if re.search(r"\b(?:403|access denied|just a moment)\b", title, re.I):
        raise DataUnavailable(f"Trading Economics returned an access-restriction page for {country}.")
    description = soup.find("meta", attrs={"name": "description"}) or soup.find("meta", id="metaDesc")
    description_text = description.get("content", "") if description else ""
    board_ten_year = board and nominal_ten_year(description_text + " " + title)
    page_identity = (nominal_ten_year(title) and any(
        normalize(title).startswith(normalize(name) + " ") for name in country_names(country)))
    metadata_problems = []
    # Public embedded JSON can retain more precision than the displayed row.
    if page_identity:
        for script in soup.find_all("script"):
            script_text = script.get_text()
            match = re.search(r"\bTEChartsMeta\s*=\s*(\[)", script_text)
            if not match:
                continue
            try:
                metadata, _ = json.JSONDecoder().raw_decode(script_text[match.start(1):])
            except json.JSONDecodeError:
                metadata_problems.append("Public quote metadata was malformed.")
                continue
            if not isinstance(metadata, list):
                continue
            for item in metadata:
                if not isinstance(item, dict):
                    continue
                identity = str(item.get("full_name") or item.get("name") or "")
                if not nominal_ten_year(identity) or not name_matches(identity, country):
                    continue
                if not all(nominal_symbol_matches(item.get(field), country) for field in ("symbol", "ticker")):
                    continue
                if item.get("type") and normalize(item["type"]) != "bond":
                    continue
                if item.get("unit") and normalize(item["unit"]) not in ("percent", "%"):
                    continue
                try:
                    quote = parse_number(item.get("last", item.get("value")))
                except DataUnavailable:
                    metadata_problems.append("Public quote metadata had an invalid numeric value.")
                    continue
                if quote is not None and quote > -100:
                    row = next((a.find_parent("tr") for a in soup.find_all("a", href=country_url)
                                if a.find_parent("tr") is not None
                                and name_matches(a.get_text(" ", strip=True), country)
                                and not inflation_linked(a.get_text(" ", strip=True))
                                and nominal_symbol_matches(a.find_parent("tr").get("data-symbol"), country)), None)
                    return quote / 100, observation_date(soup, row, country), "Public page's current 10-year quote metadata."
    # Row fallback requires both exact country path and a verified 10-year maturity.
    for anchor in soup.find_all("a", href=country_url):
        row = anchor.find_parent("tr")
        label = anchor.get_text(" ", strip=True)
        if row is None or not name_matches(label, country) or inflation_linked(label):
            continue
        if not nominal_symbol_matches(row.get("data-symbol"), country):
            continue
        if not (board_ten_year or page_identity or nominal_ten_year(label)):
            continue
        value_cell = row.select_one('[id="p"], .market-widget-last')
        if value_cell is None:
            continue
        try:
            quote = parse_number(value_cell.get_text(" ", strip=True))
        except DataUnavailable:
            continue
        if quote is not None and quote > -100:
            note = "Public nominal 10-year yield table; source display precision."
            if metadata_problems:
                note += " Table fallback used. " + " ".join(dict.fromkeys(metadata_problems))
            return quote / 100, observation_date(soup, row, None if board else country), note
    detail = " " + " ".join(dict.fromkeys(metadata_problems)) if metadata_problems else ""
    raise DataUnavailable(f"An exact {country} nominal 10-year government yield was not found on the public page." + detail)


def fetch_yield(country: str) -> Dataset:
    errors = []
    country_url = f"https://tradingeconomics.com/{COUNTRIES[country]['slug']}/government-bond-yield"
    for url in (BONDS_URL, country_url):
        try:
            response = http_get(url)
            value, reference, note = parse_yield_html(response.text, country, board=url == BONDS_URL)
            if errors:
                note += " Country-page fallback used. " + " | ".join(errors)
            return Dataset(f"Trading Economics — {country} 10-year yield", url, utc_now(),
                           {"yield": Observation(value, reference)}, note)
        except DataUnavailable as exc:
            errors.append(str(exc))
    raise DataUnavailable(f"{country} yield unavailable: " + " | ".join(errors))


def retrieve_market_bundle(country: str) -> dict[str, Dataset | FetchFailure]:
    jobs = {
        "betas": (fetch_betas, "Damodaran U.S. industry betas", BETA_XLS),
        "taxes": (fetch_taxes, "OECD combined statutory corporate income tax", OECD_API),
        "spread": (fetch_spread, "Damodaran Baa2/BBB credit spread", RATINGS_URL),
        "yield:USA": (lambda: fetch_yield("United States"), "U.S. 10-year yield", BONDS_URL),
    }
    if country != "United States":
        jobs[f"yield:{COUNTRIES[country]['iso']}"] = (
            lambda: fetch_yield(country), f"{country} 10-year yield", BONDS_URL)
    outputs = {}
    with ThreadPoolExecutor(max_workers=5) as pool:
        futures = {pool.submit(function): (key, source, url)
                   for key, (function, source, url) in jobs.items()}
        for future in as_completed(futures):
            key, source, url = futures[future]
            try:
                outputs[key] = future.result()
            except DataUnavailable as exc:
                outputs[key] = FetchFailure(source, url, utc_now(), str(exc))
    return outputs


def retrieve_for_revision(country: str, revision: str) -> dict[str, Dataset | FetchFailure]:
    """Include the application revision in Streamlit's cache key after parser upgrades."""
    return retrieve_market_bundle(country)


def reconcile_sources(bundle: dict[str, Dataset | FetchFailure], last_good: dict[str, Dataset],
                      ttl: int) -> dict[str, SourceState]:
    states = {}
    for key, result in bundle.items():
        if isinstance(result, Dataset):
            known = last_good.get(key)
            if known and datetime.fromisoformat(known.retrieved_at) > datetime.fromisoformat(result.retrieved_at):
                result = known
            last_good[key] = result
            age = (datetime.now(timezone.utc) - datetime.fromisoformat(result.retrieved_at)).total_seconds()
            status = "Cached response" if age > 5 else "Retrieved"
            if age > ttl:
                status = "STALE — cache lifetime exceeded"
            states[key] = SourceState(result.source, result.url, result, status, result.retrieved_at)
        else:
            old = last_good.get(key)
            # A country-keyed cached failure may predate a later successful common-source retrieval.
            if old and datetime.fromisoformat(old.retrieved_at) > datetime.fromisoformat(result.attempted_at):
                age = (datetime.now(timezone.utc) - datetime.fromisoformat(old.retrieved_at)).total_seconds()
                status = "Cached response — newer successful retrieval"
                if age > ttl:
                    status = "STALE — cache lifetime exceeded"
                states[key] = SourceState(old.source, old.url, old, status, old.retrieved_at)
                continue
            states[key] = SourceState(result.source, old.url if old else result.url, old,
                                      "STALE — refresh failed" if old else "Unavailable",
                                      result.attempted_at, result.error)
    return states


def lookup(state: SourceState, key: str) -> Observation | None:
    return state.dataset.observations.get(key) if state.dataset else None


# ---------------------------------------------------------------------------
# UI state, provenance, exports, and presentation.
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class EffectiveInput:
    name: str
    value: float | None
    automatic_value: float | None
    override_active: bool
    source: str
    source_url: str
    reference_date: str
    retrieved_at_utc: str
    status: str
    notes: str


def reset_manual_inputs() -> None:
    st.session_state["manual_store"] = {}
    for key in list(st.session_state):
        if key.startswith(("override_toggle:", "override_value:")):
            del st.session_state[key]


def override_control(name: str, key: str, observation: Observation | None, state: SourceState,
                     percentage: bool = True, minimum: float = 0.0,
                     maximum: float | None = None, linked: EffectiveInput | None = None) -> EffectiveInput:
    store = st.session_state.setdefault("manual_store", {})
    auto = observation.value if observation else None
    scale = 100.0 if percentage else 1.0
    seed = auto * scale if auto is not None else None
    if seed is not None and not math.isfinite(seed):
        seed = None
    saved = store.setdefault(key, {"enabled": False, "value": seed})
    toggle_key, value_key = "override_toggle:" + key, "override_value:" + key
    if toggle_key not in st.session_state:
        st.session_state[toggle_key] = saved["enabled"]
    if value_key not in st.session_state:
        st.session_state[value_key] = saved["value"]
    enabled = st.toggle("Override " + name, key=toggle_key, disabled=linked is not None)
    entered = st.number_input(
        "Saved manual " + name + (" (%)" if percentage else ""), min_value=minimum, max_value=maximum,
        step=0.0001 if percentage else 0.01, format="%.6f", key=value_key,
        disabled=not enabled or linked is not None, placeholder="Enter a verified value",
        help="Saved manual entry, used only while this override is enabled. Automatic and Used values below show the current calculation inputs.")
    saved.update(enabled=enabled, value=entered)
    used = (entered / scale if entered is not None else None) if enabled else auto
    if linked is not None:
        used = linked.value
        enabled = False
    format_value = percent if percentage else lambda v: "Unavailable" if v is None else f"{v:.8f}"
    st.caption(f"Automatic: {format_value(auto)} · Used: {format_value(used)}")
    if linked is not None:
        st.caption("United States comparison shares the U.S. yield, including its override.")
    if auto is None:
        st.caption("No matching automatic observation. Enable this override and enter a verified value if needed.")
    notes = " | ".join(filter(None, [observation.note if observation else "No matching source observation.",
                                    state.dataset.note if state.dataset else "", state.error]))
    if linked is not None:
        notes += " | Linked to effective U.S. yield; U.S. override active: " + str(linked.override_active)
    return EffectiveInput(name, used, auto, enabled, state.source, state.url,
                          observation.reference_date if observation else "Unavailable",
                          state.dataset.retrieved_at if state.dataset else "",
                          state.status if observation else "Unavailable", notes)


def project_csv(p: Inputs, result: Results, records: list[EffectiveInput], industry: str,
                country: str, tax_selection: str, generated_at: str) -> bytes:
    currency = COUNTRIES[country]["currency"]
    rows = []
    metadata = {"industry": industry, "country": country, "home_currency": "USD",
                "selected_currency": currency, "tax_selection": tax_selection,
                "generated_at_utc": generated_at,
                "active_overrides": json.dumps([r.name for r in records if r.override_active])}
    for field, value in metadata.items():
        rows.append({"section": "metadata", "field": field, "value": value, "unit": "text"})
    for field, value in asdict(p).items():
        rows.append({"section": "calculation_input", "field": field, "value": value,
                     "unit": "dimensionless" if field == "beta_u" else "decimal rate"})
    for record in records:
        row = asdict(record)
        row.update(section="market_input", field=row.pop("name"), unit=(
            "dimensionless" if record.name == "Unlevered beta" else "decimal rate"))
        rows.append(row)
    for field, value in asdict(result).items():
        scenario = ("USD" if field.endswith("_us") else currency
                    if field.endswith("_foreign") or field == "spread_foreign_auto" else "shared/derived")
        rows.append({"section": "result", "field": field, "value": value, "currency": scenario,
                     "unit": "dimensionless" if field in ("debt_to_equity", "beta_e", "factor") else "decimal rate"})
    # No display rounding: downloads contain the same full-precision results.
    return pd.DataFrame(rows).to_csv(index=False, float_format="%.17g").encode("utf-8-sig")


def all_industry_export(frame: pd.DataFrame, p: Inputs, country: str, beta_state: SourceState,
                        records: list[EffectiveInput], generated_at: str) -> bytes:
    output = frame.copy()
    output["Country"] = country
    output["Foreign currency"] = COUNTRIES[country]["currency"]
    for field, value in asdict(p).items():
        if field != "beta_u":
            output[field] = value
    output["beta_source_url"] = beta_state.url
    output["beta_retrieved_at_utc"] = beta_state.dataset.retrieved_at if beta_state.dataset else ""
    output["beta_reference_date"] = (next(iter(beta_state.dataset.observations.values())).reference_date
                                      if beta_state.dataset and beta_state.dataset.observations else "")
    output["beta_source_status"] = beta_state.status
    output["generated_at_utc"] = generated_at
    output["input_provenance_json"] = json.dumps([asdict(record) for record in records], ensure_ascii=False)
    return output.to_csv(index=False, float_format="%.17g").encode("utf-8-sig")


def render_comparison(p: Inputs, r: Results, country: str) -> None:
    currency = COUNTRIES[country]["currency"]
    a, b, c = st.columns([1, 1, 0.8])
    a.metric("USD WACC", percent(r.wacc_us))
    b.metric(f"{currency} WACC · {country}", percent(r.wacc_foreign))
    c.metric("Currency factor F", f"{r.factor:.8f}×")
    st.caption("Nominal returns depend on currency. Discount project cash flows in the same currency as the WACC.")
    if currency == "USD" and country != "United States":
        st.info("Both scenarios are denominated in USD. Any difference reflects the selected country's government-yield "
                "proxy and its sovereign-risk components. Interpret the assignment's relative-yield factor cautiously "
                "for this same-currency comparison.")
    metrics = [
        ("Risk-free yield", percent(p.rf_us), percent(p.rf_foreign)),
        ("Unlevered beta", f"{p.beta_u:.8f}", f"{p.beta_u:.8f}"),
        ("Relevered equity beta", f"{r.beta_e:.8f}", f"{r.beta_e:.8f}"),
        ("Cost of equity", percent(r.re_us), percent(r.re_foreign)),
        ("Pretax borrowing cost", percent(r.rd_us), percent(r.rd_foreign)),
        ("After-tax borrowing cost", percent(r.rd_aftertax_us), percent(r.rd_aftertax_foreign)),
        ("WACC", percent(r.wacc_us), percent(r.wacc_foreign)),
    ]
    st.dataframe(pd.DataFrame(metrics, columns=["Metric", "USD financing", f"{currency} financing · {country}"]),
                 hide_index=True, width="stretch")
    chart = go.Figure()
    labels = ["Cost of equity", "After-tax debt", "WACC"]
    for label, values, color in [
        ("USD financing", [r.re_us, r.rd_aftertax_us, r.wacc_us], "#253858"),
        (f"{currency} financing · {country}", [r.re_foreign, r.rd_aftertax_foreign, r.wacc_foreign], "#00A58E"),
    ]:
        chart.add_bar(name=label, x=labels, y=values, marker_color=color,
                      hovertemplate="%{x}: %{y:.4%}<extra>%{fullData.name}</extra>")
    chart.update_layout(barmode="group", height=350, margin=dict(l=5, r=5, t=20, b=0),
                        yaxis_tickformat=".2%", legend=dict(orientation="h", y=1.12))
    st.plotly_chart(chart, width="stretch")
    if st.checkbox("Show WACC contribution breakdown"):
        breakdown = go.Figure()
        scenarios = ["USD financing", f"{currency} financing · {country}"]
        breakdown.add_bar(name="Equity contribution", x=scenarios,
                          y=[r.equity_contribution_us, r.equity_contribution_foreign], marker_color="#253858")
        breakdown.add_bar(name="After-tax debt contribution", x=scenarios,
                          y=[r.debt_contribution_us, r.debt_contribution_foreign], marker_color="#00A58E")
        breakdown.update_layout(barmode="relative", height=300, yaxis_tickformat=".2%",
                                margin=dict(l=5, r=5, t=15, b=0))
        st.plotly_chart(breakdown, width="stretch")


def render_methodology(p: Inputs | None, r: Results | None) -> None:
    with st.expander("1 · Capital structure and project beta", expanded=True):
        st.latex(r"D/E=\frac{d}{1-d},\qquad \beta_E=\beta_U\left[1+(1-t)D/E\right]")
        st.write("Debt and tax-shield betas are assumed to be zero. The same project equity beta applies to both currencies.")
        if p and r:
            st.code(f"d = {p.debt_share:.10f}; t = {p.tax_rate:.10f}; beta_U = {p.beta_u:.10f}\n"
                    f"D/E = {p.debt_share:.10f} / (1 - {p.debt_share:.10f}) = {r.debt_to_equity:.10f}\n"
                    f"beta_E = {p.beta_u:.10f} × [1 + (1 - {p.tax_rate:.10f}) × {r.debt_to_equity:.10f}] = {r.beta_e:.10f}")
    with st.expander("2 · USD CAPM, borrowing costs, and WACC"):
        st.latex(r"R_{e,US}=R_{f,US}+\beta_E ERP,\qquad R_{d,US}=R_{f,US}+s_{US}")
        st.latex(r"R_{d,after,US}=R_{d,US}(1-t),\qquad WACC_{US}=(1-d)R_{e,US}+dR_{d,US}(1-t)")
        if p and r:
            st.code(f"Re_US = {p.rf_us:.10f} + {r.beta_e:.10f} × {p.erp:.10f} = {r.re_us:.10f}\n"
                    f"Rd_US = {p.rf_us:.10f} + {p.spread_us:.10f} = {r.rd_us:.10f}\n"
                    f"Rd_aftertax_US = {r.rd_us:.10f} × (1 - {p.tax_rate:.10f}) = {r.rd_aftertax_us:.10f}\n"
                    f"WACC_US = {r.equity_contribution_us:.10f} + {r.debt_contribution_us:.10f} = {r.wacc_us:.10f}")
    with st.expander("3 · Currency factor, foreign equity, spread, and WACC"):
        st.latex(r"F=\frac{1+R_{f,FOREIGN}}{1+R_{f,US}},\qquad R_{e,FOREIGN}=(1+R_{e,US})F-1")
        st.latex(r"s_{FOREIGN}=s_{US}F\;\text{(unless overridden)},\qquad R_{d,FOREIGN}=R_{f,FOREIGN}+s_{FOREIGN}")
        st.latex(r"R_{d,after,FOREIGN}=R_{d,FOREIGN}(1-t)")
        st.latex(r"WACC_{FOREIGN}=(1-d)R_{e,FOREIGN}+dR_{d,FOREIGN}(1-t)")
        st.write("This multiplicative gross-return adjustment reflects Fisher-equation and relative-PPP assumptions. "
                 "It is not a simple yield-difference addition. Foreign WACC is recomputed from equity and debt.")
        if p and r:
            st.code(f"F = (1 + {p.rf_foreign:.10f}) / (1 + {p.rf_us:.10f}) = {r.factor:.10f}\n"
                    f"Re_FOREIGN = (1 + {r.re_us:.10f}) × {r.factor:.10f} - 1 = {r.re_foreign:.10f}\n"
                    f"Automatic spread = {p.spread_us:.10f} × {r.factor:.10f} = {r.spread_foreign_auto:.10f}\n"
                    f"Spread used = {r.spread_foreign:.10f}; override = {p.spread_foreign_override is not None}\n"
                    f"Rd_FOREIGN = {p.rf_foreign:.10f} + {r.spread_foreign:.10f} = {r.rd_foreign:.10f}\n"
                    f"Rd_aftertax_FOREIGN = {r.rd_foreign:.10f} × (1 - {p.tax_rate:.10f}) = {r.rd_aftertax_foreign:.10f}\n"
                    f"WACC_FOREIGN = {r.equity_contribution_foreign:.10f} + {r.debt_contribution_foreign:.10f} = {r.wacc_foreign:.10f}")
    with st.expander("4 · Economic assumptions and the tax shield"):
        st.markdown(
            "- U.S. industry unlevered beta proxies project operating risk; relevering adjusts equity risk for target debt and taxes.\n"
            "- Government yields proxy nominal risk-free rates and may contain sovereign, liquidity, and term-risk components.\n"
            "- CAPM uses a U.S. market-based ERP. No separate country equity premium is added.\n"
            "- The target market-value debt share is constant. Both scenarios use the same beta, ERP, and project tax rate.\n"
            "- Interest deductions are assumed usable. The tax shield lowers the modeled debt cost by multiplying it by (1 − t).\n"
            "- The foreign BBB spread scales by F unless overridden. It is a model assumption, not an observed local BBB quote.\n"
            "- Relative risk-free yields approximate nominal currency conversion under Fisher and relative-PPP assumptions; "
            "this is not an exchange-rate forecast or a full sovereign-risk model.\n"
            "- A lower nominal WACC in one currency does not establish an economic financing advantage. Cash flows and discount rates must share a currency.")


def render_sources(states: dict[str, SourceState], records: list[EffectiveInput]) -> None:
    st.caption(f"Application revision: {APP_REVISION}")
    for key, state in sorted(states.items()):
        with st.expander(f"{state.source} · {state.status}"):
            st.markdown(f"[Source URL]({state.url})")
            st.write("Last attempt (UTC):", state.attempted_at)
            if state.dataset:
                st.write("Successful retrieval (UTC):", state.dataset.retrieved_at)
                st.write("Publication / observation references:", ", ".join(sorted(
                    {o.reference_date for o in state.dataset.observations.values()})))
                st.write(state.dataset.note)
            if state.error:
                st.warning(state.error)
    st.markdown("**Automatic observations and effective inputs**")
    st.dataframe(pd.DataFrame([asdict(record) for record in records]), hide_index=True, width="stretch")
    active = [record.name for record in records if record.override_active]
    st.write("Active manual overrides:", ", ".join(active) if active else "None")
    st.caption("Reference dates describe the underlying data, not the retrieval. An omitted observation date is reported as unknown. "
               "Refresh failures retain a session's last successful values only with an explicit STALE label. "
               "No persistent or bundled market-data snapshot is supplied.")
    st.markdown(f"[OECD dataset and explanatory material]({OECD_PAGE}) · [Public government bonds listing]({BONDS_URL})")


def compact_override(name: str, key: str, observation: Observation | None, state: SourceState,
                     checkbox_label: str, input_label: str, *, percentage: bool = True,
                     minimum: float = 0.0, maximum: float | None = None,
                     show_disabled_input: bool = False, source_caption: str | None = None,
                     show_used: bool = True) -> EffectiveInput:
    """A compact checkbox override with saved values independent of display widgets."""
    auto = observation.value if observation else None
    scale = 100.0 if percentage else 1.0
    seed = auto * scale if auto is not None else None
    if seed is not None and not math.isfinite(seed):
        seed = None
    store = st.session_state.setdefault("manual_store", {})
    saved = store.setdefault(key, {"enabled": False, "value": seed})
    toggle_key, value_key = "override_toggle:" + key, "override_value:" + key
    if toggle_key not in st.session_state:
        st.session_state[toggle_key] = saved["enabled"]
    enabled = st.checkbox(checkbox_label, key=toggle_key)
    saved["enabled"] = enabled
    if source_caption is not None:
        st.caption(source_caption)
    elif auto is not None:
        label = percent(auto) if percentage else f"{auto:.8f}"
        st.caption(f"Automatic: {label} · {observation.reference_date}")
    else:
        st.caption("Automatic source unavailable. Enable the override to enter a verified value.")
    entered = saved["value"]
    if enabled:
        if value_key not in st.session_state:
            st.session_state[value_key] = entered
        entered = st.number_input(input_label, min_value=minimum, max_value=maximum,
                                  value=None, step=0.0001 if percentage else 0.01,
                                  format="%.6f", key=value_key,
                                  placeholder="Enter a verified value")
        saved["value"] = entered
    elif show_disabled_input:
        # The inactive field shows the current source quote. Its saved manual
        # entry lives separately and is restored when the checkbox is enabled.
        st.number_input(input_label, min_value=minimum, max_value=maximum,
                        value=seed, step=0.0001 if percentage else 0.01,
                        format="%.6f", disabled=True,
                        help="Enable the override to use a saved or newly entered manual value.")
    used = entered / scale if enabled and entered is not None else None if enabled else auto
    if show_used and (enabled or source_caption is None):
        label = percent(used) if percentage else "Unavailable" if used is None else f"{used:.8f}"
        st.caption(f"Used in calculations: {label}" + (" · manual override" if enabled else ""))
    if enabled and entered is None:
        st.warning("Enter a value for this enabled override, or uncheck it to use the source.")
    notes = " | ".join(filter(None, [observation.note if observation else "No matching source observation.",
                                    state.dataset.note if state.dataset else "", state.error]))
    return EffectiveInput(name, used, auto, enabled, state.source, state.url,
                          observation.reference_date if observation else "Unavailable",
                          state.dataset.retrieved_at if state.dataset else "",
                          state.status if observation else "Unavailable", notes)


def linked_country_yield(us_yield: EffectiveInput, state: SourceState) -> EffectiveInput:
    observation = lookup(state, "yield")
    return EffectiveInput("Foreign 10-year yield", us_yield.value,
                          observation.value if observation else None, False,
                          state.source, state.url,
                          observation.reference_date if observation else "Unavailable",
                          state.dataset.retrieved_at if state.dataset else "",
                          state.status if observation else "Unavailable",
                          "Linked to the effective U.S. yield, including its override. "
                          f"U.S. yield override active: {us_yield.override_active}.")


def yield_source_caption(observation: Observation | None, state: SourceState) -> str:
    if observation is None:
        return f"Source quote unavailable. Enter a verified 10-year yield manually. [Trading Economics]({state.url})"
    status = f" · {state.status}" if state.status.startswith("STALE") else ""
    return f"Source quote: {percent(observation.value)} · {observation.reference_date}{status} · [Trading Economics]({state.url})"


def render_detailed_calculation(p: Inputs, r: Results, country: str) -> None:
    currency = COUNTRIES[country]["currency"]
    rows = [
        ("Government yield", percent(p.rf_us), percent(p.rf_foreign)),
        ("Unlevered beta", f"{p.beta_u:.8f}", f"{p.beta_u:.8f}"),
        ("Relevered equity beta", f"{r.beta_e:.8f}", f"{r.beta_e:.8f}"),
        ("Debt-to-equity ratio", f"{r.debt_to_equity:.8f}", f"{r.debt_to_equity:.8f}"),
        ("Equity risk premium", percent(p.erp), percent(p.erp)),
        ("Project corporate tax rate", percent(p.tax_rate), percent(p.tax_rate)),
        ("Debt financing share", percent(p.debt_share), percent(p.debt_share)),
        ("Borrowing spread", percent(p.spread_us), percent(r.spread_foreign)),
        ("Cost of equity", percent(r.re_us), percent(r.re_foreign)),
        ("Pretax cost of debt", percent(r.rd_us), percent(r.rd_foreign)),
        ("After-tax cost of debt", percent(r.rd_aftertax_us), percent(r.rd_aftertax_foreign)),
        ("Equity contribution to WACC", percent(r.equity_contribution_us), percent(r.equity_contribution_foreign)),
        ("Debt contribution to WACC", percent(r.debt_contribution_us), percent(r.debt_contribution_foreign)),
        ("WACC", percent(r.wacc_us), percent(r.wacc_foreign)),
    ]
    st.dataframe(pd.DataFrame(rows, columns=["Metric", "USD financing", f"{currency} financing · {country}"]),
                 hide_index=True, width="stretch")
    st.caption(f"Currency factor F = {r.factor:.8f}. The same project tax rate, beta, ERP, and debt share apply to both scenarios.")
    chart = go.Figure()
    labels = ["Cost of equity", "After-tax debt", "WACC"]
    for label, values, color in [
        ("USD financing", [r.re_us, r.rd_aftertax_us, r.wacc_us], "#353b43"),
        (f"{currency} financing · {country}", [r.re_foreign, r.rd_aftertax_foreign, r.wacc_foreign], "#b65b5b"),
    ]:
        chart.add_bar(name=label, x=labels, y=values, marker_color=color,
                      hovertemplate="%{x}: %{y:.4%}<extra>%{fullData.name}</extra>")
    chart.update_layout(barmode="group", height=330, margin=dict(l=5, r=5, t=35, b=0),
                        yaxis_tickformat=".2%", legend=dict(orientation="h", y=1.16))
    st.plotly_chart(chart, width="stretch", key="financing_comparison_chart")
    if st.checkbox("Show WACC contribution breakdown", key="show_wacc_contributions"):
        breakdown = go.Figure()
        scenarios = ["USD financing", f"{currency} financing · {country}"]
        breakdown.add_bar(name="Equity contribution", x=scenarios,
                          y=[r.equity_contribution_us, r.equity_contribution_foreign], marker_color="#353b43")
        breakdown.add_bar(name="After-tax debt contribution", x=scenarios,
                          y=[r.debt_contribution_us, r.debt_contribution_foreign], marker_color="#b65b5b")
        breakdown.update_layout(barmode="relative", height=290, yaxis_tickformat=".2%",
                                margin=dict(l=5, r=5, t=20, b=0))
        st.plotly_chart(breakdown, width="stretch", key="wacc_contributions_chart")


def render_verification_tab(p: Inputs | None, industry_frame: pd.DataFrame | None) -> None:
    """Pure, offline finance checks; never run HTTP-patching unittest fixtures here."""
    st.write("The professor's Colombia example is independent of live market observations.")
    reference = Inputs(0.80, 0.30, 0.35, 0.065, 0.04, 0.12, 0.015)
    actual = calculate(reference)
    expected = {
        "debt_to_equity": ("D/E", 3 / 7, False),
        "beta_e": ("Relevered equity beta", 179 / 175, False),
        "factor": ("Currency factor F", 14 / 13, False),
        "re_us": ("USD cost of equity", 0.10648571428571429, True),
        "rd_us": ("USD pretax debt cost", 0.055, True),
        "rd_aftertax_us": ("USD after-tax debt cost", 0.03575, True),
        "wacc_us": ("USD WACC", 0.085265, True),
        "re_foreign": ("COP cost of equity", 0.1916, True),
        "spread_foreign": ("COP borrowing spread", 0.016153846153846154, True),
        "rd_foreign": ("COP pretax debt cost", 0.13615384615384615, True),
        "rd_aftertax_foreign": ("COP after-tax debt cost", 0.0885, True),
        "wacc_foreign": ("COP WACC", 0.16067, True),
    }
    rows = []
    for field, (label, value, is_rate) in expected.items():
        computed = getattr(actual, field)
        rows.append({"Metric": label,
                     "Expected": percent(value) if is_rate else f"{value:.8f}",
                     "Calculated": percent(computed) if is_rate else f"{computed:.8f}",
                     "Passed": math.isclose(computed, value, rel_tol=1e-11, abs_tol=1e-12)})
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    equal = calculate(replace(reference, rf_foreign=reference.rf_us))
    zero = calculate(replace(reference, debt_share=0.0))
    overridden = calculate(replace(reference, spread_foreign_override=0.04))
    low_tax = calculate(replace(reference, tax_rate=0.25))
    negative = calculate(replace(reference, rf_us=-0.005, rf_foreign=-0.01))
    checks = [
        ("Equal government yields produce equal scenario WACCs", math.isclose(equal.wacc_us, equal.wacc_foreign, abs_tol=1e-12)),
        ("Zero debt makes WACC equal the corresponding equity cost", zero.wacc_us == zero.re_us and zero.wacc_foreign == zero.re_foreign),
        ("Foreign spread override preserves USD equity/debt/WACC", all(getattr(actual, k) == getattr(overridden, k) for k in ("re_us", "rd_us", "wacc_us"))),
        ("Project tax changes both scenario WACCs", low_tax.wacc_us != actual.wacc_us and low_tax.wacc_foreign != actual.wacc_foreign),
        ("United States comparison links the effective U.S. yield", comparison_yield("United States", 0.04, 0.12) == 0.04),
        ("Valid negative government yields produce finite results", all(math.isfinite(v) for v in asdict(negative).values())),
    ]
    if p is not None and industry_frame is not None:
        selected_rows = industry_frame[industry_frame["Selected"]]
        if not selected_rows.empty:
            selected = selected_rows.iloc[0]
            current = calculate(p)
            checks.append(("Current selected industry agrees with the project calculation",
                           math.isclose(selected["USD WACC (decimal)"], current.wacc_us, abs_tol=1e-12)
                           and math.isclose(selected["Foreign WACC (decimal)"], current.wacc_foreign, abs_tol=1e-12)))
    passed = all(row["Passed"] for row in rows) and all(result for _, result in checks)
    if passed:
        st.success("Colombia reference and mathematical consistency checks passed.")
    else:
        st.error("A financial verification failed. Review the rows below before using the results.")
    st.dataframe(pd.DataFrame(checks, columns=["Consistency check", "Passed"]), hide_index=True, width="stretch")
    st.caption("These live-interface checks use only pure calculations. The full embedded test suite also exercises parsing, exports, and source-failure handling.")
    st.code("python app.py --self-test", language="bash")


def main() -> None:
    st.set_page_config(page_title="Cost of capital", layout="wide", initial_sidebar_state="expanded")
    st.markdown("""<style>
    .stApp, [data-testid="stHeader"] {background:#ffffff;color:#262730;}
    .block-container {padding-top:3.4rem;padding-bottom:2.5rem;max-width:1230px;}
    [data-testid="stSidebar"] {background:#f7f7f8;border-right:1px solid #e5e5e7;}
    [data-testid="stSidebar"] .stVerticalBlock {gap:.6rem;}
    [data-testid="stSidebar"] h3 {font-size:1.35rem;line-height:1.5;}
    [data-testid="stSidebar"] label p {font-size:.88rem;}
    h1 {font-size:2.7rem;letter-spacing:-.035em;font-weight:750;}
    h2 {font-size:1.8rem;letter-spacing:-.025em;}
    h3 {font-size:1.5rem;letter-spacing:-.02em;}
    [data-testid="stMetric"] {border:0;background:transparent;padding:0;}
    [data-testid="stMetricValue"] {font-size:1.9rem;color:#262730;}
    [data-testid="stMetricLabel"] p {font-size:.85rem;color:#52525b;}
    .st-key-primary_results [data-testid="stMetricValue"] {font-size:2.65rem;}
    [data-testid="stSlider"] [role="slider"] {background-color:#e25454;}
    hr {border-color:#e4e4e7;}
    @media(max-width:700px) {
      .block-container {padding-top:3rem;padding-left:1rem;padding-right:1rem;}
      h1 {font-size:2.1rem;} [data-testid="stMetricValue"] {font-size:1.55rem;}
      .st-key-primary_results [data-testid="stMetricValue"] {font-size:2rem;}
    }
    </style>""", unsafe_allow_html=True)
    st.title("Cost of capital")
    st.caption("U.S. industry betas, corporate taxes, BBB borrowing spreads, and 10-year government bond yields")

    country = st.session_state.get("country") or "United States"
    if country not in COUNTRIES:
        country = "United States"
        st.session_state["country"] = country
    with st.sidebar:
        st.subheader("Project assumptions")
        refresh = st.button("Refresh source data", width="stretch", key="refresh_source_data")
    try:
        ttl = int(os.getenv("COC_CACHE_TTL_SECONDS", "3600"))
        if ttl < 1:
            raise ValueError("TTL must be positive")
    except ValueError:
        ttl = 3600
        st.warning("Invalid COC_CACHE_TTL_SECONDS; using a 3,600-second cache lifetime.")
    cached_fetch = st.cache_data(ttl=ttl, show_spinner=False)(retrieve_for_revision)
    if refresh:
        cached_fetch.clear()
    with st.spinner("Retrieving source data…"):
        bundle = cached_fetch(country, APP_REVISION)
    states = reconcile_sources(bundle, st.session_state.setdefault("last_good_sources", {}), ttl)
    beta_state, tax_state, spread_state = states["betas"], states["taxes"], states["spread"]
    us_yield_state = states["yield:USA"]
    foreign_yield_state = states[f"yield:{COUNTRIES[country]['iso']}"]
    betas = beta_state.dataset.observations if beta_state.dataset else {}
    prior_industry = st.session_state.get("industry") or st.session_state.get("last_selected_industry", DEFAULT_INDUSTRY)
    industries = sorted(set(betas) | {prior_industry})
    with st.sidebar:
        if "industry" in st.session_state and st.session_state["industry"] is None:
            del st.session_state["industry"]
        industry = st.selectbox("U.S. industry", industries, index=industries.index(prior_industry), key="industry")
        st.session_state["last_selected_industry"] = industry
        st.selectbox("Unlevered beta measure", ["Unlevered beta"], key="beta_measure",
                     help="Uses the exact published Unlevered beta column. No cash-corrected or levered beta is substituted.")
        erp = st.number_input("Annual equity risk premium (%)", min_value=0.0, value=6.5,
                              step=0.05, format="%.2f", key="erp") / 100
        # Keep older sessions usable when migrating from the previous 99.99% input.
        if st.session_state.get("debt") is not None and not 0 <= st.session_state["debt"] <= 99:
            st.session_state["debt"] = min(99.0, max(0.0, float(st.session_state["debt"])))
        debt = st.slider("Fraction financed with debt (%)", min_value=0.0, max_value=99.0,
                         value=30.0, step=0.1, format="%.2f%%", key="debt") / 100
        automatic_spread = lookup(spread_state, "BBB")
        spread_caption = (f"BBB spread: {percent(automatic_spread.value, 2)} "
                          f"({automatic_spread.value * 10000:.0f} basis points) · {automatic_spread.reference_date}"
                          if automatic_spread else "BBB spread unavailable. Enter a verified borrowing spread below.")
        if spread_state.status.startswith("STALE"):
            spread_caption += " · " + spread_state.status
        us_spread = compact_override("U.S. BBB spread", "spread:USA", automatic_spread, spread_state,
                                      "Override BBB borrowing spread",
                                      "USD borrowing spread over U.S. 10-year yield (percentage points)",
                                      show_disabled_input=True, source_caption=spread_caption)
        country = st.selectbox("Project country / currency", sorted(COUNTRIES),
                               index=sorted(COUNTRIES).index(country), key="country")
        st.caption(f"Currency: {COUNTRIES[country]['currency']}")
        project_tax_slot = st.container()
        beta_observation = lookup(beta_state, industry)
        if beta_observation is None:
            st.warning("The selected industry's source beta is unavailable. Its selection is preserved; enter a beta in Advanced assumptions and manual overrides.")
        advanced_missing = beta_observation is None
        with st.expander("Advanced assumptions and manual overrides", expanded=advanced_missing):
            beta = compact_override("Unlevered beta", f"beta:{industry}", beta_observation, beta_state,
                                    "Override industry unlevered beta", "Industry unlevered beta", percentage=False)
            st.caption("U.S. and selected-country yield overrides are directly available in the 10-year yields section on the main page.")
            us_tax = compact_override("U.S. corporate tax rate", "tax:USA", lookup(tax_state, "USA"), tax_state,
                                      "Override U.S. corporate tax data", "U.S. corporate tax data (%)", maximum=100.0)
            foreign_tax = compact_override("Foreign corporate tax rate", f"foreign_tax:{COUNTRIES[country]['iso']}",
                                           lookup(tax_state, COUNTRIES[country]["iso"]), tax_state,
                                           "Override selected-country corporate tax data", "Selected-country corporate tax data (%)", maximum=100.0)
            foreign_spread_slot = st.container()
            tax_selection = st.radio("Corporate tax jurisdiction",
                                     ["U.S. statutory corporate tax rate", "Selected country's statutory corporate tax rate"],
                                     key="tax_selection")
            st.button("Reset all manual overrides", on_click=reset_manual_inputs, width="stretch", key="reset_manual_overrides")
        selected_source_tax = us_tax if tax_selection.startswith("U.S.") else foreign_tax
        tax_jurisdiction = "United States" if tax_selection.startswith("U.S.") else country
        source_tax_observation = lookup(tax_state, "USA" if tax_selection.startswith("U.S.") else COUNTRIES[country]["iso"])
        effective_tax_observation = (Observation(selected_source_tax.value, selected_source_tax.reference_date,
                                                "Inherited from the selected effective statutory-tax source. "
                                                f"Source tax override active: {selected_source_tax.override_active}.")
                                     if selected_source_tax.value is not None else None)
        raw_tax_caption = (f"OECD combined statutory rate: {percent(source_tax_observation.value, 2)} · {source_tax_observation.reference_date}"
                           if source_tax_observation else "OECD combined statutory rate unavailable. Enter the project rate manually.")
        if tax_state.status.startswith("STALE"):
            raw_tax_caption += " · " + tax_state.status
        with project_tax_slot:
            project_tax = compact_override("Project corporate tax rate", "tax:project", effective_tax_observation, tax_state,
                                           "Override corporate tax rate", "Project corporate tax rate (%)", maximum=100.0,
                                           source_caption=raw_tax_caption)
            if selected_source_tax.override_active and not project_tax.override_active:
                st.caption(f"Using {tax_jurisdiction} source-tax override: {percent(project_tax.value)}.")
        st.caption("One project tax rate is applied to both financing scenarios.")

    st.subheader("10-year yields used in the calculation")
    yield_left, yield_right = st.columns(2, gap="large")
    with yield_left:
        us_yield = compact_override("U.S. 10-year yield", "yield:USA", lookup(us_yield_state, "yield"), us_yield_state,
                                     "Enter U.S. 10-year yield manually", "U.S. 10-year government yield (%)",
                                     minimum=-99.9999,
                                     source_caption=yield_source_caption(lookup(us_yield_state, "yield"), us_yield_state))
    with yield_right:
        st.write(f"Selected country: {country} ({COUNTRIES[country]['currency']})")
        if country == "United States":
            foreign_yield = linked_country_yield(us_yield, foreign_yield_state)
            st.caption("The selected-country yield equals the U.S. yield. No currency conversion is needed.")
        else:
            foreign_yield = compact_override("Foreign 10-year yield", f"foreign_yield:{COUNTRIES[country]['iso']}",
                                             lookup(foreign_yield_state, "yield"), foreign_yield_state,
                                             "Enter selected-country 10-year yield manually", "Selected-country 10-year government yield (%)",
                                             minimum=-99.9999,
                                             source_caption=yield_source_caption(lookup(foreign_yield_state, "yield"), foreign_yield_state))
    effective_foreign_yield = comparison_yield(country, us_yield.value, foreign_yield.value)
    auto_foreign = None
    if us_spread.value is not None and us_yield.value is not None and effective_foreign_yield is not None:
        auto_foreign = us_spread.value * ((1 + effective_foreign_yield) / (1 + us_yield.value))
        if not math.isfinite(auto_foreign):
            auto_foreign = None
    derived = Dataset("Model-derived foreign BBB spread", "", utc_now(), {},
                      "s_FOREIGN = effective s_US × (1 + effective Rf_FOREIGN) / (1 + effective Rf_US).")
    derived_state = SourceState(derived.source, "", derived, "Derived from current inputs", derived.retrieved_at)
    with foreign_spread_slot:
        foreign_spread = compact_override("Foreign BBB spread", f"foreign_spread:{COUNTRIES[country]['iso']}",
                                          Observation(auto_foreign, "Derived, not a market quote") if auto_foreign is not None else None,
                                          derived_state, "Override foreign-currency BBB borrowing spread",
                                          "Foreign borrowing spread (percentage points)")
    records = [beta, us_tax, foreign_tax, project_tax, us_yield, foreign_yield, us_spread, foreign_spread]
    required = {"industry unlevered beta (advanced assumptions)": beta.value,
                "project corporate tax rate (sidebar)": project_tax.value,
                "U.S. 10-year yield (main page)": us_yield.value,
                "USD borrowing spread (sidebar)": us_spread.value}
    if country != "United States":
        required["selected-country 10-year yield (main page)"] = effective_foreign_yield
    if foreign_spread.override_active:
        required["foreign BBB spread override (advanced assumptions)"] = foreign_spread.value
    missing = [name for name, value in required.items() if value is None]
    p, result = None, None
    if not missing:
        try:
            p = Inputs(beta.value, debt, project_tax.value, erp, us_yield.value,
                       effective_foreign_yield, us_spread.value,
                       foreign_spread.value if foreign_spread.override_active else None)
            result = calculate(p)
        except InputError as exc:
            p = None
            st.error(str(exc))

    st.subheader("Borrowing spreads")
    currency = COUNTRIES[country]["currency"]
    spread_columns = st.columns(2) if country != "United States" else [st.container()]
    with spread_columns[0]:
        st.metric("Borrowing spread · USD", percent(us_spread.value))
        if us_spread.value is not None:
            st.caption(f"{us_spread.value * 10000:.2f} basis points over the U.S. 10-year yield")
    if country != "United States":
        with spread_columns[1]:
            foreign_spread_label = currency if currency != "USD" else f"USD · {country}"
            st.metric(f"Borrowing spread · {foreign_spread_label}", percent(foreign_spread.value))
            st.caption("Manual foreign spread override" if foreign_spread.override_active else "Automatically adjusted using the currency factor F")
    st.header(industry)
    firms = getattr(beta_observation, "firm_count", None)
    firms_text = f"{firms:,} U.S. firms." if firms is not None else "U.S. firm count unavailable."
    beta_date = beta_observation.reference_date if beta_observation else "unavailable"
    override_text = " · manual project tax override" if project_tax.override_active else ""
    st.caption(f"{firms_text} Beta data dated {beta_date}. Tax jurisdiction: {tax_jurisdiction}{override_text}; "
               "the same tax rate is used in both currency scenarios.")
    assumption_metrics = st.columns(4)
    assumption_metrics[0].metric("U.S. 10-year yield", percent(us_yield.value, 3))
    assumption_metrics[1].metric("Selected country 10-year yield", percent(effective_foreign_yield, 3))
    assumption_metrics[2].metric("Corporate tax rate", percent(project_tax.value, 2))
    assumption_metrics[3].metric("Unlevered beta", "Unavailable" if beta.value is None else f"{beta.value:.4f}")
    st.divider()
    with st.container(key="primary_results"):
        primary = st.columns(2)
        primary[0].metric("Equity cost of capital · USD", percent(result.re_us, 2) if result else "Unavailable")
        primary[1].metric("WACC · USD", percent(result.wacc_us, 2) if result else "Unavailable")
        changed_us_comparison = (result is not None and country == "United States"
                                 and not math.isclose(result.wacc_us, result.wacc_foreign, rel_tol=1e-12, abs_tol=1e-12))
        if country != "United States" or changed_us_comparison:
            alternate = st.columns(2)
            scenario_label = (currency if currency != "USD" else f"USD · {country}") if country != "United States" else "USD · selected scenario"
            alternate[0].metric(f"Equity cost of capital · {scenario_label}", percent(result.re_foreign, 2) if result else "Unavailable")
            alternate[1].metric(f"WACC · {scenario_label}", percent(result.wacc_foreign, 2) if result else "Unavailable")
            if changed_us_comparison:
                st.caption("The selected USD scenario differs because its foreign-spread override is active; the currency factor remains 1.")
    st.caption("Nominal required returns depend on currency. Project cash flows and discount rates must use a consistent currency.")
    if country != "United States" and currency == "USD":
        st.info(f"{country} also uses USD. Its government yield can differ because of sovereign and liquidity risk. "
                "The scenario difference here reflects the country-yield proxy; it is not a currency conversion or evidence of cheaper financing.")
    if missing:
        st.warning("Complete the following inputs to calculate results: " + "; ".join(missing) + ".")
    source_issues = [state for state in states.values() if state.error or state.status.startswith("STALE")]
    if refresh:
        failures = sum(isinstance(item, FetchFailure) for item in bundle.values())
        if failures:
            st.warning(f"Source refresh completed with {failures} error(s). Saved manual inputs are preserved; see Data sources.")
        else:
            st.success("Source data refreshed. Project selections and manual values are preserved.")
    if source_issues:
        with st.expander(f"Source status · {len(source_issues)} source issue(s)"):
            for state in source_issues:
                st.warning(f"{state.source}: {state.status}. {state.error}")
    industry_frame, industry_error = None, None
    if p is not None and betas:
        try:
            industry_frame = industry_comparison(betas, p, industry, beta.override_active)
        except InputError as exc:
            industry_error = str(exc)
    generated_at = utc_now()
    detailed_tab, industries_tab, methodology_tab, sources_tab, verification_tab, downloads_tab = st.tabs(
        ["Detailed calculation", "All industries", "Methodology and assumptions", "Data sources", "Verification", "Downloads"])
    with detailed_tab:
        if p and result:
            render_detailed_calculation(p, result, country)
        else:
            st.info("The detailed calculation appears when the required source data or manual inputs are available.")
    with industries_tab:
        if not betas:
            st.warning("The industry dataset is unavailable. Manual project inputs still work; all-industry comparison requires the source dataset.")
        elif industry_error:
            st.error(f"The all-industry comparison cannot calculate every industry with these assumptions: {industry_error}")
        elif industry_frame is not None:
            if beta.override_active:
                st.info("Only the selected industry uses its manual beta. Other industries retain their individual source betas.")
            if industry not in betas:
                st.info("The selected project is included as an additional manual-beta row; its source beta is unavailable.")
            search = st.text_input("Search industries", placeholder="Type part of an industry name", key="industry_search")
            visible = industry_frame[industry_frame["Industry"].str.contains(search, case=False, regex=False)] if search else industry_frame
            st.caption(f"{len(betas)} source industries · Selected: {industry}. Select a column header to sort. Downloads include every row.")
            formats = {column: "{:.4%}" for column in industry_frame.columns if "(decimal)" in column}
            formats.update({"Source beta": "{:.8f}", "Beta used": "{:.8f}", "Equity beta": "{:.8f}", "U.S. firms": "{:.0f}"})
            st.dataframe(visible.style.format(formats, na_rep="Unavailable"), hide_index=True, width="stretch", height=500)
        else:
            st.info("Complete the required project inputs to compare all industries.")
    with methodology_tab:
        render_methodology(p, result)
        st.caption("Calculations use full numerical precision. The home screen rounds headline results for readability; detailed tables and CSV files retain greater precision.")
    with sources_tab:
        render_sources(states, records)
        st.caption(f"Cache lifetime: {ttl:,} seconds. The project tax override takes precedence over the selected jurisdiction's effective source rate.")
    with verification_tab:
        render_verification_tab(p, industry_frame)
    with downloads_tab:
        st.write("Downloads contain the current project assumptions, full-precision results, currency labels, active overrides, and source provenance. Rates are exported as decimals.")
        if p and result:
            st.download_button("Download selected-project CSV",
                               project_csv(p, result, records, industry, country, tax_selection, generated_at),
                               file_name="selected_project_cost_of_capital.csv", mime="text/csv", on_click="ignore")
        else:
            st.info("Complete the required project inputs to enable the selected-project download.")
        if p and industry_frame is not None:
            st.download_button("Download all-industry CSV",
                               all_industry_export(industry_frame, p, country, beta_state, records, generated_at),
                               file_name="all_industries_cost_of_capital.csv", mime="text/csv", on_click="ignore")
        elif industry_error:
            st.warning("The all-industry download is unavailable until the invalid comparison assumptions are corrected.")
    st.caption(f"Educational model · Application revision: {APP_REVISION}")


# ---------------------------------------------------------------------------
# Embedded verification: no server, network, or additional test file required.
# ---------------------------------------------------------------------------
REFERENCE_INPUTS = Inputs(0.80, 0.30, 0.35, 0.065, 0.04, 0.12, 0.015)


class FinancialVerification(unittest.TestCase):
    def test_colombia_reference(self):
        result = calculate(REFERENCE_INPUTS)
        expected = {
            "debt_to_equity": 3 / 7, "beta_e": 179 / 175, "factor": 14 / 13,
            "re_us": 0.10648571428571429, "rd_us": 0.055,
            "rd_aftertax_us": 0.03575, "wacc_us": 0.085265,
            "re_foreign": 0.1916, "spread_foreign": 0.016153846153846154,
            "rd_foreign": 0.13615384615384615, "rd_aftertax_foreign": 0.0885,
            "wacc_foreign": 0.16067,
        }
        for field, value in expected.items():
            with self.subTest(metric=field):
                self.assertTrue(math.isclose(getattr(result, field), value, rel_tol=1e-11, abs_tol=1e-12))

    def test_a_equal_yields(self):
        r = calculate(replace(REFERENCE_INPUTS, rf_foreign=REFERENCE_INPUTS.rf_us))
        self.assertEqual(r.factor, 1)
        for us, foreign in ((r.re_us, r.re_foreign), (r.rd_us, r.rd_foreign),
                            (r.rd_aftertax_us, r.rd_aftertax_foreign), (r.wacc_us, r.wacc_foreign)):
            self.assertAlmostEqual(us, foreign, places=14)

    def test_b_zero_debt(self):
        p = replace(REFERENCE_INPUTS, debt_share=0)
        r = calculate(p)
        self.assertEqual(r.debt_to_equity, 0)
        self.assertEqual(r.beta_e, p.beta_u)
        self.assertEqual(r.wacc_us, r.re_us)
        self.assertEqual(r.wacc_foreign, r.re_foreign)

    def test_c_foreign_spread_override(self):
        base = calculate(REFERENCE_INPUTS)
        changed = calculate(replace(REFERENCE_INPUTS, spread_foreign_override=0.04))
        for field in ("re_us", "rd_us", "rd_aftertax_us", "wacc_us", "re_foreign"):
            self.assertEqual(getattr(base, field), getattr(changed, field))
        for field in ("rd_foreign", "rd_aftertax_foreign", "wacc_foreign"):
            self.assertNotEqual(getattr(base, field), getattr(changed, field))
        expected_delta = REFERENCE_INPUTS.debt_share * (0.04 - base.spread_foreign) * (1 - REFERENCE_INPUTS.tax_rate)
        self.assertAlmostEqual(changed.wacc_foreign - base.wacc_foreign, expected_delta, places=14)
        zero = calculate(replace(REFERENCE_INPUTS, spread_foreign_override=0.0))
        self.assertEqual(zero.rd_foreign, REFERENCE_INPUTS.rf_foreign)

    def test_d_tax_selection(self):
        for tax in (0.255, 0.35):
            with self.subTest(tax=tax):
                p = replace(REFERENCE_INPUTS, tax_rate=tax)
                r = calculate(p)
                expected_beta = p.beta_u * (1 + (1 - tax) * p.debt_share / (1 - p.debt_share))
                self.assertAlmostEqual(r.beta_e, expected_beta)
                self.assertAlmostEqual(r.re_us, p.rf_us + expected_beta * p.erp)
                self.assertAlmostEqual(r.re_foreign, (1 + r.re_us) * r.factor - 1)
                self.assertAlmostEqual(r.rd_aftertax_us, r.rd_us * (1 - tax))
                self.assertAlmostEqual(r.rd_aftertax_foreign, r.rd_foreign * (1 - tax))
                self.assertAlmostEqual(r.wacc_us, (1 - p.debt_share) * r.re_us + p.debt_share * r.rd_us * (1 - tax))
                self.assertAlmostEqual(r.wacc_foreign, (1 - p.debt_share) * r.re_foreign + p.debt_share * r.rd_foreign * (1 - tax))
        lower, higher = (calculate(replace(REFERENCE_INPUTS, tax_rate=t)) for t in (0.255, 0.35))
        for field in ("beta_e", "re_us", "re_foreign", "rd_aftertax_us", "rd_aftertax_foreign", "wacc_us", "wacc_foreign"):
            self.assertNotEqual(getattr(lower, field), getattr(higher, field))

    def test_e_united_states_link(self):
        us = 0.0475  # Emulates an active manual U.S. yield override.
        foreign = comparison_yield("United States", us, 0.12)
        self.assertEqual(foreign, us)
        r = calculate(replace(REFERENCE_INPUTS, rf_us=us, rf_foreign=foreign))
        self.assertAlmostEqual(r.wacc_us, r.wacc_foreign, places=14)
        self.assertAlmostEqual(r.re_us, r.re_foreign, places=14)
        self.assertAlmostEqual(r.rd_us, r.rd_foreign, places=14)
        self.assertAlmostEqual(r.rd_aftertax_us, r.rd_aftertax_foreign, places=14)
        self.assertEqual(comparison_yield("Colombia", us, 0.12), 0.12)

    def test_f_all_industry_consistency(self):
        betas = {"Selected": Observation(0.8, "2026"), "Other": Observation(1.2, "2026"),
                 "Total Market": Observation(99, "2026")}
        for overridden in (False, True):
            p = replace(REFERENCE_INPUTS, beta_u=0.9 if overridden else 0.8)
            frame = industry_comparison(betas, p, "Selected", overridden).set_index("Industry")
            r = calculate(p)
            self.assertEqual(len(frame), 2)
            self.assertEqual(frame.loc["Selected", "USD WACC (decimal)"], r.wacc_us)
            self.assertEqual(frame.loc["Selected", "Foreign WACC (decimal)"], r.wacc_foreign)
            self.assertEqual(frame.loc["Other", "Beta used"], 1.2)
            self.assertEqual(bool(frame.loc["Selected", "Beta overridden"]), overridden)

    def test_g_valid_boundaries(self):
        for changed in ({"debt_share": 0.9999}, {"erp": 0}, {"rf_us": -0.005, "rf_foreign": -0.01},
                        {"tax_rate": 0}, {"tax_rate": 1}, {"beta_u": 0}, {"spread_us": 0}):
            with self.subTest(changed=changed):
                r = calculate(replace(REFERENCE_INPUTS, **changed))
                self.assertTrue(all(math.isfinite(v) for v in asdict(r).values()))
        zero_erp = calculate(replace(REFERENCE_INPUTS, erp=0))
        self.assertEqual(zero_erp.re_us, REFERENCE_INPUTS.rf_us)
        self.assertAlmostEqual(zero_erp.re_foreign, REFERENCE_INPUTS.rf_foreign)

    def test_g_invalid_and_missing_inputs(self):
        invalid = [("debt_share", 1), ("debt_share", -0.1), ("tax_rate", 1.01), ("tax_rate", -0.01),
                   ("rf_us", -1), ("rf_foreign", -1.1), ("beta_u", -0.1), ("erp", -0.01),
                   ("spread_us", -0.01), ("spread_foreign_override", -0.01),
                   ("rf_us", None), ("beta_u", float("nan")), ("erp", float("inf")), ("tax_rate", "35%")]
        for field, value in invalid:
            with self.subTest(field=field, value=value), self.assertRaises(InputError):
                calculate(replace(REFERENCE_INPUTS, **{field: value}))

    def test_beta_column_identity(self):
        html = """<p>Date of Analysis: Data used is as of January 2026</p><table>
        <tr><th>Industry Name</th><th>Beta</th><th>Unlevered beta</th><th>Unlevered beta corrected for cash</th></tr>
        <tr><td>Software</td><td>9.9</td><td>0.8</td><td>1.2</td></tr>
        <tr><td>Total Market</td><td>1</td><td>1</td><td>1</td></tr></table>"""
        values, date = parse_beta_html(html)
        self.assertEqual(values, {"Software": 0.8})
        self.assertEqual(date, "January 2026")

    def test_industry_firm_count_metadata(self):
        html = """<p>Date of Analysis: January 2026</p><table>
        <tr><th>Industry Name</th><th>Number of firms</th><th>Beta</th><th>Unlevered beta</th></tr>
        <tr><td>Software</td><td>309</td><td>9.9</td><td>0.8</td></tr>
        <tr><td>Other</td><td>N/A</td><td>2</td><td>1.2</td></tr></table>"""
        counts = {}
        values, _ = parse_beta_html(html, counts)
        self.assertEqual(counts, {"Software": 309})
        self.assertEqual(values, {"Software": 0.8, "Other": 1.2})
        frame = industry_comparison({"Software": Observation(0.8, "2026", firm_count=309)},
                                    REFERENCE_INPUTS, "Software", False)
        row = frame.iloc[0]
        result = calculate(REFERENCE_INPUTS)
        self.assertEqual(row["U.S. firms"], 309)
        self.assertEqual(row["USD pretax debt (decimal)"], result.rd_us)
        self.assertEqual(row["Foreign pretax debt (decimal)"], result.rd_foreign)
        self.assertEqual(row["USD after-tax debt (decimal)"], result.rd_aftertax_us)
        self.assertEqual(row["Foreign after-tax debt (decimal)"], result.rd_aftertax_foreign)

    def test_excel_html_inline_fragments_preserve_industry_identity(self):
        html = """<table><tr><td>Industry Name</td><td>Number of firms</td><td>Unlevered<br>beta</td></tr>
        <tr><td>Financial Svcs. (Non-bank &amp; Insur<span>ance)</span></td><td>42</td><td>0.8</td></tr></table>"""
        counts = {}
        values, _ = parse_beta_html(html, counts)
        self.assertEqual(values, {"Financial Svcs. (Non-bank & Insurance)": 0.8})
        self.assertEqual(counts, {"Financial Svcs. (Non-bank & Insurance)": 42})

    def test_bbb_group_and_units(self):
        html = """<p>Date of Analysis: Data used is as of January 2026</p><table>
        <tr><td colspan='3'>For large non-financial service firms</td><td></td><td></td>
        <td colspan='4'>For financial service firms</td></tr>
        <tr><td>&gt;</td><td>to</td><td>Rating is</td><td>Spread is</td><td></td>
        <td>&gt;</td><td>to</td><td>Rating is</td><td>Spread is</td></tr>
        <tr><td>2.5</td><td>2.9999</td><td>Baa2/BBB</td><td>1.5%</td><td></td>
        <td>0.9</td><td>1.2</td><td>Baa2/BBB</td><td>8.8%</td></tr></table>"""
        spread, _ = parse_bbb_spread(html)
        self.assertEqual(spread, 0.015)

    def test_oecd_statutory_combined_latest(self):
        header = "REF_AREA,FREQ,MEASURE,TARGETING,UNIT_MEASURE,SECTOR,DIV_TREATMENT,RATE_STRUCTURE,TAX_BASE,TIME_PERIOD,OBS_VALUE,UNIT_MULT\n"
        rows = ("USA,A,CIT_C,ST,PT_INC_TAX,S13,_Z,_Z,_Z,2025,25.57,0\n"
                "USA,A,CIT_C,ST,PT_INC_TAX,S13,_Z,_Z,_Z,2026,25.53,0\n"
                "USA,A,CIT,ST,PT_INC_TAX,S1311,_Z,_Z,_Z,2026,21,0\n"
                "COL,A,CIT_C,TA,PT_INC_TAX,S13,_Z,_Z,_Z,2026,9,0\n")
        values = parse_oecd_csv(header + rows, 2026)
        self.assertAlmostEqual(values["USA"].value, 0.2553)
        self.assertEqual(values["USA"].reference_date, "2026")
        self.assertNotIn("COL", values)
        with self.assertRaises(DataUnavailable):
            parse_oecd_csv("wrong,column\n1,2", 2026)

    def test_oecd_truncated_and_misaligned_rows(self):
        header = "REF_AREA,FREQ,MEASURE,TARGETING,UNIT_MEASURE,SECTOR,DIV_TREATMENT,RATE_STRUCTURE,TAX_BASE,TIME_PERIOD,OBS_VALUE,UNIT_MULT\n"
        prefix = "USA,A,CIT_C,ST,PT_INC_TAX,S13,_Z,_Z,_Z"
        for row in (prefix, prefix + ",2026,25", prefix + ",2026,25,0,unexpected"):
            with self.subTest(row=row), self.assertRaises(DataUnavailable):
                parse_oecd_csv(header + row, 2026)

    def test_oecd_oversized_error_page_is_recoverable(self):
        # A provider can return an HTML error document with HTTP 200.
        with self.assertRaises(DataUnavailable):
            parse_oecd_csv("<html>" + "x" * (csv.field_size_limit() + 1) + "</html>", 2026)

    def test_corrupt_xls_uses_official_html_fallback(self):
        # Exercise the provider boundary with actual exception types emitted by xlrd.
        html = "<p>Date of Analysis: January 2026</p><table><tr><th>Industry Name</th><th>Unlevered beta</th></tr>"
        html += "".join(f"<tr><td>Industry {i}</td><td>0.8</td></tr>" for i in range(20)) + "</table>"
        for failure in (struct.error("truncated XLS"), IndexError("missing sector"), EOFError("incomplete XLS")):
            responses = []
            for url, content in ((BETA_XLS, b"truncated workbook"), (BETA_HTML, html.encode())):
                response = requests.Response()
                response.status_code, response.url, response._content, response.encoding = 200, url, content, "utf-8"
                responses.append(response)
            with self.subTest(error=type(failure).__name__), \
                    patch(__name__ + ".http_get", side_effect=responses), \
                    patch.object(xlrd, "open_workbook", side_effect=failure):
                dataset = fetch_betas()
                self.assertEqual(dataset.url, BETA_HTML)
                self.assertEqual(len(dataset.observations), 20)
                self.assertIn("HTML fallback", dataset.note)

    def test_yield_country_and_maturity(self):
        html = """<title>Colombia 10-Year Government Bond Yield</title>
        <meta name="description" content="The yield on Colombia 10Y eased on November 13, 2026.">
        <script>TEChartsMeta = [{"name":"Colombia 10Y","last":12.698}];</script>"""
        quote, date, _ = parse_yield_html(html, "Colombia")
        self.assertAlmostEqual(quote, 0.12698)
        self.assertEqual(date, "2026-11-13")
        for wrong in ("United States", "Chile"):
            with self.assertRaises(DataUnavailable):
                parse_yield_html(html, wrong)
        with self.assertRaises(DataUnavailable):
            parse_yield_html(html.replace("10-Year", "5-Year").replace("10Y", "5Y"), "Colombia")

    def test_us_yield_short_country_name_metadata(self):
        # Synthetic quote; mirrors the public U.S. page's actual instrument names.
        html = """<title>US 10 Year Treasury Note Yield - Quote - Chart</title>
        <script>TEChartsMeta = [
        {"symbol":"USGG5YR:IND","name":"US 5Y","full_name":"US 5 Year Note Bond Yield","last":9.9},
        {"symbol":"USGG10YR:IND","name":"US 10Y","full_name":"US 10 Year Note Bond Yield","last":4.1234}
        ];</script>
        <script type="application/ld+json">{"@type":"Dataset",
        "name":"US 10 Year Treasury Note Yield Data",
        "temporalCoverage":"1912-01-01/2026-11-13"}</script>"""
        quote, date, _ = parse_yield_html(html, "United States")
        self.assertAlmostEqual(quote, 0.041234, places=14)
        self.assertEqual(date, "2026-11-13")
        with self.assertRaises(DataUnavailable):
            parse_yield_html(html, "United Kingdom")

    def test_us_yield_short_country_name_table(self):
        html = """<title>Bond Yields by Country</title>
        <meta name="description" content="Government Bond 10y">
        <table><tr><td><a href="/united-states/government-bond-yield">US 2Y</a></td>
        <td id="p">9.9</td></tr><tr>
        <td><a href="/united-states/government-bond-yield">US 10Y</a></td>
        <td id="p">4.12</td><td id="date">Nov/13</td></tr></table>"""
        quote, date, _ = parse_yield_html(html, "United States", board=True)
        self.assertEqual(quote, 0.0412)
        self.assertIn("source omits year", date)
        with self.assertRaises(DataUnavailable):
            parse_yield_html(html.replace("US 10Y", "US 5Y"), "United States", board=True)

    def test_uk_yield_short_country_name(self):
        html = """<title>UK 10 Year Bond Yield - Quote - Chart</title>
        <script>TEChartsMeta = [{"name":"UK 10Y","full_name":"UK 10 Year Bond Yield",
        "symbol":"GUKG10:IND","type":"Bond","last":4.25}];</script>"""
        quote, _, _ = parse_yield_html(html, "United Kingdom")
        self.assertAlmostEqual(quote, 0.0425)
        with self.assertRaises(DataUnavailable):
            parse_yield_html(html, "United States")

    def test_nominal_yield_rejects_tips_and_conflicting_symbols(self):
        for identity, symbol in (("US 10Y TIPS", "USGG10YR:IND"),
                                 ("US 10Y Inflation-Linked", "USGG10YR:IND"),
                                 ("US 10Y Real Yield", "USGG10YR:IND"),
                                 ("US 10Y", "USGG5YR:IND")):
            html = '<title>US 10 Year Treasury Note Yield</title><script>TEChartsMeta = ' + json.dumps(
                [{"full_name": identity, "symbol": symbol, "last":2.3}]) + ';</script>'
            with self.subTest(identity=identity, symbol=symbol), self.assertRaises(DataUnavailable):
                parse_yield_html(html, "United States")
        html = """<title>Government Bond 10y</title><meta name="description" content="Government Bond 10y">
        <table><tr data-symbol="USGG10YR:IND"><td>
        <a href="/united-states/government-bond-yield">US 10Y TIPS</a></td><td id="p">2.3</td></tr></table>"""
        with self.assertRaises(DataUnavailable):
            parse_yield_html(html, "United States", board=True)
        with self.assertRaises(DataUnavailable):
            parse_yield_html(html.replace("US 10Y TIPS", "US 10Y").replace("USGG10YR:IND", "USGG5YR:IND"),
                             "United States", board=True)

    def test_malformed_optional_metadata_keeps_valid_table_fallback(self):
        table = """<table><tr data-symbol="USGG10YR:IND"><td>
        <a href="/united-states/government-bond-yield">US 10Y</a></td>
        <td id="p">4.12</td><td id="date">Nov/13</td></tr></table>"""
        payloads = ['[invalid JSON', '[{"name":"US 10Y","last":true}]',
                    '[{"name":"US 10Y","last":[4.1]}]']
        for payload in payloads:
            html = '<title>US 10 Year Treasury Note Yield</title><script>TEChartsMeta = ' + payload + ';</script>' + table
            with self.subTest(payload=payload):
                quote, _, note = parse_yield_html(html, "United States")
                self.assertEqual(quote, 0.0412)
                self.assertIn("Table fallback", note)

    def test_yield_date_is_not_historical_news(self):
        html = """<title>Colombia 10-Year Government Bond Yield</title>
        <script>TEChartsMeta = [{"name":"Colombia 10Y","last":12.698}];</script>
        <article>Old news on January 2, 2020.</article>"""
        _, date, _ = parse_yield_html(html, "Colombia")
        self.assertEqual(date, "Not provided / not identified")
        html += '<script type="application/ld+json">{"@type":"Dataset","name":"Colombia 10-Year Government Bond Yield Data","temporalCoverage":"2002-09-16/2026-11-13"}</script>'
        self.assertEqual(parse_yield_html(html, "Colombia")[1], "2026-11-13")

    def test_source_numbers(self):
        self.assertEqual(parse_number("1,234.50%"), 1234.5)
        self.assertEqual(parse_number("−0.5"), -0.5)
        self.assertEqual(parse_number(".5"), 0.5)
        self.assertIsNone(parse_number("N/A"))
        for malformed in ("1.2.3", "1,5", "unknown", True, False, 10 ** 400):
            with self.subTest(value=malformed), self.assertRaises(DataUnavailable):
                parse_number(malformed)

    def test_huge_integer_calculation_input_is_controlled(self):
        with self.assertRaises(InputError):
            calculate(replace(REFERENCE_INPUTS, beta_u=10 ** 400))

    def test_newer_source_values_survive_older_country_cache(self):
        now = datetime.now(timezone.utc)
        old_time = (now - timedelta(minutes=1)).isoformat()
        old = Dataset("Source", "https://example.org", old_time, {"yield": Observation(0.04, "2026")})
        new = replace(old, retrieved_at=now.isoformat(), observations={"yield": Observation(0.05, "2026")})
        last_good = {"yield:USA": new}
        state = reconcile_sources({"yield:USA": old}, last_good, 3600)["yield:USA"]
        self.assertEqual(lookup(state, "yield").value, 0.05)
        self.assertEqual(last_good["yield:USA"].retrieved_at, new.retrieved_at)
        failure = FetchFailure(old.source, old.url, old_time, "Earlier failed request")
        state = reconcile_sources({"yield:USA": failure}, last_good, 3600)["yield:USA"]
        self.assertEqual(lookup(state, "yield").value, 0.05)
        self.assertEqual(state.error, "")

    def test_missing_selected_source_beta_retains_explicit_manual_row(self):
        frame = industry_comparison({"Other": Observation(1.2, "2026")}, REFERENCE_INPUTS, "Selected", True)
        selected = frame[frame["Selected"]].iloc[0]
        self.assertTrue(pd.isna(selected["Source beta"]))
        self.assertEqual(selected["Beta used"], REFERENCE_INPUTS.beta_u)
        self.assertTrue(selected["Beta overridden"])
        self.assertEqual(selected["USD WACC (decimal)"], calculate(REFERENCE_INPUTS).wacc_us)

    def test_source_failures_and_stale_label(self):
        failure = FetchFailure("Test source", "https://example.org", utc_now(), "HTTP 403")
        empty = reconcile_sources({"x": failure}, {}, 3600)["x"]
        self.assertIsNone(empty.dataset)
        self.assertEqual(empty.status, "Unavailable")
        old = Dataset("Test source", failure.url, "2026-01-01T00:00:00+00:00", {"x": Observation(0.04, "2025")})
        state = reconcile_sources({"x": failure}, {"x": old}, 3600)["x"]
        self.assertTrue(state.status.startswith("STALE"))
        self.assertEqual(state.dataset.retrieved_at, old.retrieved_at)
        self.assertIsNone(lookup(state, "missing"))

    def test_export_matches_results(self):
        p = REFERENCE_INPUTS
        r = calculate(p)
        text = project_csv(p, r, [], "Software", "Colombia", "Foreign", utc_now()).decode("utf-8-sig")
        rows = list(csv.DictReader(io.StringIO(text)))
        results = {row["field"]: float(row["value"]) for row in rows if row["section"] == "result"}
        for field, value in asdict(r).items():
            self.assertEqual(results[field], value)
        self.assertEqual(next(row["value"] for row in rows if row["field"] == "selected_currency"), "COP")
        self.assertEqual(next(row["currency"] for row in rows if row["field"] == "spread_foreign_auto"), "COP")

    def test_no_direct_wacc_conversion(self):
        r = calculate(REFERENCE_INPUTS)
        self.assertFalse(math.isclose(r.wacc_foreign, (1 + r.wacc_us) * r.factor - 1, abs_tol=1e-9))


def run_self_tests() -> int:
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(FinancialVerification)
    outcome = unittest.TextTestRunner(verbosity=2).run(suite)
    if outcome.wasSuccessful():
        r = calculate(REFERENCE_INPUTS)
        print(f"\nColombia verified: USD WACC {percent(r.wacc_us)} | COP WACC {percent(r.wacc_foreign)}")
        print("Offline tests passed. They do not certify external-source availability.")
    return 0 if outcome.wasSuccessful() else 1


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        raise SystemExit(run_self_tests())
    main()
