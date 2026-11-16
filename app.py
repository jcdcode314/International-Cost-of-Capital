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
import sys
import time
import unicodedata
import unittest
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, replace
from datetime import datetime, timezone
from typing import Any

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
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
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
    for industry, observation in sorted(betas.items()):
        if is_aggregate(industry):
            continue
        overridden = industry == selected and beta_override
        beta_used = p.beta_u if overridden else observation.value
        result = calculate(replace(p, beta_u=beta_used))
        rows.append({
            "Industry": industry, "Selected": industry == selected,
            "Beta overridden": overridden, "Source beta": observation.value,
            "Beta used": beta_used, "Equity beta": result.beta_e,
            "USD equity (decimal)": result.re_us,
            "Foreign equity (decimal)": result.re_foreign,
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
    if isinstance(value, (float, int)):
        return float(value) if math.isfinite(value) else None
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
            for rr in range(rowspan):
                for cc in range(colspan):
                    cells[row_number + rr, column + cc] = cell.get_text(" ", strip=True)
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


def parse_beta_xls(content: bytes) -> tuple[dict[str, float], str]:
    try:
        book = xlrd.open_workbook(file_contents=content)
    except (xlrd.XLRDError, xlrd.compdoc.CompDocError, ValueError) as exc:
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
        return beta_rows(grid), date
    raise DataUnavailable("Damodaran workbook has no recognized U.S. industry beta sheet.")


def parse_beta_html(text: str) -> tuple[dict[str, float], str]:
    soup = BeautifulSoup(text, "html.parser")
    for table in soup.find_all("table"):
        grid = table_grid(table)
        if any("industry name" in [normalize(v) for v in row]
               and "unlevered beta" in [normalize(v) for v in row] for row in grid):
            return beta_rows(grid), published_date(soup.get_text(" ", strip=True))
    raise DataUnavailable("Official beta HTML has no recognized unlevered-beta table.")


def fetch_betas() -> Dataset:
    errors = []
    for url, parser in ((BETA_XLS, parse_beta_xls), (BETA_HTML, parse_beta_html)):
        try:
            response = http_get(url, "*/*")
            betas, reference = parser(response.content if url == BETA_XLS else response.text)
            if len(betas) < 20:
                raise DataUnavailable("Fewer than 20 valid industries; refusing a likely incomplete table.")
            note = "Exact Unlevered beta column; U.S. industries; aggregate market rows excluded."
            if errors:
                note += " Official HTML fallback used, at its published rounded precision. " + " | ".join(errors)
            return Dataset("Damodaran U.S. industry betas", url, utc_now(),
                           {name: Observation(beta, reference) for name, beta in betas.items()}, note)
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
    reader = csv.DictReader(io.StringIO(text.lstrip("\ufeff")))
    required = {"REF_AREA", "FREQ", "MEASURE", "TARGETING", "UNIT_MEASURE", "SECTOR",
                "DIV_TREATMENT", "RATE_STRUCTURE", "TAX_BASE", "TIME_PERIOD", "OBS_VALUE", "UNIT_MULT"}
    if not required.issubset(reader.fieldnames or []):
        raise DataUnavailable("OECD CSV schema does not contain the required statutory-rate dimensions.")
    latest: dict[str, Observation] = {}
    exact = {"FREQ": "A", "MEASURE": "CIT_C", "TARGETING": "ST",
             "UNIT_MEASURE": "PT_INC_TAX", "SECTOR": "S13",
             "DIV_TREATMENT": "_Z", "RATE_STRUCTURE": "_Z", "TAX_BASE": "_Z"}
    for row in reader:
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
                    if not (has_ten_year(name) and name_matches(name, country)):
                        continue
                    end = str(dataset.get("temporalCoverage", "")).split("/")[-1]
                    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", end):
                        try:
                            return datetime.strptime(end, "%Y-%m-%d").date().isoformat()
                        except ValueError:
                            continue
    # Inspect only a current summary/description, never dated news or forecasts.
    summary = soup.select_one("#market-summary, #historical-desc #description")
    description = soup.find("meta", attrs={"name": "description"})
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
    board_ten_year = board and has_ten_year(description_text + " " + title)
    page_identity = (has_ten_year(title) and any(
        normalize(title).startswith(normalize(name) + " ") for name in country_names(country)))
    # Public embedded JSON can retain more precision than the displayed row.
    if page_identity:
        for script in soup.find_all("script"):
            script_text = script.get_text()
            match = re.search(r"\bTEChartsMeta\s*=\s*(\[)", script_text)
            if not match:
                continue
            try:
                metadata, _ = json.JSONDecoder().raw_decode(script_text[match.start(1):])
            except json.JSONDecodeError as exc:
                raise DataUnavailable("Public bond metadata is no longer valid JSON.") from exc
            if not isinstance(metadata, list):
                continue
            for item in metadata:
                if not isinstance(item, dict):
                    continue
                identity = str(item.get("full_name") or item.get("name") or "")
                if not has_ten_year(identity) or not name_matches(identity, country):
                    continue
                quote = parse_number(item.get("last", item.get("value")))
                if quote is not None and quote > -100:
                    row = next((a.find_parent("tr") for a in soup.find_all("a", href=country_url)
                                if a.find_parent("tr") is not None), None)
                    return quote / 100, observation_date(soup, row, country), "Public page's current 10-year quote metadata."
    # Row fallback requires both exact country path and a verified 10-year maturity.
    for anchor in soup.find_all("a", href=country_url):
        row = anchor.find_parent("tr")
        if row is None or not name_matches(anchor.get_text(" ", strip=True), country):
            continue
        if not (board_ten_year or page_identity or has_ten_year(anchor.get_text(" ", strip=True))):
            continue
        value_cell = row.select_one('[id="p"], .market-widget-last')
        if value_cell is None:
            continue
        quote = parse_number(value_cell.get_text(" ", strip=True))
        if quote is not None and quote > -100:
            return quote / 100, observation_date(soup, row, None if board else country), "Public 10-year yield table; source display precision."
    raise DataUnavailable(f"An exact {country} 10-year government yield was not found on the public page.")


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


def reconcile_sources(bundle: dict[str, Dataset | FetchFailure], last_good: dict[str, Dataset],
                      ttl: int) -> dict[str, SourceState]:
    states = {}
    for key, result in bundle.items():
        if isinstance(result, Dataset):
            last_good[key] = result
            age = (datetime.now(timezone.utc) - datetime.fromisoformat(result.retrieved_at)).total_seconds()
            status = "Cached response" if age > 5 else "Retrieved"
            if age > ttl:
                status = "STALE — cache lifetime exceeded"
            states[key] = SourceState(result.source, result.url, result, status, result.retrieved_at)
        else:
            old = last_good.get(key)
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
    saved = store.setdefault(key, {"enabled": False, "value": auto * scale if auto is not None else None})
    toggle_key, value_key = "override_toggle:" + key, "override_value:" + key
    if toggle_key not in st.session_state:
        st.session_state[toggle_key] = saved["enabled"]
    if value_key not in st.session_state:
        st.session_state[value_key] = saved["value"]
    enabled = st.toggle("Override " + name, key=toggle_key, disabled=linked is not None)
    entered = st.number_input(
        name + (" (%)" if percentage else ""), min_value=minimum, max_value=maximum,
        step=0.0001 if percentage else 0.01, format="%.6f", key=value_key,
        disabled=not enabled or linked is not None, placeholder="Enter a verified value")
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
        scenario = "USD" if field.endswith("_us") else currency if field.endswith("_foreign") else "shared/derived"
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


def main() -> None:
    st.set_page_config(page_title="International Cost of Capital", page_icon="🌐", layout="wide")
    st.markdown("""<style>
    .stApp {background:var(--background-color);}
    [data-testid="stMetric"] {border:1px solid #dce4eb;border-radius:12px;padding:18px;}
    [data-testid="stMetricValue"] {color:#008e7b;}
    h1 {letter-spacing:-.045em;} .block-container {padding-top:2rem;max-width:1450px;}
    [data-testid="stSidebar"] {border-right:1px solid #dce4eb;}
    </style>""", unsafe_allow_html=True)
    st.caption("INTERNATIONAL FINANCE  /  VALUATION LAB")
    st.title("International Cost of Capital Calculator")
    st.write("Compare project equity, borrowing costs, and WACC in U.S. dollars and a selected country's currency.")
    with st.sidebar:
        st.header("Project assumptions")
        country = st.selectbox("Target country", sorted(COUNTRIES),
                               index=sorted(COUNTRIES).index("United States"), key="country")
        st.caption(f"Selected currency: {COUNTRIES[country]['currency']} · Country: {COUNTRIES[country]['iso']}")
        refresh = st.button("Refresh Market Data", width="stretch", type="primary")
    raw_ttl = os.getenv("COC_CACHE_TTL_SECONDS", "3600")
    try:
        ttl = int(raw_ttl)
        if ttl < 1:
            raise ValueError("TTL must be positive")
    except ValueError:
        ttl = 3600
        st.warning("Invalid COC_CACHE_TTL_SECONDS; using a 3,600-second cache lifetime.")
    cached_fetch = st.cache_data(ttl=ttl, show_spinner=False)(retrieve_market_bundle)
    if refresh:
        cached_fetch.clear()
    with st.spinner("Retrieving public beta, statutory tax, BBB spread, and 10-year yield data…"):
        bundle = cached_fetch(country)
    last_good = st.session_state.setdefault("last_good_sources", {})
    states = reconcile_sources(bundle, last_good, ttl)
    if refresh:
        failed = sum(isinstance(item, FetchFailure) for item in bundle.values())
        if failed:
            st.warning(f"Refresh completed with {failed} source error(s). Review Sources & status; manual inputs are preserved.")
        else:
            st.success("Market data refreshed. Project assumptions and manual values are preserved.")
    for state in states.values():
        if state.error or state.status.startswith("STALE"):
            st.warning(f"{state.source}: {state.status}. " + (state.error or "Cached data exceeded its lifetime."))
    beta_state, tax_state, spread_state = states["betas"], states["taxes"], states["spread"]
    us_yield_state = states["yield:USA"]
    foreign_yield_state = states[f"yield:{COUNTRIES[country]['iso']}"]
    betas = beta_state.dataset.observations if beta_state.dataset else {}
    industries = sorted(betas) if betas else [DEFAULT_INDUSTRY]
    with st.sidebar:
        prior_industry = st.session_state.get("industry", DEFAULT_INDUSTRY)
        if prior_industry not in industries:
            st.session_state["industry"] = DEFAULT_INDUSTRY if DEFAULT_INDUSTRY in industries else industries[0]
        industry = st.selectbox("Industry", industries,
                               index=industries.index(DEFAULT_INDUSTRY) if DEFAULT_INDUSTRY in industries else 0,
                               key="industry")
        erp = st.number_input("Equity risk premium (%)", min_value=0.0, value=6.5,
                              step=0.1, format="%.4f", key="erp") / 100
        debt = st.number_input("Target debt financing share (%)", min_value=0.0,
                               max_value=99.99, value=30.0, step=1.0, format="%.4f", key="debt") / 100
        tax_selection = st.radio("Project corporate tax source",
                                 ["U.S. statutory corporate tax rate", "Selected country's statutory corporate tax rate"],
                                 key="tax_selection")
        st.caption("The selected project tax rate applies to BOTH financing scenarios.")
        st.button("Reset all manual overrides", on_click=reset_manual_inputs, width="stretch")
        with st.expander("Manual overrides & automatic values", expanded=True):
            beta = override_control("Unlevered beta", f"beta:{industry}", lookup(beta_state, industry),
                                    beta_state, percentage=False)
            us_tax = override_control("U.S. corporate tax rate", "tax:USA", lookup(tax_state, "USA"),
                                      tax_state, maximum=100.0)
            foreign_tax = override_control("Foreign corporate tax rate", f"foreign_tax:{COUNTRIES[country]['iso']}",
                                           lookup(tax_state, COUNTRIES[country]["iso"]), tax_state, maximum=100.0)
            us_yield = override_control("U.S. 10-year yield", "yield:USA", lookup(us_yield_state, "yield"),
                                        us_yield_state, minimum=-99.9999)
            foreign_yield = override_control("Foreign 10-year yield", f"foreign_yield:{COUNTRIES[country]['iso']}",
                                             lookup(foreign_yield_state, "yield"), foreign_yield_state,
                                             minimum=-99.9999, linked=us_yield if country == "United States" else None)
            us_spread = override_control("U.S. BBB spread", "spread:USA", lookup(spread_state, "BBB"), spread_state)
            # This automatic spread is derived from EFFECTIVE source inputs, including overrides.
            auto_foreign = None
            effective_foreign_yield = comparison_yield(country, us_yield.value, foreign_yield.value)
            if us_spread.value is not None and us_yield.value is not None and effective_foreign_yield is not None:
                auto_foreign = us_spread.value * (1 + effective_foreign_yield) / (1 + us_yield.value)
                if not math.isfinite(auto_foreign):
                    auto_foreign = None
            derived = Dataset("Model-derived foreign BBB spread", "", utc_now(), {},
                              "s_FOREIGN = effective s_US × (1 + effective Rf_FOREIGN) / (1 + effective Rf_US).")
            derived_state = SourceState(derived.source, "", derived, "Derived from current inputs", derived.retrieved_at)
            foreign_spread = override_control("Foreign BBB spread", f"foreign_spread:{COUNTRIES[country]['iso']}",
                                              Observation(auto_foreign, "Derived, not a market quote") if auto_foreign is not None else None,
                                              derived_state)
        st.caption(f"Cache lifetime: {ttl:,} seconds. Rates are entered in percent; spreads in percentage points.")
    records = [beta, us_tax, foreign_tax, us_yield, foreign_yield, us_spread, foreign_spread]
    selected_tax = us_tax if tax_selection.startswith("U.S.") else foreign_tax
    required = {"Unlevered beta": beta.value, "Project corporate tax rate": selected_tax.value,
                "U.S. 10-year yield": us_yield.value, "U.S. BBB spread": us_spread.value}
    if country != "United States":
        required["Foreign 10-year yield"] = effective_foreign_yield
    if foreign_spread.override_active:
        required["Foreign BBB spread override"] = foreign_spread.value
    missing = [name for name, value in required.items() if value is None]
    st.markdown("**Market assumptions**")
    assumption_rows = [{"Assumption": record.name,
                        "Automatic": f"{record.automatic_value:.8f}" if record.name == "Unlevered beta" and record.automatic_value is not None else percent(record.automatic_value),
                        "Used": f"{record.value:.8f}" if record.name == "Unlevered beta" and record.value is not None else percent(record.value),
                        "Mode": "Manual override" if record.override_active else "Derived" if record is foreign_spread else "Automatic / linked",
                        "Reference": record.reference_date, "Source status": record.status}
                       for record in records]
    assumption_rows += [
        {"Assumption": "Equity risk premium", "Used": percent(erp), "Mode": "Project assumption"},
        {"Assumption": "Target debt share", "Used": percent(debt), "Mode": "Project assumption"},
        {"Assumption": "Project tax applied to both scenarios", "Used": percent(selected_tax.value), "Mode": tax_selection},
        {"Assumption": "Selected country's currency", "Used": COUNTRIES[country]["currency"], "Mode": country},
    ]
    with st.expander("Automatic and effective market assumptions", expanded=False):
        st.dataframe(pd.DataFrame(assumption_rows), hide_index=True, width="stretch")
    p, result = None, None
    if missing:
        st.warning("Complete the following inputs using the independent sidebar overrides: " + ", ".join(missing) + ".")
    else:
        try:
            p = Inputs(beta.value, debt, selected_tax.value, erp, us_yield.value,
                       effective_foreign_yield, us_spread.value,
                       foreign_spread.value if foreign_spread.override_active else None)
            result = calculate(p)
        except InputError as exc:
            p = None
            st.error(str(exc))
    generated_at = utc_now()
    comparison_tab, industries_tab, methodology_tab, sources_tab = st.tabs(
        ["Financing comparison", "All industries", "Methodology", "Sources & status"])
    with comparison_tab:
        if p and result:
            st.caption(f"{industry} · {country} · Project tax {percent(p.tax_rate)} · Debt share {percent(p.debt_share)}")
            render_comparison(p, result, country)
            st.download_button("Download selected-project CSV",
                               project_csv(p, result, records, industry, country, tax_selection, generated_at),
                               file_name="selected_project_cost_of_capital.csv", mime="text/csv", on_click="ignore")
        else:
            st.info("Results appear when the required data or manual inputs are available. No example values are substituted.")
    with industries_tab:
        if not betas:
            st.warning("The industry dataset is unavailable. A manual beta can support the selected project; "
                       "all-industry comparison requires a successfully retrieved industry dataset.")
        elif p:
            if beta.override_active:
                st.info("Selected industry's beta is manually overridden. Only its row uses that override; "
                        "other industries retain their individual source betas.")
            table = industry_comparison(betas, p, industry, beta.override_active)
            search = st.text_input("Search industries", placeholder="Type part of an industry name")
            visible = table[table["Industry"].str.contains(search, case=False, regex=False)] if search else table
            st.caption(f"{len(table)} valid industries · Selected: {industry}. Click a column header to sort. Export includes all rows.")
            formats = {column: "{:.4%}" for column in table.columns if "(decimal)" in column}
            formats.update({"Source beta": "{:.8f}", "Beta used": "{:.8f}", "Equity beta": "{:.8f}"})
            st.dataframe(visible.style.format(formats), hide_index=True, width="stretch", height=520)
            st.download_button("Download all-industry CSV",
                               all_industry_export(table, p, country, beta_state, records, generated_at),
                               file_name="all_industries_cost_of_capital.csv", mime="text/csv", on_click="ignore")
        else:
            st.info("Complete the required project assumptions to calculate all industries.")
    with methodology_tab:
        render_methodology(p, result)
        st.caption("All calculations use full numerical precision. Displayed substitutions are rounded for readability; CSV rates are decimals.")
    with sources_tab:
        render_sources(states, records)
    st.caption("Educational valuation model · Public-source data is best effort · Python 3.12 · Run python app.py --self-test for offline verification.")


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
        for malformed in ("1.2.3", "1,5", "unknown"):
            with self.subTest(value=malformed), self.assertRaises(DataUnavailable):
                parse_number(malformed)

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
