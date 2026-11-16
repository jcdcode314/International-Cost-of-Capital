# Cost of capital

A Python 3.12 / Streamlit dashboard for the Kellogg International Finance exercise, comparing a project's cost of equity, pretax debt, after-tax debt, and WACC in USD and a selected country's currency. The equations implement the methodology specified in the assignment, including separate currency conversion of equity and debt. The original assignment DOCX, all 13 embedded images in the supplied Cost of Capital slides DOCX, and the supplied home-screen reference were inspected. The actual slides confirm the beta relevering, multiplicative currency conversion, separate equity/debt costs, and Colombia example implemented here.

## Repository

Exactly three application files belong at the repository root:

```text
app.py
requirements.txt
README.md
```

No API keys, secrets, configuration files, databases, or supplementary project modules are required. Dependency versions in `requirements.txt` were installed and exercised with Python 3.12.

## Features

- Public Damodaran U.S. industry unlevered betas and large non-financial Baa2/BBB credit spread.
- OECD combined general-government statutory corporate tax rates, using the current 2.0 dataflow and the latest available reference year per jurisdiction, including 2026 where available.
- Best-effort public Trading Economics 10-year government yields; exact country and maturity checks.
- A broad country selector with explicit ISO country and currency mappings, including countries with unavailable public yields.
- Independent overrides for beta, source and project tax rates, yields, and borrowing spreads. The foreign yield is linked to the U.S. yield when the United States is selected, to enforce the assignment's equality requirement.
- Screenshot-matched white home screen with a light-gray sidebar, two-column yields, borrowing spreads, industry firm counts, four assumption metrics, and prominent equity/WACC values. Detailed tables and charts sit below the primary results.
- Sortable, searchable comparisons across all valid source industries, excluding aggregate market rows.
- Full-precision selected-project and all-industry CSV downloads with assumptions, manual-override indicators, and source provenance.
- Configurable caching, explicit stale-data labels, refresh without resetting manual values, and embedded offline verification.

## Install and run locally

Install **Python 3.12** and open a terminal in the repository root.

```bash
python --version
python -m venv .venv
```

Activate the environment on macOS/Linux:

```bash
source .venv/bin/activate
```

On Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Or on Windows Command Prompt:

```bat
.venv\Scripts\activate.bat
```

Then install and launch:

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Open the local URL printed by Streamlit, usually `http://localhost:8501`. The first run retrieves the external datasets. If a source cannot be retrieved, the dashboard identifies the missing input and provides its manual override; no illustrative quote replaces it.

To change the default 3,600-second cache lifetime, set the optional environment variable before launching. It must be a positive integer.

```bash
export COC_CACHE_TTL_SECONDS=1800
streamlit run app.py
```

In PowerShell, use `$env:COC_CACHE_TTL_SECONDS="1800"`. An invalid setting is reported and falls back to 3,600 seconds. Cache expiry is checked when the app next reruns; this is not a background streaming feed.

## Streamlit Community Cloud deployment

1. Confirm `app.py`, `requirements.txt`, and `README.md` are at the GitHub repository root.
2. Sign in at [Streamlit Community Cloud](https://share.streamlit.io/) and select **Create app** / **Deploy an app**.
3. Select this GitHub repository and its branch, normally `main`.
4. Set **Main file path** to **`app.py`**.
5. Open **Advanced settings** and choose **Python 3.12** from the Python version selector before deploying. If the platform changes this interface, use its runtime-version setting; do not rely on the platform's default Python version.
6. Deploy. No secrets are needed. Review the build log if installation fails, and use the app's **Data sources** tab to inspect retrieval failures from the deployed host.

The hosted app requires outbound HTTPS to the source domains. A site that is accessible locally may block the cloud host. The manual inputs work independently of those sources. This repository is ready to deploy; committing the files does not itself deploy a Streamlit server.

If uploading the files manually, open the GitHub repository, select **Add file > Create new file**, enter `app.py`, paste only the Python source, and select **Commit changes**. Repeat for `requirements.txt`. Edit the existing `README.md` (or create it if absent), replace its contents, and commit. Confirm all three files are at the root. Filename labels and Markdown fence characters must not be pasted into Python or requirements files; the Python filename is `app.py`, never `app.py.txt`.

## Dashboard controls and defaults

The initial industry is **Software (System & Application)**, the country is **United States**, ERP is **6.5%**, target debt financing is **30%**, and the project tax source is **U.S. statutory corporate tax rate**. All manual toggles begin off. Automatic beta, yields, taxes, and the U.S. spread come only from successful retrievals; the foreign spread is calculated from effective inputs.

The sidebar follows the reference: **Refresh source data**, **U.S. industry**, the supported **Unlevered beta measure**, **Annual equity risk premium (%)**, the **Fraction financed with debt (%)** slider, **Override BBB borrowing spread**, **Project country / currency**, and **Override corporate tax rate**. The beta measure offers only the exact ordinary unlevered beta supported by the validated source parser. The 0–99% debt slider defaults to 30%.

The home screen shows government yields in two columns, then borrowing spreads, the selected industry with actual source firm count/date, a row of four assumptions, and USD equity cost/WACC below a divider. Foreign-currency results appear as an additional row for a foreign project. Identical U.S. results are shown once; an intentionally different foreign-spread override is identified. Screenshot values are illustrative: the live source values are used at runtime.

Enter rates as percentages: `4` means a 4% yield; `1.5` means a borrowing spread of 1.5 percentage points (150 basis points). Beta is dimensionless. Source quotes, dates, and effective values remain visible. Missing observations stay unavailable; enabling an empty manual input requires a verified value and never silently substitutes zero.

**Enter U.S. 10-year yield manually** is available directly in the main yield section; selected-country manual yield entry is available alongside it for foreign projects. **Advanced assumptions and manual overrides** holds beta overrides, independent U.S./foreign tax-source overrides, the foreign BBB spread override and its automatic value, the tax-jurisdiction radio, and explicit reset. Source selection defaults to the U.S. statutory rate. The top-level project tax override takes precedence over the selected source's effective rate; either way, one project tax rate applies to both financing scenarios. Each active controlling override is included in export provenance.

Manual beta values are retained by industry, and foreign rates/spreads by country. Returning to a selection restores its saved settings. **Refresh source data** clears the cached retrieval function and performs new requests, preserving selections and manual values. **Reset all manual overrides** explicitly clears saved manual settings. When United States is selected, both scenarios always share the effective U.S. government yield, including its manual entry; the separate foreign-yield input is disabled.

Below the primary results are **Detailed calculation**, **All industries**, **Methodology and assumptions**, **Data sources**, **Verification**, and **Downloads** tabs. The verification tab presents the fixed Colombia example and pure mathematical checks without making source requests or replacing live assumptions. The full embedded test suite runs separately using `python app.py --self-test`.

The all-industry tab uses each industry's own source beta. If the selected industry has a beta override, only its selected row uses that manual beta, and both the source beta and effective beta are shown. All other assumptions are shared. A refresh that removes a source industry preserves the selected project and its saved manual beta instead of silently choosing another industry. When a source dataset is available, an explicitly overridden selected project with no current source row is included as an additional manual row with **Unavailable** source beta. If no industry dataset is available at all, a manual beta can support the selected project, but an invented all-industry table is never generated. Numerical overflow in an all-industry comparison produces a clear error while retaining valid selected-project results and source information.

## Methodology and variable definitions

All rates below are **decimals**, all weights are market-value weights, and calculations retain full floating-point precision. Only display values are rounded, to two or three decimal places on the compact home screen, four decimal places for detailed financial percentages, and additional precision for intermediate betas and factors.

| Variable | Meaning |
|---|---|
| `d` | Target debt share, `D / (D + E)` |
| `D`, `E` | Market values of project debt and equity |
| `D/E` | Target debt-to-equity ratio |
| `beta_U` | Industry unlevered beta, excluding the cash-corrected column |
| `beta_E` | Relevered project equity beta |
| `t` | Selected project statutory corporate tax rate |
| `ERP` | U.S. market-based equity risk premium |
| `Rf_US`, `Rf_FOREIGN` | U.S. and selected-country 10-year government yields |
| `s_US` | Large non-financial Baa2/BBB U.S. credit spread |
| `s_FOREIGN` | Modeled or manually overridden foreign BBB credit spread |
| `F` | Relative gross risk-free-rate currency factor |
| `Re_US`, `Re_FOREIGN` | USD and selected-currency equity costs |
| `Rd_US`, `Rd_FOREIGN` | Pretax borrowing costs |
| `Rd_aftertax_US`, `Rd_aftertax_FOREIGN` | Borrowing costs after the assumed interest tax shield |
| `WACC_US`, `WACC_FOREIGN` | Market-value-weighted average costs of capital |

The model implements:

```text
D/E = d / (1 - d)
beta_E = beta_U * [1 + (1 - t) * D/E]
Re_US = Rf_US + beta_E * ERP
Rd_US = Rf_US + s_US
Rd_aftertax_US = Rd_US * (1 - t)
WACC_US = (1 - d) * Re_US + d * Rd_aftertax_US

F = (1 + Rf_FOREIGN) / (1 + Rf_US)
Re_FOREIGN = (1 + Re_US) * F - 1
s_FOREIGN_auto = s_US * F
s_FOREIGN = s_FOREIGN_auto, unless its manual override is enabled
Rd_FOREIGN = Rf_FOREIGN + s_FOREIGN
Rd_aftertax_FOREIGN = Rd_FOREIGN * (1 - t)
WACC_FOREIGN = (1 - d) * Re_FOREIGN + d * Rd_aftertax_FOREIGN
```

The model assumes zero debt beta and zero tax-shield beta. It uses one project beta, tax rate, ERP, and debt share in both scenarios. No separate foreign-country equity premium is added. A foreign spread override affects foreign debt costs and WACC only; it does not affect USD results or foreign equity cost. Foreign WACC is **not** obtained by converting the final USD WACC directly.

Debt shares satisfy `0 <= d < 1`. The UI slider allows up to 99%; the calculation function also accepts other finite shares strictly below 100%. Zero debt is valid. Tax rates must be within 0–100%; this implementation requires nonnegative betas, ERP, and credit spreads. Government yields may be negative but must exceed -100%, making both gross risk-free returns positive. Missing, non-finite, or numerically overflowing inputs are rejected.

## Sources, reference dates, and retrieval dates

| Input | Public source and validation |
|---|---|
| U.S. industry beta | [Damodaran XLS](https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls), with [official HTML fallback](https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/Betas.html). Searches workbook sheets for the actual table; uses the exact **Unlevered beta** header, not `Beta` or the cash-corrected measure. XLS preserves more precision; the HTML fallback has the source's rounded values. The exact Number of firms column supplies the industry firm-count metadata when published. |
| BBB spread | [Damodaran ratings](https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/ratings.html). Expands HTML table spans and selects **Baa2/BBB**, **large non-financial service firms**, and **Spread**, rather than coverage or the financial-services block. Source percentage points are divided by 100. |
| Statutory taxes | [OECD database](https://www.oecd.org/en/data/datasets/corporate-income-tax-rates-database.html); [documented Data Explorer flow](https://data-explorer.oecd.org/vis?df%5Bag%5D=OECD.CTP.TPS&df%5Bid%5D=DSD_TAX_CIT%40DF_CIT&df%5Bvs%5D=2.0); [public SDMX CSV query](https://sdmx.oecd.org/public/rest/v1/data/OECD.CTP.TPS,DSD_TAX_CIT@DF_CIT,2.0/.A.CIT_C.ST..S13...?startPeriod=2000&dimensionAtObservation=AllDimensions). Uses the HTTP Accept header `text/csv;version=2.0`. |
| Government yields | [Trading Economics public bonds](https://tradingeconomics.com/bonds), then the corresponding public `/{country}/government-bond-yield` page if necessary. For the U.S., the fallback is [US 10 Year Treasury Note Yield](https://tradingeconomics.com/united-states/government-bond-yield). Recognizes `US`/`United States` and `UK`/`United Kingdom` labels. Parses public table rows or embedded current quote JSON; rejects real/TIPS/inflation-linked instruments and contradictory maturity identifiers. Validates the known U.S. nominal symbol `USGG10YR:IND` when present. Does not call private chart APIs, use paid credentials, or substitute a different maturity. |

The OECD parser validates annual frequency `A`, combined corporate rate `CIT_C`, statutory targeting `ST`, general-government sector `S13`, percentage-of-taxable-income units `PT_INC_TAX`, unscaled units `UNIT_MULT=0`, and non-applicable dividend/rate-structure/tax-base dimensions `_Z`. It matches explicit ISO country codes and selects each jurisdiction's newest valid reference year up to the current calendar year. Rates are not effective tax rates and do not combine corporate and shareholder dividend taxes. Different countries can have different latest reference years. The [original 2024 assignment reference](https://www.oecd.org/en/publications/corporate-tax-statistics-2024_9c27d6e8-en/full-report/statutory-corporate-income-tax-rates_440f9286.html) is background, rather than a frozen source of current rates.

The beta workbook's **Date updated**, the official tables' **Date of Analysis**, and OECD **reference year** are displayed separately from successful **retrieval timestamps in UTC**. For bond yields, an identifiable dataset temporal-coverage end date or dated current summary is used. A row's abbreviated date is explicitly labeled if it omits its year/time zone. A page refresh date or a historical news date is not treated as the observation date; unidentified dates remain unknown.

HTTP uses connect/read timeouts of 4/12 seconds per attempt, with one retry for temporary connection failures, timeouts, HTTP 429, and selected 5xx errors. The independent sources are fetched concurrently. Successful responses and source-specific failures are cached until the configurable TTL or an explicit refresh. Changing country creates a separate cached bundle and may retrieve shared sources again; returning to an older country bundle retains any newer successful common-source observation already known to the session. Superseded cached failures do not overwrite newer successes. The cache key includes the application revision, so a parser upgrade obtains new source responses instead of reusing a prior revision's failed cache. The app does not automatically retry repeatedly while idle.

Truncated XLS responses are converted to source errors and can use the official HTML fallback. Malformed or misaligned OECD CSV observations are rejected without crashing the dashboard. Boolean values are rejected as financial observations. Optional malformed bond metadata can fall back to a separately validated nominal 10-year table on the same public page.

If refresh fails, a **session's** last successfully retrieved dataset can remain in use only with a visible **STALE — refresh failed** label, its original retrieval timestamp, and the latest error. There is no packaged market snapshot or durable local fallback. Without a previous successful observation, an input is unavailable until retrieved or manually supplied. A missing country tax rate is distinct from failure of the OECD request itself.

Trading Economics scraping is **best effort**, not guaranteed. Its pages can be dynamic, reformatted, or blocked by location or hosting provider, including HTTP-200 pages that contain no usable quote. The bonds listing returned no quote rows during development checks, but automatic U.S. retrieval was verified through the dedicated public U.S. page after correcting recognition of its `US 10Y` label. No API key, browser automation, or additional dependency is needed for that fallback. Use a verified current 10-year quote with the corresponding override if both public retrieval paths fail. Manual values are identified as user inputs; the app cannot independently establish their source or observation date. Consult a professional tax source for project-specific law and deductibility constraints.

## Colombia slide verification

Reference assumptions: beta `0.80`, debt share `30%`, project tax `35%`, ERP `6.5%`, U.S. yield `4%`, Colombia yield `12%`, U.S. spread `1.5` percentage points, and foreign-spread override **off**.

| Metric | Expected |
|---|---:|
| D/E | 0.42857143 |
| Relevered equity beta | 1.02285714 |
| Currency factor F | 1.07692308 |
| USD cost of equity | 10.6485714% |
| USD pretax debt | 5.5000% |
| USD after-tax debt | 3.5750% |
| USD WACC | **8.5265%** |
| COP cost of equity | 19.1600% |
| COP BBB spread | 1.6153846% |
| COP pretax debt | 13.6153846% |
| COP after-tax debt | 8.8500% |
| COP WACC | **16.0670%** |

These are test fixtures, never default market observations. To reproduce the example in the UI, select Colombia, enable overrides for beta, the relevant project tax rate, both yields, and the U.S. spread, then enter these reference assumptions. Keep the foreign spread override disabled and use the matching tax-source radio selection.

## Tests and verification

```bash
python -m py_compile app.py
python app.py --self-test
python -m pip check
```

Self-test mode runs without launching a Streamlit server or requesting market data and exits nonzero on failure. It verifies the Colombia results against full-precision expected values using tight floating-point tolerances; equal yields; zero debt; isolated foreign-spread overrides including a zero spread; both tax selections; linked U.S. yields; all-industry consistency and override isolation; high valid leverage; zero ERP; valid negative yields; missing inputs; invalid rates, denominators and non-finite inputs; exact source columns and rating groups; OECD dimensions and reference-year selection; country/maturity checks; stale-data handling; and exported numerical agreement. Self-tests are embedded in `app.py`; no test file is required.

Revision **2026-11-16.3** contains **31 embedded tests**, including corrupted-XLS fallback, truncated OECD data and oversized error pages, U.S./U.K. aliases, nominal-versus-TIPS identity, optional metadata fallback, Boolean/oversized-number rejection, newest-source preservation, and a missing source beta with an explicit selected-project override, industry firm-count extraction, preservation of industry names split across Excel HTML spans, and all-industry borrowing costs. The revision is visible in **Data sources**. Independent review also checked 1,000 random financial cases against a 60-digit Decimal oracle and 1,000 currency-conversion/foreign-spread sensitivity identity cases. No financial-formula corrections were required.

Streamlit interaction checks cover default values, the Colombia inputs, tax switching, foreign-spread isolation, industry and country changes, refresh failures, malformed providers, explicit reset, disappearing/reappearing industries, empty selector recovery, fully manual operation, and controlled all-industry overflow. Browser checks additionally exercise the rendered dashboard and CSV download buttons. Live public-source availability is a separate observation, not guaranteed by these tests.

For integration verification, start the app, verify its defaults, enter verified manual data for missing inputs, switch the tax radio, and inspect the substituted equations. Change the foreign BBB spread and confirm USD results remain fixed. Refresh and confirm manual values persist. The selected industry's all-industry row should equal the main dashboard using its effective beta. Export and compare decimal CSV rates against percentage display values (multiply CSV rates by 100). External availability must be checked separately from the offline mathematical tests.

## Download results

Use **Download selected-project CSV** or **Download all-industry CSV** in **Downloads**. Both appear once the required inputs are complete; the latter additionally requires the industry dataset. The all-industry export always includes all industries, even when the displayed table is filtered by search.

Selected-project CSV uses long-form sections for metadata, exact calculation inputs, automatic/effective source inputs, active overrides, and calculated results. The all-industry CSV uses one row per industry and includes source/effective beta, U.S. firm counts where available, both currencies' equity, pretax debt, after-tax debt and WACC results, shared inputs, beta provenance, and JSON input provenance. Both carry currency labels and a generation timestamp. Rates are explicitly labeled as decimals and serialized with up to 17 significant digits. Empty optional values are missing data, not zero. Source timestamps attached to overridden fields describe the **automatic** observation, not an independent verification of the user's manual value.

## Financial assumptions and limitations

The operating-risk proxy is a U.S. industry average and may not represent a particular foreign project. Government yields can include sovereign, liquidity, and term components, and countries sharing a currency can still have different government yields. The mappings specify local legal currency, not a guarantee that an available government security is a pure risk-free benchmark. Currency-regime and rate-source changes require maintenance.

For USD-using countries such as Ecuador, Panama, and El Salvador, both financing scenarios can carry a USD label while their national government-yield proxies differ. The app preserves the assigned equations and displays a targeted explanation: any modeled difference reflects those country/sovereign-yield proxies. The relative-yield factor needs particular caution in a same-currency comparison.

CAPM uses a U.S. market ERP without an additional country equity premium. The foreign BBB spread is modeled by `s_US * F`, not retrieved as a local corporate quote. Fisher and relative-PPP assumptions motivate the nominal gross-return adjustment but need not hold over a project's horizon. No exchange-rate path, hedging cost, country-risk decomposition, or FX forecast is estimated.

The model assumes constant target market-value financing weights, zero debt/tax-shield beta, and fully usable interest deductions at one selected project tax rate in both scenarios. Statutory headline rates are not project-specific effective rates, and interest-deduction restrictions are not modeled. Negative debt yields are permitted algebraically; the resulting tax adjustment follows the assigned formula even where practical tax treatment would differ.

**A lower nominal WACC does not make a currency economically cheaper.** Match project cash-flow currency, inflation assumptions, and discount-rate currency before interpreting a valuation. This is an educational implementation of the supplied methodology.
