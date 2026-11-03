# Context document: "SP vs FCST" Dashboard (latest file FY27/2611/REAS_2_261103.xlsm)

> **Instructions for any AI assistant reading this document**
> - This is a self-contained, plain-text description of an Excel demand-planning dashboard (Coty, Barcelona Supply Chain Hub). It does not depend on any specific AI tool, plugin or file access.
> - Treat it as the source of truth to answer questions about what the dashboard shows, where the data comes from, how each figure is calculated, how to update it, its limitations and how to interpret the results.
> - If a question is not covered here, say so explicitly instead of guessing. Do not invent figures: the only numbers you can quote are the ones in this document (sections 12 and 18) or numbers the user gives you.
> - To test or audit the sell-in / sell-out trends, use section 18 (formulas, cell map, worked examples and reference values).
> - The dashboard was built with the help of an AI assistant and validated by recalculating the figures independently.
> - Version history: the 30-Oct-2026 version compared SP vs BASE (+ INNO) with a "Remove INNO" switch. **ver2 (data of 01-Nov-2026)** adds the OSP layer, a layer selector (BASE / INNO / OSP), a Planner filter, charts by House, a 7-slide export (9 slides since the gap heatmaps were added) and automatic detection of the last closed month. **ver2 was chosen as the final version** and is stored as `FY27/2611/REAS_2_261101.xlsm` (final versions live in `<FY>/<YYMM>/`). **ver3** (`extra_features/REAS_2_261101_ver3.xlsm`, extra-features folder, not part of the final version) is ver2 plus the exclusion of up to 3 product lines; the items marked "(ver3 only)" below exist only in that file. **REAS_2_261103.xlsm (03-Nov-2026)** is the same tool with a new BASE / INNO / OSP extract (SP, OSP-SP, EPOS and PLANNER unchanged) plus the rule "past months: SP = FCST" (§5.2). Sections 12 and 18 keep the figures of the 01-Nov file; the 03-Nov figures are in §12.1. **REAS_2_261103_DA.xlsm** (repository root, in development) is REAS_2_261103 plus the DA and GSP input tabs and the new **DA IMPACT** tab (impact of the Demand Assumptions on the FCST, §19).

---

## 1. Purpose

The dashboard compares the **Sales Plan (SP)** with **our forecast (FCST)**, split into three **layers**:

| Layer | Our forecast (FCST side) | Sales plan (SP side) |
|---|---|---|
| **BASE** | BASE tab (system forecast volume) | BASE part of the SP tab = MAX(0, SP − INNO) |
| **INNO** | INNO tab (innovation volume) | INNO part of the SP tab = MIN(SP, INNO) |
| **OSP** (one-shot promotions) | OSP tab | OSP-SP tab (OSP volume planned by Sales) |

SP tab = BASE + INNO together (they have to be split using the INNO forecast, as before). OSP is planned separately by Sales in the OSP-SP tab, so it is **added**, not subtracted.

Comparison by house, IPPL Line and product line (PL), month by month, for fiscal year FY27 (Jul-26 → Jun-27).

It also shows:
- **Sell-in trends** (volume of the selected layers, last 6 and 12 closed months vs last year).
- **Sell-out trends** (EPOS in $, last 26 and 52 weeks vs last year).
- **LY** (last-year volume) and **trend projections** (LY × (1 + trend)).
- **Forecast accuracy at lag 3**: **SPP3** (bias) and **WAPE3** (weighted absolute percentage error), closed months.
- An **EXPORT** button that produces a 7-slide PDF for reporting.

---

## 2. File structure

The file is a **macro-enabled Excel workbook (.xlsm)**. **Only the Dashboard tab is visible**: the file is saved that way and, when macros are enabled, every time the workbook is opened all other tabs are hidden again (Workbook_Open). To see a tab: right-click a tab → Unhide….

| Tab | Type | Content |
|---|---|---|
| **SP** | Manual input | Sales Plan (BASE + INNO): House, IPPL Line and 12 months (row 1 = months Jul-26…Jun-27). Rows 2–201 are read. |
| **OSP-SP** | Manual input | Sales Plan of the one-shot promotions, same structure as SP. Rows 2–201 are read. |
| **BASE** | Manual input | System extract. One row per product line; column A = IPPL LINE (XLOOKUP formula on IPPL MAP 1), column B = Product_line; column blocks per FY / half / quarter / month with measures (Act + Fcst, Year Ago, % vs YA, Snapshot, % L12M vs YA, % L6M vs YA, SPP3 %, WAPE3, Cuts…). Rows 5–504 are read. |
| **INNO** | Manual input | Same structure as BASE, innovations only. Rows 5–504 are read. |
| **OSP** | Manual input | Same structure as BASE, one-shot promotions only. Rows 5–504 are read. |
| **EPOS** | Manual input | Sell-out (Power BI export): Category, House, Brand, Line; $ and YoY Δ / YoY % for L26W and L52W. Rows 3–602 are read. |
| **PLANNER** | Master data | Team ownership: House (column A) → Planner (column B). "?" or empty = "(unassigned)". Houses missing from the tab are also "(unassigned)". |
| **IPPL MAP 1 / IPPL MAP 2** | Master data | Mapping GMM Line / Product line → IPPL Line and House. |
| **TABLE** | Automatic (formulas only) | Consolidates SP, BASE, INNO, OSP and OSP-SP by IPPL Line and month. Never typed into. |
| **Dashboard** | Visible | Filters, KPIs, tables, charts and EXPORT button. |
| **DASH CALC** | Automatic (hidden) | Calculation engine of the dashboard. Do not edit. |
| **EXPORT** | Automatic (hidden) | Layout of the 7 slides exported to PDF. |

**Units:** BASE, INNO and OSP come in thousands in the system; TABLE and the dashboard multiply them × 1000 so they are comparable with SP / OSP-SP (volume in units). EPOS is in $ and is only used for % trends.

---

## 3. How TABLE works (automatic consolidation)

- **IPPL Line list:** de-duplicated union of the IPPL Lines in SP (first, in SP order), then OSP-SP, BASE, INNO and OSP. Capacity: 150 lines.
- **House of each line:** taken from SP; if not there, from OSP-SP; if not there, from IPPL MAP 1; otherwise **"(no house)"** (only visible with House = All and Planner = All).
- **Planner of each house:** from the PLANNER tab (column X of TABLE).
- **Blocks (column A = TYPE):** rows 2–151 = SP, 152–301 = BASE, 302–451 = INNO, 452–601 = OSP, 602–751 = OSP-SP.
- **Values:** SP / OSP-SP = sum of the tab for that IPPL Line and month (empty month stays empty, not 0). BASE / INNO / OSP = sum of the month's "Act + Fcst" column for that IPPL Line × 1000.
- **Adding data:** new houses, IPPL Lines, PLs or planners added in the input tabs appear automatically in TABLE, the dropdowns and every calculation.

Monthly "Act + Fcst" columns used in BASE / INNO / OSP: Jul-26 CP, Aug-26 DC, Sep-26 DP, Oct-26 EC, Nov-26 EP, Dec-26 FC, Jan-27 FP, Feb-27 GC, Mar-27 GP, Apr-27 HC, May-27 HP, Jun-27 IC. The Year Ago (YA) column of each month is the next column; SPP3 % is 9 columns to the right and WAPE3 10 columns to the right (e.g. Jul-26: CQ, CY, CZ).

---

## 4. Dashboard filters (yellow cells)

| Filter | Cell | Options | Effect |
|---|---|---|---|
| Planner | C5 | All + planners of the PLANNER tab | Keeps only the houses of that planner (and filters the House and IPPL Line dropdowns). |
| House | C6 | All + houses (of the selected planner) | Filters everything. If the house does not belong to the planner, a warning appears and all houses of the planner are shown. |
| IPPL Line | C7 | All + lines of the selected house / planner | If the line is not in the filter, a warning appears and all lines are shown. |
| Month from / Month to | C8 / C9 | Jul-26…Jun-27 | Range for totals, KPIs and "period total" details. Default **Nov-26 → Jun-27** (SP / OSP-SP have no volume before Nov-26). |
| Metric for per-line charts | C10 | FCST / SP / Gap | Metric of the monthly history charts by House and by IPPL Line. |
| Include BASE | C11 | Yes / No | Includes the BASE layer on both sides. |
| Include INNO | C12 | Yes / No | Includes the INNO layer on both sides. |
| Include OSP | C13 | Yes / No | Includes the OSP layer on both sides. |
| Product line | C14 | All + PLs within the Planner/House/IPPL filter | Product-line view (see §8). If the PL does not fit the filter (or is excluded), a warning appears and everything is shown. |
| Month for trends | C15 | Closed months (default: last closed month) | End month of the sell-in trend windows; uses the "% L6M vs YA" / "% L12M vs YA" columns of that month (see §7.1). It also defines the past months where SP = FCST (§5.2, 261103 onwards). |
| Exclude product line 1 / 2 / 3 (ver3 only) | C15 / C16 / C17 | (none) + PLs within the Planner/House/IPPL filter | Removes up to 3 product lines (see §5.1). |
| Remove excluded PLs from SP too (ver3 only) | C18 | Yes / No | Yes (default): the forecast of the excluded PLs is also deducted from the SP. No: they are removed from FCST only. |

Warnings and notices are shown in F15 (ver3: the list of excluded product lines in F16). Trends and forecast accuracy **do not depend** on Month from / Month to (they always use the closed months).

---

## 5. Comparison logic and layer selection

For each IPPL Line and month (the lowest level at which SP exists):

- **SP BASE part** = MAX(0, SP − MAX(0, INNO))
- **SP INNO part** = MAX(0, MIN(SP, MAX(0, INNO)))
- **SP OSP part** = OSP-SP
- **SP compared** = Include BASE × SP BASE part + Include INNO × SP INNO part + Include OSP × OSP-SP
- **FCST compared** = Include BASE × BASE + Include INNO × INNO + Include OSP × OSP

Only then is it summed to house / planner / total. Consequences:
- All layers included (default): SP + OSP-SP vs BASE + INNO + OSP.
- BASE + OSP (equivalent of the old "Remove INNO = Yes" plus OSP): (SP − INNO, floored at 0) + OSP-SP vs BASE + OSP.
- BASE + INNO only: exactly the old "Remove INNO = No" comparison (SP vs BASE + INNO).
- SP − INNO is never negative (SP and the INNO plan are planned by different people). A negative INNO forecast is treated as 0 when splitting the SP.
- No layer selected: everything is 0 and a warning appears.

### 5.2 Past months: SP = FCST (from REAS_2_261103)
- The SP / OSP-SP tabs have no past months. For every month up to the **Month for trends** (C15; default = last closed month, Oct-26), the SP of each layer is set equal to its FCST, per IPPL Line and month: SP BASE part = BASE, SP INNO part = INNO, OSP part = OSP (i.e. the actuals already in the forecast).
- Effect: those months show SP = FCST (gap 0) in every table, chart and heatmap, and ranges that include them (e.g. Jul-26 → Jun-27) give full-year SP vs FCST. Months after the Month for trends use the SP tabs as before; a month between the Month for trends and the first SP month (e.g. Month for trends = Aug-26 → Sep–Oct 26) has no SP.
- DASH CALC: row 67 = flag (month ≤ B28); rows 3201–3750 (SP parts) use the FCST blocks of TABLE when the flag is 1; row 65 = SP available (flag or SP tabs); row 68 = the SP tabs have data; B27 = first month with SP in the tabs.

### 5.1 Excluding product lines (ver3 only)
- The excluded PLs (matched by name, in any layer) are removed from: FCST (every layer, total, by House / IPPL Line / month), sell-in trends, LY and projections, forecast accuracy, the PL rankings, charts and table, and sell-out (EPOS line with the same name).
- SP does not exist by product line. With **Remove excluded PLs from SP too = Yes**, the forecast of the excluded PLs of each layer is deducted from the SP part of that layer, per IPPL Line and month, floored at 0 (same idea as splitting INNO out of the SP): SP BASE part − excluded BASE, SP INNO part − excluded INNO, OSP-SP − excluded OSP. With **No**, the SP is unchanged and only FCST loses the volume (the gap then shows the full volume of the excluded PLs).
- The excluded PLs are shown in F16 and in the filter line of every exported slide.

Labels show the layers in use, e.g. "SP (BASE + INNO + OSP)" and "FCST (BASE + INNO + OSP)". At product-line level there is no SP (only FCST).

---

## 6. What the Dashboard shows

1. **Period KPIs:** SP, FCST, Gap (FCST − SP) and Gap % = FCST / SP − 1; last closed month; period.
2. **By layer (range total)** (F11:M14): FCST, SP part and Gap % for BASE, INNO and OSP (excluded layers are labelled "(excluded)").
3. **Trend tiles** (traffic light): Sell-in 6M, Sell-in 12M, Sell-out 6M, Sell-out 12M. Green ≥ +5%, amber between −5% and +5%, red ≤ −5%.
4. **Monthly table:** SP, FCST, Gap, Gap % (colour scale: red = FCST below SP, white ≈ 0, green = FCST above SP; full colour at ±30%), BASE, INNO and OSP; "Range total" column.
5. **SP vs FCST by month chart** (lines) with the panel "Extra lines on the chart": **LY** (sum of the Year Ago columns of the selected layers) and **projections** LY × (1 + trend) for Sell-in 6M / 12M, Sell-out 6M / 12M and a **manual trend**.
6. **FCST by layer by month:** stacked bars BASE + INNO + OSP with the SP line (excluded layers are hidden).
7. **Forecast accuracy (lag 3):** SPP3 bars and WAPE3 line for the 12 FY months (values only for closed months), with dotted thresholds WAPE 40% and SPP ±5%.
8. **Detail by House** (period total, SP vs FCST, horizontal bars, sorted by the larger of SP and FCST) and **Detail by IPPL Line** (top 30 by SP), side by side.
9. **Detail by product line** (FCST, top 25 by volume).
10. **Monthly history by House** (one line per house; metric per filter C10), **by IPPL Line** (top 30 by SP) and **by product line** (FCST, top 25).
11. **Gap FCST vs SP by month — by House and by IPPL Line** (heatmap tables, above the trends section): FCST − SP per month for every house / the top 30 IPPL Lines in the filter, with a range total; red = FCST below SP, green = FCST above SP. Only months in the range that have SP are filled; follows every filter (planner, house, IPPL Line, layers); n/a in the product-line view. (Final version only; not in ver3.)
12. **Sell-in vs sell-out trends by IPPL Line** (chart) and tables by IPPL Line and by product line (PL table: BASE, INNO, OSP, FCST, trends, IPPL Line).

---

## 7. Key definitions and calculations

### 7.0 Last closed month (automatic)
A month is closed when any BASE / INNO / OSP row has an SPP3 value for it. With the data of 01-Nov-2026 the closed months are **Jul–Oct 26** (last closed month **Oct-26**). Trend windows and accuracy move automatically when a new month is closed; nothing has to be edited.

### 7.1 Sell-in trends (selected layers, volume)
- Windows end at the **Month for trends** (filter C15, closed months only; default = last closed month): **L6M** = the 6 months ending there, **L12M** = the 12 months ending there (e.g. Oct-26: L6M = May-26 → Oct-26, L12M = Nov-25 → Oct-26). FY26 months come from the Year Ago columns of the equivalent FY27 months; FY27 months from "Act + Fcst".
- Last-year volume of each row: **YA = Act / (1 + %)**, with the system "% L6M vs YA" (6M) or "% L12M vs YA" (12M) **of the selected month** (month-level columns, e.g. Oct-26: EK and EJ). The aggregate is volume-weighted: **trend = ΣAct / ΣYA − 1**.
- Rows without that % (mostly discontinued items with no current volume, where the % would be −100%, and launches): their YA is **estimated** from their Year Ago columns of the window months available in FY27, scaled to 6 / 12 months (launches → YA = 0). The Dashboard note (row 16) shows the share of volume that uses the system % ("coverage"; Oct-26: 70% for 6M, 77% for 12M). Earlier versions counted these rows with YA = 0, which inflated the trends (e.g. Total 6M +14.1% instead of +5.8%).
- Only rows of the included layers count.

### 7.2 Sell-out trends (EPOS, $)
- 6M = **L26W**, 12M = **L52W**. **trend = Σ$ / Σ($ − YoY Δ) − 1.**
- Each EPOS line is mapped to its IPPL Line via IPPL MAP 1. At product-line level it is matched by exact line name in EPOS; if there is no match, n/a.
- The layer selection does not affect sell-out (EPOS does not split layers).

### 7.3 Forecast accuracy (lag 3)
- **SPP3 = Actual / Forecast lag3 − 1** (negative = over-forecast; positive = under-forecast).
- **WAPE3 = Σ|Actual − Forecast lag3| / ΣActual.**
- **Aggregation (percentages are never summed):** for each row and month the dashboard rebuilds **Forecast = Actual / (1 + SPP3)** and **|Error| = WAPE3 × Actual**, sums them over the selected layers and filters, and recalculates the %.
- Rows with Actual = 0 or without KPI cannot be rebuilt and are excluded (small volume).
- Reference thresholds: WAPE 40%, SPP ±5%.

### 7.4 LY and projections
- Monthly LY = Σ Year Ago of the selected layers × 1000, filtered.
- Projection = LY × (1 + selected trend).

---

## 8. Product-line view (filter C14)
- SP does not exist at PL level: it is hidden (KPIs "n/a", no SP line in charts) and a notice appears.
- FCST (selected layers), LY, projections, sell-in / sell-out trends and SPP3 / WAPE3 are calculated for that PL only.
- Charts, tables and trends by House and by IPPL Line show only the House / IPPL Line that holds that PL, with the FCST of that PL only (SP bars = 0 and "n/a" in the IPPL table, because SP does not exist by PL); sell-in / sell-out trends by IPPL Line use only that PL.
- PL list = PLs of BASE, then PLs of INNO / OSP that are not in BASE.

---

## 9. EXPORT button (macro)

- Located top right of the Dashboard. Macros must be enabled when opening the file (and, if the file comes from email/SharePoint, it may need to be unblocked in File Properties).
- Produces a **9-slide PDF** (A4 landscape, one slide per page) **without asking**, named **REAS_2_EXPORT_YYMMDDHHMMSS.pdf** (always unique), and opens it when done.
- **Location:** same folder as the Excel file. If the file is on OneDrive/SharePoint, the macro finds the locally synced folder via the OneDrive registry; if that fails it tries the web path; as a last resort it saves to Downloads and shows the path.
- **Slide content** (each slide has the title, the **active filters** — Planner, House, IPPL Line, PL, layers, metric — and what is shown at the top, and a **COMMENTS** box at the bottom):
  1. KPIs + by-layer table + trend tiles + monthly table + SP vs FCST chart (with the LY/projection lines switched on).
  2. FCST by layer (stacked) + forecast accuracy (SPP3/WAPE3).
  3. Detail by House + detail by IPPL Line (bars).
  4. Detail by product line (bars).
  5. Monthly history by House (lines).
  6. Monthly history by IPPL Line.
  7. Monthly history by product line.
  8. Gap FCST vs SP by month — by House (heatmap).
  9. Gap FCST vs SP by month — by IPPL Line (heatmap).
- Not exported (they remain in the dashboard): the sell-in vs sell-out trends by IPPL Line chart and the two final tables.

---

## 10. How to update the data
1. Unhide the relevant tab (right-click the Dashboard tab → Unhide…).
2. **SP / OSP-SP:** paste House, IPPL Line and the 12 months from row 2 (row 1 = months).
3. **BASE / INNO / OSP:** paste the new extract with the same column structure; make sure the XLOOKUP formula in column A (IPPL LINE) reaches every row.
4. **EPOS:** paste the Power BI export with the same structure.
5. **PLANNER:** keep one row per house with its planner.
6. Save. The other tabs are hidden again the next time the file is opened (or hide them manually). Everything else (TABLE, DASH CALC, Dashboard, EXPORT) recalculates automatically, including the last closed month.

Capacities: SP 200 rows, OSP-SP 200, BASE / INNO / OSP 500 each, EPOS 600, 150 IPPL Lines, 30 houses, 200 PLANNER rows, 1,000 product lines. Charts by IPPL Line show the top 30 by SP; by House all houses (30).

---

## 11. Technical notes (ver2)
- The engine uses only classic Excel functions (SUMIFS, SUMPRODUCT, INDEX/MATCH, COUNTIFS); the only dynamic function is the existing XLOOKUP in column A of BASE / INNO / OSP.
- The workbook is set to recalculate fully when opened.
- Macros: ExportModule (ExportSlides, ShowOnlyDashboard) and ThisWorkbook (Workbook_Open → ShowOnlyDashboard). The VBA project is stored as source code only, so Excel compiles it on first open.

---

## 12. Reference figures (data of 01-Nov-2026, all filters = All, Nov-26 → Jun-27)

| Indicator | All layers (default) |
|---|---|
| SP (BASE + INNO + OSP) | 531,020,999 |
| FCST (BASE + INNO + OSP) | 565,041,364 |
| Gap / Gap % | 34,020,365 / +6.4% |
| BASE: FCST vs SP BASE part | 291,976,463 vs 305,926,681 (−4.6%) |
| INNO: FCST vs SP INNO part | 149,150,360 vs 130,887,972 (+14.0%) |
| OSP: FCST vs OSP-SP | 123,914,541 vs 94,206,347 (+31.5%) |
| Sell-in 6M / 12M (Month for trends = Oct-26) | +5.8% / +4.4% (BASE −5.0% / +4.2%, INNO +59.4% / +21.8%, OSP −1.9% / −5.7%) |
| Sell-out 6M / 12M (EPOS) | +2.7% / −0.2% |
| SPP3 Jul / Aug / Sep / Oct 26 | −0.7% / −19.2% / −9.9% / −12.1% |
| WAPE3 Jul / Aug / Sep / Oct 26 | 45.0% / 58.9% / 41.5% / 42.7% |
| LY (Nov-26 → Jun-27) | 556,365,725 |

Other examples (all layers, Nov-26 → Jun-27):
- (ver3 only) Excluding Burberry Her EDP, Goddess Amber Vanilla and Burberry Multiline (removed from SP too): SP 500,192,902 vs FCST 526,094,915 (+5.2%).
- Planner = Adria: SP 237,954,161 vs FCST 243,924,326 (+2.5%).
- Planner = German, House = Gucci: SP 90,730,106 vs FCST 103,345,071 (+13.9%).
- House = Gucci, BASE only: SP BASE part 49,650,568 vs BASE 48,410,379 (−2.5%).
- Top houses (SP vs FCST): BURBERRY 148,894,679 vs 149,043,847; Gucci 90,730,106 vs 103,345,071; CP-Marc Jacobs 78,307,751 vs 82,044,992.

---

### 12.1 REAS_2_261103.xlsm (data of 03-Nov-2026, all filters = All, Month for trends = Oct-26)
| Indicator | Nov-26 → Jun-27 (default range) | Jul-26 → Jun-27 (full year) |
|---|---|---|
| SP (BASE + INNO + OSP) | 531,020,999 | 913,405,705 |
| FCST (BASE + INNO + OSP) | 526,034,262 | 908,418,968 |
| Gap | −4,986,737 (−0.9%) | −4,986,737 (Jul–Oct gap = 0) |
| Sell-in 6M / 12M | −2.6% / −11.2% | (trends do not depend on the range) |
| Volume with system % (coverage) 6M / 12M | 50% / 44% | |

Full year by layer (FCST vs SP part): BASE 453,521,596 vs 493,509,048; INNO 228,897,745 vs 211,228,872; OSP 225,999,627 vs 208,667,785.

Note on the 03-Nov extract: the month-level "% L6M / % L12M vs YA" columns are much emptier than on 01-Nov (rows with % L6M in Oct-26: BASE 158 of 381, INNO 4 of 193, OSP 14 of 158; on 01-Nov: 268 of 402, 56 of 214, 89 of 187), so most of the trend is the YA-column estimate (§7.1).

## 13. Validation performed (ver2)
- An independent recalculation (Python, straight from the input tabs, without the workbook formulas) was compared with the workbook recalculated in LibreOffice in 11 scenarios (default; planners Adria / German / Magalie / (unassigned); houses Gucci / BURBERRY; IPPL Line; product line; house not belonging to the planner; no layer; single layers and layer combinations; Oct–Dec range): SP, FCST, each layer, SP parts, sell-in trends, SPP3 / WAPE3 for the 12 months, LY by month and SP / FCST by house match exactly in every scenario, with no formula errors.
- ver3: the same comparison in 15 scenarios (the 11 above plus exclusions of 1–3 PLs, with and without removal from SP, combined with planner / house / layer filters, and a PL that is both selected and excluded): no mismatches and no formula errors. With no exclusion, ver3 gives exactly the ver2 figures.
- The exported slides were rendered and checked (7 pages, charts within their slide).
- REAS_2_261103: the values Excel calculated and saved in the file match the independent re-calculation (42 checks: totals, layers, SP parts, trends, accuracy, houses). After adding "past months: SP = FCST", 15 scenarios (including full-year ranges, Month for trends = Aug-26 and planner / house / layer filters) match with no formula errors.

---

## 14. Limitations and points of attention
- SP and OSP-SP exist only by IPPL Line (not by PL).
- SP / OSP-SP have no volume before Nov-26; that is why the default range starts in Nov-26. Months in the range without SP count as SP = 0 in the totals.
- The SP split into BASE and INNO uses our INNO forecast (INNO tab) because the SP tab does not separate them.
- OSP-SP is assumed to be additional to the SP tab (it contains lines that are not in SP and months where OSP-SP > SP).
- Trends and accuracy rely on the % provided by the system; if the system changes its definitions they must be reviewed.
- BASE / INNO / OSP rows whose product line has no IPPL Line in IPPL MAP 1 are not assigned to any line.
- Some chart legends show empty entries when there are fewer lines than slots (Excel limitation without extra macros).
- EPOS will be reviewed later (structure and mapping).

---

## 15. Next steps / improvements

1. **Consider cuts in trends and actuals.** The extracts include "o9 Cuts" and "Cuts YA" columns that are not used today. Add the cut volume back to actuals (and year-ago actuals) so that sell-in trends, LY and accuracy reflect unconstrained demand, with an option to show both views.
2. **Identify more forecast layers.** OSP is now a separate layer; next candidates: statistical baseline, promotions/events, management adjustments.
3. **Identify past promotional or exceptional activities.** Flag months/products affected by promotions, one-off deals, launches, stock-outs or other events, so that trends, LY projections and accuracy can be read excluding those effects.
4. **Flexibility to add more houses and PLs to the scope.** Keep IPPL MAP 1 and PLANNER complete for every new product line / house, raise capacities if needed and review chart readability when many lines are selected.

---

## 16. Frequently asked questions (short answers)

**What is FCST?** Our forecast: the sum of the selected layers (BASE, INNO and OSP tabs).

**What is OSP?** One-shot promotions. Our forecast is in the OSP tab; the Sales plan for them is in the OSP-SP tab.

**Why is the SP split into BASE and INNO parts?** The SP tab contains BASE and INNO together; the INNO forecast is used to separate them (SP INNO part = MIN(SP, INNO), BASE part = the rest).

**Can the SP BASE part be negative?** No: it is floored at 0 per IPPL Line and month.

**How do I compare only BASE (like the old "Remove INNO")?** Include BASE = Yes, Include INNO = No, Include OSP = No (or Yes to keep OSP on both sides).

**What does a positive Gap % mean?** FCST is above SP: our forecast expects more volume than the plan.

**What does a negative SPP3 mean?** Over-forecast: actuals came in below what was forecast 3 months earlier.

**Why don't trends change when I move Month from/to?** They always measure the last 6/12 closed months vs last year.

**When does the last closed month change?** Automatically, when the system extract brings SPP3 values for a new month.

**What does the Planner filter do?** It keeps only the houses owned by that planner (PLANNER tab) in every figure, chart and dropdown.

**Why does a trend show n/a?** There is no last-year volume (new product or line) or no match in EPOS.

**Where is the exported PDF saved?** In the same folder as the Excel file (or in Downloads, with a message, if that folder cannot be written). Name REAS_2_EXPORT_YYMMDDHHMMSS.pdf.

**Why can't I see the data tabs?** They are hidden on purpose (and hidden again on every open); right-click a tab → Unhide.

---

## 17. Glossary
- **SP:** Sales Plan (SP tab = BASE + INNO).
- **OSP:** One-shot promotions (OSP tab = our forecast; OSP-SP tab = Sales plan).
- **FCST:** our forecast = selected layers of BASE + INNO + OSP.
- **BASE:** system forecast by product line (formerly REAS).
- **INNO:** innovation volume.
- **IPPL Line:** product planning group (the level at which SP exists).
- **PL:** Product line.
- **Planner:** team member owning a house (PLANNER tab).
- **LY / YA:** Last Year / Year Ago (same month of the previous year).
- **L6M / L12M:** last 6 / 12 closed months.
- **L26W / L52W:** last 26 / 52 weeks (sell-out).
- **EPOS:** point-of-sale sell-out data ($).
- **SPP3:** forecast bias at lag 3 (Act / Fcst − 1).
- **WAPE3:** weighted absolute percentage error at lag 3 (Σ|Act − Fcst| / ΣAct).
- **Lag 3:** forecast made 3 months before the month being measured.
- **Cuts:** volume that could not be served (cut) versus the original demand.

## 18. Testing the sell-in and sell-out trends (audit guide)

This section is written so that the trends can be re-calculated by hand (or in another tool / conversation) and compared with the file. All references are to `FY27/2611/REAS_2_261101.xlsm` (data of 01-Nov-2026).

### 18.1 Inputs used (BASE, INNO and OSP tabs, one row per product line, rows 5–504)
Column A = IPPL Line, column B = Product line. Volumes are in thousands (the dashboard multiplies by 1000). For each FY27 month the block has 13 columns; the ones used for trends are:

| Month | Act + Fcst | Year Ago (YA) | % L12M vs YA | % L6M vs YA | SPP3 % (closed-month flag) |
|---|---|---|---|---|---|
| Jul-26 | CP | CQ | CW | CX | CY |
| Aug-26 | DC | DD | DJ | DK | DL |
| Sep-26 | DP | DQ | DW | DX | DY |
| Oct-26 | EC | ED | EJ | EK | EL |
| Nov-26 | EP | EQ | EW | EX | EY |
| Dec-26 | FC | FD | FJ | FK | FL |
| Jan-27 | FP | FQ | FW | FX | FY |
| Feb-27 | GC | GD | GJ | GK | GL |
| Mar-27 | GP | GQ | GW | GX | GY |
| Apr-27 | HC | HD | HJ | HK | HL |
| May-27 | HP | HQ | HW | HX | HY |
| Jun-27 | IC | ID | IJ | IK | IL |

- "Act + Fcst" of a closed month = actual sell-in. "Year Ago" of a FY27 month = actual of the same month in FY26.
- "% L6M vs YA" / "% L12M vs YA" of month m = the system trend of that product line for the 6 / 12 months ending in m (they are only filled for closed months: Jul–Oct 26 in this data).
- Coverage of those % in this data (Oct-26, rows with % L6M / % L12M): BASE 268 / 290 of 402 rows, INNO 56 / 109 of 214, OSP 89 / 113 of 187.

### 18.2 Step by step
1. **Closed months:** month k (1 = Jul-26 … 12 = Jun-27) is closed if any BASE / INNO / OSP row has an SPP3 value in that month. Closed = Jul–Oct 26 → 4 (DASH CALC B25).
2. **Trend month t:** Dashboard C15 "Month for trends" (default = last closed month). DASH CALC B28 = its index (Oct-26 = 4; a non-closed or empty choice falls back to B25), B29 = its label.
3. **Window weights** (DASH CALC rows 58–61, columns C..N = months 1..12):
   - L6M actual: k in (t−6, t] → use "Act + Fcst" of month k.
   - L6M last year: k > t + 6 → use "Year Ago" of month k (= actual of the FY26 month).
   - L12M actual: k ≤ t → "Act + Fcst". L12M last year: k > t → "Year Ago".
   - Example t = Oct-26 (4): L6M = May-26, Jun-26 (YA columns HQ, ID of May-27, Jun-27) + Jul–Oct 26 (CP, DC, DP, EC). L12M = Nov-25 → Jun-26 (YA columns EQ, FD, FQ, GD, GQ, HD, HQ, ID) + Jul–Oct 26.
4. **Per row** (sell-in rows of DASH CALC: BASE rows 1201–1700 = tab rows 5–504; INNO 1701–2200; OSP 3801–4300):
   - F = Act L6M = 1000 × Σ window volumes (step 3). H = Act L12M.
   - % used = "% L6M vs YA" (6M) or "% L12M vs YA" (12M) **of month t** (Oct-26: EK / EJ).
   - If the % exists and 1 + % > 0: **YA = Act / (1 + %)** (G for 6M, I for 12M). Flag O (6M) / BL (12M) = 1.
   - If not (blank %: discontinued items whose % would be −100%, launches, rows without history): **YA = 1000 × Σ "Year Ago" of the FY27 months inside the window × (6 / number of those months)** for 6M, or × (12 / number of those months) for 12M. With t = Oct-26: 6M → (CQ + DD + DQ + ED) × 6/4; 12M → (CQ + DD + DQ + ED) × 12/4. Flag = 0.
   - E = the row is in the filter: layer included, Planner / House (via the IPPL Line of column A), IPPL Line, Product line.
5. **Aggregate:** Sell-in 6M = Σ(E × F) / Σ(E × G) − 1 (DASH CALC B14, Dashboard F10). Sell-in 12M = Σ(E × H) / Σ(E × I) − 1 (B15, Dashboard H10). Coverage = share of Σ(E × Act) in rows with flag 1 (B30 for 6M, B31 for 12M; shown in Dashboard B16).
6. The same row values feed the trend tables by IPPL Line (DASH CALC AD/AE 5301+) and by product line (AD/AE 101+), and the projections (LY × (1 + trend)).

### 18.3 Worked examples (t = Oct-26; tab values are in thousands, results in units)
- **Gorgeous Gardenia (BASE row 9, has %):** Act Jul–Oct 26 = 1,652.8 + 1,516.9 + 2,255.0 + 996.4; May–Jun 26 (HQ, ID) = 1,343.8 + 2,039.2 → Act L6M = 9,803,994. % L6M Oct (EK) = −22.33% → YA L6M = 9,803,994 / (1 − 0.2233) = 12,622,122 → row trend −22.3%. (In the PL view the PL also has an OSP row, so the PL trend is −23.9%.)
- **Goddess EDP (BASE row 5, has %):** Act L6M = 15,626,993; % L6M Oct = −16.97% → YA L6M = 18,821,047.
- **Fragrances Radiant Grace (INNO row 106, no %, discontinued):** Act L6M = 0; YA Jul–Oct 25 (CQ, DD, DQ, ED) = 17.4 + 40.2 + 46.3 + 143.8 → YA L6M estimate = 1000 × 247.6 × 6/4 = 371,433. It now pulls the trend down; before this change it counted as YA = 0.

### 18.4 Reference values (all layers, Nov-26 → Jun-27; trends do not depend on the month range)
Sell-in 6M / 12M by Month for trends (all filters = All):

| Month for trends | Total | BASE only | INNO only | OSP only |
|---|---|---|---|---|
| Oct-26 (default) | +5.8% / +4.4% | −5.0% / +4.2% | +59.4% / +21.8% | −1.9% / −5.7% |
| Sep-26 | +4.2% / +10.5% | −4.0% / +17.3% | +36.1% / +9.4% | −0.4% / +0.2% |
| Aug-26 | +9.9% / +21.5% | +7.9% / +47.4% | +40.1% / −0.6% | −6.1% / +2.1% |
| Jul-26 | +16.5% / +27.6% | +6.5% / +67.1% | +49.4% / −10.3% | +15.2% / +9.5% |

Month for trends = Oct-26, all layers:

| Filter | Sell-in 6M | Sell-in 12M | Coverage 6M / 12M |
|---|---|---|---|
| All | +5.8% | +4.4% | 70% / 77% |
| House = BURBERRY | −2.8% | +4.2% | 74% / 77% |
| House = Gucci | −4.6% | −7.0% | |
| House = CP-Marc Jacobs | +22.5% | +14.6% | 71% / 79% |
| House = Kylie Makeup | +9.8% | +30.3% | |
| House = Hugo Boss | −20.5% | −7.6% | |
| Planner = Adria | +6.7% | +7.7% | |
| Planner = German | +22.8% | +10.2% | |
| Planner = Magalie | −16.0% | −12.7% | 90% / 93% |
| Planner = Elvira | −31.9% | −9.9% | |
| Planner = (unassigned) | +10.0% | +16.8% | |
| Product line = Gorgeous Gardenia | −23.9% | −11.9% | 100% / 100% |
| Product line = Burberry Her EDP | −17.7% | −16.5% | |
| Product line = Goddess EDP | −9.1% | −10.6% | |
| BASE only, Month for trends = Sep-26 | −4.0% | +17.3% | 90% / 93% |

The method before the "Month for trends" change (FY % columns, rows without % = YA 0) gave Total +14.1% / +6.9% and INNO +146.8% for 6M; use that only to confirm that a file is the old version.

### 18.5 Sell-out (EPOS) test
- EPOS rows 3–602: D = line (product line name), E = $ L26W, F = YoY Δ L26W, H = $ L52W, I = YoY Δ L52W.
- Last-year $ = $ − YoY Δ. Sell-out 6M = Σ $ L26W / Σ($ L26W − Δ) − 1; 12M the same with L52W. Lines are mapped to IPPL Line / House through IPPL MAP 1 (column C line → column E IPPL Line) and filtered like the sell-in rows; "Total" rows are ignored.
- Reference (all filters = All): Sell-out 6M +2.7%, 12M −0.2%. By house it matches the YoY % of the Power BI export.

### 18.6 How these values were checked
The figures in 18.4 were produced by an independent Python re-calculation (reading the input tabs directly) and match the workbook recalculated in LibreOffice to 6 decimals in every case tested (15 filter scenarios, trend months Jul–Oct).

---

## 19. DA impact (Demand Assumptions) — `REAS_2_261103_DA.xlsm` (in development, separate tab)

**Status:** new feature, built in its own tabs (**DA IMPACT**, visible, and **DA CALC**, hidden). It is not in the Dashboard nor in the EXPORT slides yet ("first a separate tab, later we decide how to incorporate it"). The file lives in the repository root while it is being developed; the rest of the workbook (Dashboard, EXPORT, TABLE, DASH CALC) is unchanged.

### 19.1 What a DA is and what the tab shows
- **DAs (Demand Assumptions)** are adjustments that planners load on top of the system forecast (pipe fills, new doors, promotions, deletes, "Prometheus" removals, etc.). They come in **units**, by EAN and month (DA tab).
- **The DAs are already included in the FCST.** The tab does not change the forecast: it splits it into **Base FCST (without DA) + DA impact = FCST final**. The FCST final is always the same number as in the Dashboard for the same filters.
  - Example: FCST = 100k $, DA = −10k $ → Base FCST = 110k $, DA = −10k $, FCST final = 100k $.
  - Example: FCST = 100k $, DA = +10k $ → Base FCST = 90k $, DA = +10k $, FCST final = 100k $.
- To compare with the FCST (which is in $), the DA units are valued with the **GSP** (GSP tab, price per EAN).

### 19.2 New input tabs
| Tab | Content |
|---|---|
| **DA** | One row per DA and EAN. Row 1 = month of each value column (dates, J1…), row 2 = headers, data from row 3. Columns: A Reason_Code, B DA Description, C Initiative Type, D Created by, E Submission Date, F House, G Product Line, H EAN, I Measure, J… monthly units (8 months Nov-26 → Jun-27 in this extract; up to 12 columns J–U are read; months outside FY27 are ignored). Capacity rows 3–4002. |
| **GSP** | Price per EAN: A CFG, B EAN, C Currency, D Product Line ("Name (code)"), E Brand, F House, G Month ("2026.M08"), H GSP EAN Raw. Capacity rows 2–3001. |

### 19.3 Calculation (DA CALC)
For each DA row:
1. **Units by FY month**: the date in DA row 1 is matched to the FY27 months (Jul-26 … Jun-27).
2. **GSP**: value of the EAN in the GSP tab (if the EAN appears in several months, the latest month). If the EAN is not in the GSP tab → **average GSP of its product line** (GSP product line mapped with IPPL MAP 2, or the name without the "(code)"), then **average GSP of its house**; if none → 0 (counted as $0 and flagged).
3. **DA $ = units × GSP.**
4. **Mapping**: the product line (DA column G) is looked up in BASE, INNO and OSP (exact name, then trimmed) → IPPL Line (column A of that tab) → TABLE line → House and Planner (exactly as the FCST in the Dashboard). PLs not found, or without an IPPL Line in TABLE, are "not mapped to the FCST" and shown separately.
5. **Layer split**: a product line can be in BASE, INNO and OSP. The DA $ of each month is split between the layers **in proportion to the FCST of that product line in that month** (Act + Fcst of the month, negatives = 0). If the product line has no FCST that month: all to BASE, else INNO, else OSP. (The extract is by product line, so the layer of each EAN is not known; this is the assumption to review.)
6. **Aggregation** with the filters: FCST final = TABLE (selected layers, filtered lines, months in range) — same as the Dashboard; DA = Σ DA $ of the selected layers; **Base FCST = FCST final − DA**. DA + = sum of the positive DA values (row × month), DA − = sum of the negative ones.

### 19.4 DA IMPACT tab (layout)
- **Filters (yellow)**: Planner, House, IPPL Line, Month from / to, Include BASE / INNO / OSP, DA reason code (All or one). Invalid combinations fall back to "All" with a warning in row 15 (same logic as the Dashboard).
- **KPIs**: FCST final, Base FCST (without DA), DA impact (net), DA % of FCST, DA + (adds to Base), DA − (removes from Base), DA units (net), period. **Breakdown** box: Base FCST + DA = FCST final (same as the Dashboard).
- **Monthly table** (12 FY months + range total): FCST final, Base FCST, DA net, DA +, DA −, DA % of FCST, DA by layer, DA units. Months outside the range in grey.
- **Charts**: (1) by month: grey = Base FCST kept in the FCST final, green = DA that adds (grey + green = FCST final), red hatched = DA that removes (grey + red = Base FCST), dark line with labels = FCST final; (2) bridge Base FCST → + DA → − DA → FCST final (range total); (3) DA by reason code; (4) DA by House (top 15).
- **Tables**: by House (all, sorted by |DA|), by IPPL Line (top 30 by |DA|), by reason code (DA +, DA −, net, %, units, # lines), top 25 DAs by |$| (description, $, units, # lines, created by, reason code).
- **Data quality**: share of DA units priced with the EAN GSP / PL average / house average / no GSP; DA not mapped to the FCST; IPPL Lines where the DA $ is larger than the FCST (Base FCST < 0).
- Colours: positive DA figures green, negative red (conditional formatting). "n/m" = % not meaningful (FCST ≤ 0 or |%| > 999%).

### 19.5 Reference figures (data of 03-Nov-2026, all filters = All, Nov-26 → Jun-27)
| Indicator | Value |
|---|---|
| FCST final (BASE + INNO + OSP) | 526,034,262 (= Dashboard) |
| Base FCST (without DA) | 431,632,030 |
| DA impact (net) | +94,402,232 (17.9% of FCST, +21.9% vs Base) |
| DA + / DA − | +216,170,205 / −121,767,973 |
| DA units (net) | 2,018,691 |
| DA by layer | BASE +31,189,041 · INNO +45,410,399 · OSP +17,802,792 |
| Top houses | BURBERRY +34.7M · Gucci +32.8M · Kylie Jenner Fragrances +15.8M · Hugo Boss −5.7M |
| Top reason codes | OSP & Initiatives +72.7M · Market Specifics +13.2M · Promo +7.4M · Reasonability/Completeness −5.5M |
| GSP source (DA units) | EAN 16% · PL average 38% · house average 44% · no GSP 2% (Gucci Make up, Infiniment Coty Paris, Orveda) |
| DA not mapped to the FCST | 9 DA lines, 1,293 units (CK One Summer, Orveda Sun / Treatment Kit) |
| IPPL Lines with Base FCST < 0 | 8 (−11.1M in total), e.g. Gucci Guilty Men's Extension, Burberry Her Parfum, Gorgeous Orchid Intense, Goddess Amber Vanilla — mostly priced with PL / house averages |

Other filters: Planner = Adria → FCST 241,439,282, DA +36,494,624; House = Gucci → FCST 103,309,215, DA +32,790,460; Jul-26 → Jun-27 → FCST 908,418,968, DA +94,402,232 (no DA before Nov-26).

### 19.6 Points of attention
- **GSP coverage is low**: only 16% of the DA units have the exact EAN in the GSP tab (the GSP tab has 1,025 EANs, the DA tab 1,203, and 901 DA EANs are missing — mostly make-up). The rest uses averages; adding the missing EANs to the GSP tab improves the result automatically.
- **Base FCST < 0** in some lines means the DA $ is larger than the FCST of the line: either the GSP used is higher than the price behind the FCST $ (averages, new launches) or the DA units do not fully reach the FCST. Check those EANs.
- The layer split (19.3 step 5) is an assumption because the DA has no layer; if the EAN → layer is known in the future it can replace it.
- Macro: `ShowOnlyDashboard` (Workbook_Open) now keeps **Dashboard and DA IMPACT** visible. The VBA source was changed and the compiled cache invalidated, so Excel recompiles it the first time the file is opened.

### 19.7 Validation
An independent Python re-calculation (reading DA, GSP, BASE / INNO / OSP and TABLE directly) was compared with the workbook recalculated in LibreOffice (BASE / INNO / OSP column A frozen to the values saved by Excel, because LibreOffice 24.2 has no XLOOKUP) in 14 filter scenarios (default; planners Adria / Magalie; houses Gucci / BURBERRY; house not of the planner; IPPL Line; INNO only; no INNO; no layer; reason codes; full year; reversed month range): FCST final, Base FCST, DA, DA +, DA −, units, DA by month, by house and by reason code match, and the FCST final equals the Dashboard. No formula errors in the DA IMPACT tab.
