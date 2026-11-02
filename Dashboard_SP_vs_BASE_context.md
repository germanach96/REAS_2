# Context document: "SP vs BASE" Dashboard (file REAS_2_301026_dashboard.xlsm)

> **Instructions for any AI assistant reading this document**
> - This is a self-contained, plain-text description of an Excel demand-planning dashboard (Coty, Barcelona Supply Chain Hub). It does not depend on any specific AI tool, plugin or file access.
> - Treat it as the source of truth to answer questions about what the dashboard shows, where the data comes from, how each figure is calculated, how to update it, its limitations and how to interpret the results.
> - If a question is not covered here, say so explicitly instead of guessing. Do not invent figures: the only numbers you can quote are the ones in this document (section 11) or numbers the user gives you.
> - The dashboard was built with the help of an AI assistant and validated by recalculating the figures independently.

---

## 1. Purpose

The dashboard compares the **Sales Plan (SP)** with the **BASE forecast** (system forecast volume) plus **innovation volume (INNO)**, by house, IPPL Line and product line (PL), month by month, for fiscal year FY27 (Jul-26 → Jun-27).

It also shows:
- **Sell-in trends** (BASE + INNO volume, last 6 and 12 closed months vs last year).
- **Sell-out trends** (EPOS in $, last 26 and 52 weeks vs last year).
- **LY** (last-year volume) and **trend projections** (LY × (1 + trend)).
- **Forecast accuracy at lag 3**: **SPP3** (bias) and **WAPE3** (weighted absolute percentage error).
- An **EXPORT** button that produces a 5-slide PDF for reporting.

Typical use: check whether the plan (SP) is aligned with what the system forecast and the innovation plan say, and support the discussion with trends and accuracy.

---

## 2. File structure

The file is a **macro-enabled Excel workbook (.xlsm)**. By default **only the Dashboard tab is visible**; all other tabs are hidden (right-click a tab → Unhide…).

| Tab | Type | Content |
|---|---|---|
| **SP** | Manual input | Sales Plan: House, IPPL Line and 12 months (row 1 = months Jul-26…Jun-27). Data is pasted from row 2 (rows 2–201 are read). |
| **BASE** | Manual input | System extract (formerly called "REAS"). One row per product line; column A = IPPL LINE (XLOOKUP formula on IPPL MAP 1), column B = Product_line; column blocks per FY / half / quarter / month with measures (Act + Fcst, Year Ago, % vs YA, Snapshot, % L12M vs YA, % L6M vs YA, SPP3 %, WAPE3, Cuts…). Rows 5–504 are read. |
| **INNO** | Manual input | Same structure as BASE, innovations only. Rows 5–504 are read. |
| **EPOS** | Manual input | Sell-out (Power BI export): Category, House, Brand, Line; $ and YoY Δ / YoY % for L26W and L52W. Rows 3–602 are read. (To be reviewed later.) |
| **IPPL MAP 1 / IPPL MAP 2** | Master data | Mapping GMM Line / Product line → IPPL Line and House. |
| **TABLE** | Automatic (formulas only) | Consolidates SP, BASE and INNO by IPPL Line and month. Never typed into. |
| **Dashboard** | Visible | Filters, KPIs, tables, charts and EXPORT button. |
| **DASH CALC** | Automatic (hidden) | Calculation engine of the dashboard. Do not edit. |
| **EXPORT** | Automatic (hidden) | Layout of the 5 slides exported to PDF. |

**Units:** BASE and INNO come in thousands in the system; TABLE and the dashboard multiply them × 1000 so they are comparable with SP (volume in units). EPOS is in $ and is only used for % trends.

---

## 3. How TABLE works (automatic consolidation)

- **IPPL Line list:** built automatically as the de-duplicated union of the IPPL Lines in SP (first, in SP order) and then those that exist only in BASE or INNO. Capacity: 150 lines.
- **House of each line:** taken from SP; if the line is not in SP, looked up in IPPL MAP 1; if not found there either, it is labelled **"(no house)"** (only visible with House = All).
- **Blocks:** rows 2–151 = SP, 152–301 = BASE, 302–451 = INNO (column A = TYPE).
- **Values:**
  - SP = sum of SP for that IPPL Line and month. If the month is empty in SP it stays empty (not 0); e.g. Jul–Sep 26 have no SP.
  - BASE = sum of the month's "Act + Fcst" column in the BASE tab for that IPPL Line × 1000.
  - INNO = same from the INNO tab × 1000.
- **Adding data:** new houses, IPPL Lines or product lines added in SP, BASE or INNO appear automatically in TABLE, in the dropdowns and in every calculation (tested with a test house, a test IPPL Line and test PLs).

Monthly "Act + Fcst" columns used in BASE/INNO: Jul-26 CP, Aug-26 DC, Sep-26 DP, Oct-26 EC, Nov-26 EP, Dec-26 FC, Jan-27 FP, Feb-27 GC, Mar-27 GP, Apr-27 HC, May-27 HP, Jun-27 IC. The Year Ago (YA) column of each month is the next column (CQ, DD, DQ, ED, EQ, FD, FQ, GD, GQ, HD, HQ, ID).

---

## 4. Dashboard filters (yellow cells)

| Filter | Cell | Options | Effect |
|---|---|---|---|
| House | C5 | All + existing houses | Filters everything. |
| IPPL Line | C6 | All + lines of the selected house | If the line does not belong to the house, a warning appears and all lines of the house are shown. |
| Remove INNO | C7 | No / Yes | Changes the comparison basis (see §5). |
| Month from / Month to | C8 / C9 | Jul-26…Jun-27 | Range for totals, KPIs and "period total" details. Default Oct-26 → Jun-27 (SP has no volume in Jul–Sep 26). |
| Metric for per-line charts | C10 | BASE / SP / Gap | Metric of the monthly history by IPPL Line chart. |
| Product line | C11 | All + PLs within the House/IPPL filter | Product-line view (see §8). If the PL does not fit the House/IPPL filter, a warning appears and everything is shown. |

Trends and forecast accuracy **do not depend** on Month from / Month to (they always use the last closed months).

---

## 5. Comparison logic and "Remove INNO"

| Remove INNO | SP side | BASE side |
|---|---|---|
| **No** (default) | SP | **BASE + INNO** |
| **Yes** | **SP − INNO** | **BASE** |

Reason: SP already includes innovations; the system forecast keeps them separately in INNO. With "Yes" both sides are compared without innovation.

**No-negatives rule (when INNO is removed):** SP and the INNO plan are planned by different people, so they may not match. Therefore, with Remove INNO = Yes:
- SP − INNO is calculated **per IPPL Line and per month** (the lowest level at which SP exists) and **floored at 0**: MAX(0, SP − INNO). Only then is it summed to house / total.
- BASE never goes below BASE (removing INNO leaves exactly BASE).
- Example: SP 100, INNO 120, BASE 20 → SP − INNO = 0 and BASE = 20.
- Consequence: the volume gap can differ between the two modes when, in some line/month, INNO exceeds SP.

At product-line level there is no SP, so the subtraction does not apply (only BASE or BASE + INNO).

---

## 6. What the Dashboard shows

1. **Period KPIs:** SP, BASE (+ INNO), Gap (BASE vs SP) and Gap % = BASE / SP − 1, plus INNO in the range.
2. **Trend tiles** (traffic light): Sell-in 6M, Sell-in 12M, Sell-out 6M, Sell-out 12M. Green ≥ +5%, amber between −5% and +5%, red ≤ −5%.
3. **Monthly table:** SP, BASE (+ INNO), Gap, Gap % (colour scale: red = BASE below SP, white ≈ 0, green = BASE above SP; full colour at ±30%) and INNO; "Range total" column.
4. **SP vs BASE by month chart** (lines). With switches (panel "Extra lines on the chart"):
   - **LY (last year):** sum of the Year Ago columns of BASE + INNO (BASE only if Remove INNO = Yes).
   - **Projections:** LY × (1 + trend) for Sell-in 6M, Sell-in 12M, Sell-out 6M, Sell-out 12M and a **manual trend** (cell R28). Several can be shown at once. The panel shows the applied trend and the total in the range.
5. **BASE vs INNO breakdown:** stacked bars BASE + INNO by month with the SP line. With Remove INNO = Yes, INNO disappears and the line becomes SP − INNO.
6. **Forecast accuracy (lag 3):** SPP3 bars and WAPE3 line for closed months (Jul, Aug, Sep 26), with dotted thresholds WAPE 40% and SPP ±5%.
7. **Detail by IPPL Line** (period total, SP vs BASE) and **by product line** (top 25 by volume).
8. **Monthly history by IPPL Line** (one line per IPPL Line, top 30 by SP; metric per filter C10) and **by product line** (top 25).
9. **Sell-in vs sell-out trends by IPPL Line** (chart) and tables by IPPL Line and by product line (all PLs in the filter, sorted, with trends).

---

## 7. Key definitions and calculations

### 7.1 Sell-in trends (BASE + INNO, volume)
- Windows: **L6M = Apr-26 → Sep-26**, **L12M = Oct-25 → Sep-26** (last closed month: Sep-26). Apr–Jun 26 and Oct-25–Jun-26 come from the Year Ago columns of the equivalent FY27 months; Jul–Sep 26 from "Act + Fcst".
- The last-year volume of each row is derived from the % provided by the system: **YA = Act / (1 + "% L6M vs YA")** (or L12M). Each row therefore reproduces the system % exactly and the aggregate is volume-weighted: **trend = ΣAct / ΣYA − 1**.
- Products with no history (launches) count with YA = 0; if the whole selection is new, the trend shows **n/a**.
- With Remove INNO = Yes, BASE only.

### 7.2 Sell-out trends (EPOS, $)
- 6M = **L26W**, 12M = **L52W**. **trend = Σ$ / Σ($ − YoY Δ) − 1.**
- Each EPOS line is mapped to its IPPL Line via IPPL MAP 1 (all current lines map).
- Validated: by house and in total it matches exactly the YoY % calculated by Power BI in the EPOS tab.
- At product-line level it is matched by exact line name in EPOS; if there is no match, n/a.
- Remove INNO does not affect sell-out (EPOS does not split INNO).

### 7.3 Forecast accuracy (lag 3)
- **SPP3 = Actual / Forecast lag3 − 1** (negative = over-forecast; positive = under-forecast).
  - The formula was inferred from the data: SPP3 ranges from −100% to +999% (only consistent with Act/Fcst − 1), and in 216 single-SKU rows WAPE matches exactly |Act − Fcst| / Act using that forecast.
- **WAPE3 = Σ|Actual − Forecast lag3| / ΣActual.**
- **Aggregation (percentages are never summed):** for each row and month the dashboard rebuilds **Forecast = Actual / (1 + SPP3)** and **|Error| = WAPE3 × Actual**, sums them and recalculates the % for the filter. Validated: filtering a single PL returns exactly its values from the tab (e.g. Gorgeous Gardenia Sep-26: SPP3 −12.9%, WAPE3 16.1%).
- Closed months with results only (currently Jul, Aug, Sep 26). Rows with Actual = 0 or without KPI cannot be rebuilt and are excluded (small volume).
- Reference thresholds: WAPE 40%, SPP ±5%.

### 7.4 LY and projections
- Monthly LY = Σ Year Ago (BASE + INNO, or BASE only with Remove INNO) × 1000, filtered.
- Projection = LY × (1 + selected trend).

---

## 8. Product-line view (filter C11)
- SP does not exist at PL level: it is hidden (KPIs "n/a", no SP line in charts) and a notice appears.
- BASE (+ INNO), LY, projections, sell-in / sell-out trends and SPP3 / WAPE3 are calculated for that PL only.
- Charts by IPPL Line keep showing the full IPPL Line as a reference.

---

## 9. EXPORT button (macro)

- Located top right of the Dashboard. Macros must be enabled when opening the file (and, if the file comes from email/SharePoint, it may need to be unblocked in File Properties).
- Produces a **5-slide PDF** (A4 landscape, one slide per page) **without asking**, named **REAS_2_EXPORT_YYMMDDHHMMSS.pdf** (always unique), and opens it when done.
- **Location:** same folder as the Excel file. If the file is on OneDrive/SharePoint, the macro finds the locally synced folder via the OneDrive registry; if that fails it tries the web path; as a last resort it saves to Downloads and shows the path.
- **Slide content** (each slide has the title, the **active filters** and what is shown at the top, and a **COMMENTS** box at the bottom for notes):
  1. KPIs + trend tiles + monthly table + SP vs BASE chart (with the LY/projection lines switched on).
  2. BASE vs INNO breakdown + forecast accuracy (SPP3/WAPE3).
  3. Detail by IPPL Line + detail by product line.
  4. Monthly history by IPPL Line.
  5. Monthly history by product line.
- Not exported (they remain in the dashboard): the sell-in vs sell-out trends by IPPL Line chart and the two final tables (by IPPL Line and by product line).
- Charts keep the same size and proportions as in the dashboard.

---

## 10. How to update the data
1. Unhide the relevant tab (right-click the Dashboard tab → Unhide…).
2. **SP:** paste House, IPPL Line and the 12 months from row 2 (row 1 = months).
3. **BASE / INNO:** paste the new extract with the same column structure; make sure the XLOOKUP formula in column A (IPPL LINE) reaches every row.
4. **EPOS:** paste the Power BI export with the same structure.
5. Hide the tabs again if desired. Everything else (TABLE, DASH CALC, Dashboard, EXPORT) recalculates automatically.

Capacities: SP 200 rows, BASE 500, INNO 500, EPOS 600, 150 IPPL Lines, 30 houses. Charts by IPPL Line show the top 30 by SP.

---

## 11. Reference figures (data of 30-Oct-2026, House = All, Oct-26 → Jun-27)

| Indicator | Remove INNO = No | Remove INNO = Yes |
|---|---|---|
| SP / SP − INNO | 144,391,730 | 102,874,741 |
| BASE + INNO / BASE | 151,865,659 | 104,886,170 |
| Gap | 7,473,929 | 2,011,429 |
| Gap % | +5.2% | +2.0% |
| INNO in range | 46,979,489 | 46,979,489 |
| Sell-in 6M / 12M | +20.4% / +15.5% | +4.5% / +19.5% |
| Sell-out 6M / 12M (EPOS) | +7.7% / +3.5% | +7.7% / +3.5% |
| SPP3 Jul / Aug / Sep 26 | +0.3% / −8.1% / −9.9% | 0.0% / −19.8% / −13.2% |
| WAPE3 Jul / Aug / Sep 26 | 44.7% / 50.1% / 37.0% | 38.2% / 42.7% / 41.4% |
| LY (Oct-26 → Jun-27) | 152,057,210 | 108,052,009 |

Other examples (Remove INNO = No): Gucci: SP 74,285,617 vs BASE + INNO 75,893,746 (+2.2%). Kylie Makeup / All Kylie Eyes: SP 678,871 vs 839,828 (+23.7%).

Compared with the 28-Oct version, BASE and INNO changed (forecast of future months; total BASE + INNO went from 154.16M to 151.87M). SP and EPOS did not change, and neither did trends/accuracy (Sep-26 is still the last closed month).

---

## 12. Validation performed
- SP, BASE and INNO totals match a direct consolidation of the input tabs in every scenario tested (All, by house, by IPPL Line, by PL, with and without INNO).
- Sell-out by house = Power BI YoY % (exact).
- Sell-in and SPP3/WAPE3 recalculated independently row by row: they match.
- The restructuring into SP/BASE/INNO inputs + automatic TABLE reproduces exactly the previous figures, and was tested by adding a new house, IPPL Line and PLs.
- The exported slides show the same values as the Dashboard.

---

## 13. Limitations and points of attention
- SP exists only by IPPL Line (not by PL).
- SP has no volume in Jul–Sep 26; that is why the default range starts in Oct-26.
- Trends and accuracy rely on the % provided by the system; if the system changes its definitions they must be reviewed.
- System Q1/FY values for SPP3/WAPE3 are not used: only closed months are shown.
- BASE/INNO rows whose product line has no IPPL Line in IPPL MAP 1 are not assigned to any line.
- Some chart legends show empty entries when there are fewer lines than slots (Excel limitation without extra macros).
- EPOS will be reviewed later (structure and mapping).

---

## 14. Next steps / improvements

1. **Consider cuts in trends and actuals.** The BASE/INNO extracts include "o9 Cuts" and "Cuts YA" columns that are not used today. Next step: add the cut volume back to actuals (and year-ago actuals) so that sell-in trends, LY and accuracy reflect unconstrained demand rather than shipped volume after cuts, with an option to show both views.
2. **Identify forecast layers.** Split the forecast into its layers (e.g. statistical baseline, promotions/events, innovation, management adjustments) so the gap vs SP can be explained by layer instead of only as a total.
3. **Identify past promotional or exceptional activities.** Flag months/products affected by promotions, one-off deals, launches, stock-outs or other exceptional events, so that trends, LY projections and accuracy can be read (or recalculated) excluding those effects.
4. **Flexibility to add more houses and PLs to the scope.** The input tabs and TABLE already pick up new houses, IPPL Lines and PLs automatically within the current capacities (SP 200 rows, BASE/INNO 500 rows each, 150 IPPL Lines, 30 houses). Next step: extend the scope to more houses/brands, keep IPPL MAP 1 complete for every new product line, raise capacities if needed and review chart readability when many lines are selected.

---

## 15. Frequently asked questions (short answers)

**What is BASE?** The system forecast by product line ("Act + Fcst" column), formerly called REAS. It contains actuals in closed months and forecast in future months.

**Why is SP compared with BASE + INNO?** Because SP includes innovations and BASE does not. To compare without innovations use Remove INNO = Yes (SP − INNO vs BASE).

**Can SP − INNO be negative?** No: it is floored at 0 per IPPL Line and month.

**What does a positive Gap % mean?** BASE (+ INNO) is above SP: the system forecast expects more volume than the plan.

**What does a negative SPP3 mean?** Over-forecast: actuals came in below what was forecast 3 months earlier.

**Why don't trends change when I move Month from/to?** They always measure the last 6/12 closed months (to Sep-26) vs last year.

**Why does a trend show n/a?** There is no last-year volume (new product or line) or no match in EPOS.

**How is the projection calculated?** LY of the month × (1 + selected trend); the trend can be sell-in, sell-out or manual.

**Where is the exported PDF saved?** In the same folder as the Excel file (or in Downloads, with a message, if that folder cannot be written). Name REAS_2_EXPORT_YYMMDDHHMMSS.pdf.

**What do I do if I add a new house or line?** Just paste it in SP/BASE/INNO; it appears automatically in TABLE, dropdowns and charts.

**Why can't I see the data tabs?** They are hidden on purpose; right-click a tab → Unhide.

---

## 16. Glossary
- **SP:** Sales Plan.
- **BASE:** system forecast by product line (formerly REAS).
- **INNO:** innovation volume.
- **IPPL Line:** product planning group (the level at which SP exists).
- **PL:** Product line.
- **LY / YA:** Last Year / Year Ago (same month of the previous year).
- **L6M / L12M:** last 6 / 12 closed months.
- **L26W / L52W:** last 26 / 52 weeks (sell-out).
- **EPOS:** point-of-sale sell-out data ($).
- **SPP3:** forecast bias at lag 3 (Act / Fcst − 1).
- **WAPE3:** weighted absolute percentage error at lag 3 (Σ|Act − Fcst| / ΣAct).
- **Lag 3:** forecast made 3 months before the month being measured.
- **Cuts:** volume that could not be served (cut) versus the original demand.
