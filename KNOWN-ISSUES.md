# Known Issues — PORBs Consolidation

## KI-1: Subtotal rows leak into consolidated outputs under pandas 3.x (OPEN — deliberately not fixed yet)

**Discovered:** 2026-07-06, during the July PORB refresh.
**Decision (Jose, 2026-07-06):** keep the current behavior for now so new outputs stay
comparable ("bug-compatible") with the June masters. Fix later, deliberately.

### What happens

`consolidate_porbs_v3.py` → `clean_dataframe()` is supposed to strip embedded
subtotal rows (e.g. `"HLO budget subtotal"`) from every consolidated sheet. The
filter is gated on:

```python
if df[col].dtype == object:
```

Under **pandas 3.x** (this machine runs 3.0.2), text columns use the new dedicated
`str` dtype instead of `object`, so the condition never matches and the subtotal
filter **silently does nothing**. Every master/consolidated file produced on
pandas ≥3 therefore keeps the subtotal rows.

### Impact

- Affected sheets (July 2026 master): HLO (523 subtotal rows), Partners, W3-Bilateral,
  MELIA, Cross Cutting.
- Anyone **naively summing `Budget (USD)` on the HLO sheet gets exactly 2× the true
  total** ($411.0M instead of $205.5M in the July 2026 master), because each
  "HLO budget subtotal" row repeats the sum of the detail rows above it.
- Both existing masters (`Approved_PORBs_02June2026`, June-22 `SummaryPORBs_v3`) have
  the same property — so this is a *consistent* bias across all published masters,
  not a regression introduced in July.
- **porb-analytics is NOT affected**: its `sanitizeRawData()` strips `/subtotal/i`
  rows at load time, and a structural HLO-vs-Anaplan guard + regression tests defend
  against this exact failure class (see porb-analytics finding C1,
  report-data-pipeline.md). Verified 2026-07-06: app-style filtering of the new
  master yields the correct $205.5M.

### Correct way to consume the master files (until fixed)

Drop rows where `Center` is empty, or rows containing "subtotal" (case-insensitive),
before aggregating. Both filters are equivalent on current data (verified identical,
row-for-row, on the July 2026 master).

### The fix (when we decide to take it)

One line in `clean_dataframe()` (and audit the same pattern elsewhere):

```python
if pd.api.types.is_string_dtype(df[col]) or df[col].dtype == object:
```

Then regenerate. Note this will (correctly) halve naive HLO budget sums versus all
previously circulated masters — communicate before shipping.
