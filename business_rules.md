# Business Rules

This document records the current calculation rules as implemented in the codebase.
It is intentionally descriptive, not aspirational.

## 1. Shipment profit
- `base_profit = customer_net - platform_cost`
- Implemented in `scripts/build_report.py` inside `collect_shipments()`
- Example:
  - customer net: 80.00
  - platform cost: 55.00
  - base profit: 25.00

## 2. Tax deduction
- Where applicable, gross values are reduced by 15%.
- Implemented in `scripts/build_report.py` inside `read_prices()`
- Also reflected in `web/generate_site.py` for return calculations.
- Example:
  - gross 100.00 -> net 85.00

## 3. Overweight threshold
- Overweight starts after 15 kg.
- Implemented in `scripts/build_report.py` inside `collect_shipments()`
- Example:
  - weight 17.5 kg -> extra 2.5 kg

## 4. Overweight profit
- `extra_profit = extra_kg * extra_profit_net`
- Implemented in `scripts/build_report.py` inside `collect_shipments()`
- Example:
  - extra kg: 2
  - net extra kilo profit: 4.00
  - overweight profit: 8.00

## 5. COD profit
- `cod_profit = customer_cod_fee - platform_cod_cost`
- Implemented in `scripts/build_report.py` inside `read_prices()` and `collect_shipments()`
- Example:
  - customer COD fee: 10.00
  - platform COD cost: 3.00
  - COD profit: 7.00

## 6. Excluded shipments
- Shipments with status `ملغي` or `مسودة` are excluded from profit.
- Implemented in `scripts/build_report.py` inside `collect_shipments()`
- Example:
  - status `ملغي` -> included in totals? no

## 7. June-only scope
- The current report pipeline only includes June records.
- Implemented in `scripts/build_report.py` inside `collect_shipments()`
- Example:
  - 2026-05-30 -> excluded
  - 2026-06-02 -> included

## 8. Special merchants
- `عبدالرحمن المطيري`
- `مؤسسة اواني القيصرية التجارية`
- These merchants have special pricing handling.
- Implemented in `scripts/build_report.py` inside `customer_price_from_sheet()` and `platform_base_cost()`
- Example:
  - customer gross 100.00 may remain 100.00 instead of being reduced

## 9. Customer price lookup order
- First try merchant-specific pricing.
- If missing, try carrier pricing.
- If still missing, use the shipment row fallback.
- Implemented in `scripts/build_report.py` inside `customer_price_from_sheet()`

## 10. Platform cost lookup order
- First try merchant-specific pricing.
- If missing, try carrier pricing.
- Implemented in `scripts/build_report.py` inside `platform_base_cost()`

## 11. Unique shipment identity
- Shipments are deduplicated by `order_id` in the report pipeline.
- Implemented in `scripts/build_report.py` inside `collect_shipments()`
- The sync path also upserts by the same primary key logic in `scripts/sync_from_lead.py`.

## 12. Totals
- `total = base_profit + extra_profit + cod_profit`
- Implemented in `scripts/build_report.py` inside `collect_shipments()` and `aggregate()`
- Example:
  - base 20 + extra 8 + COD 7 = total 35

## 13. Revenue card
- Dashboard revenue is labeled `إجمالي الإيرادات` and described as `إجمالي ما دفعه العملاء`.
- Implemented in `web/generate_site.py`

## 14. Return profit
- Returns are handled separately in the dashboard logic.
- Implemented in `web/generate_site.py` inside `return_profit_from_details()`

## 15. Financial operations categories
- Bank deposits
- Moyasar deposits
- Shipping return deduction
- Tax deduction
- Shipping refund
- Implemented in `scripts/build_report.py` inside `style_operations_sheet()`
- Reflected in `web/generate_site.py`

## 16. Last updated
- The dashboard footer uses the modified time of `output/lead6_report.xlsx`.
- Implemented in `web/generate_site.py`

## 17. June period label
- The visible period is `1` to the last day present in the data.
- Implemented in `scripts/build_report.py` and `web/generate_site.py`

## 18. Missing sequence numbers
- Missing order numbers are calculated between the minimum and maximum numeric order IDs found in the data.
- Implemented in `web/generate_site.py`

## 19. Ranking lists
- Top merchants, cities, and carriers are sorted by total profit descending.
- Implemented in `web/generate_site.py`

## 20. Current state
- PostgreSQL is being introduced in parallel only.
- Excel remains the operational output until validation is complete.

