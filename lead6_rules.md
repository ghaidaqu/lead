# lead6 rules

Use this workbook logic for the June dataset.

## Scope
- Process only June 1 to June 12.
- Use the raw Excel file as the source.
- Keep the output organized and clean like the last version.

## Profit rules
- Customer shipping price: take from the prices sheet.
- Platform shipping cost: take from the prices sheet.
- Deduct 15% tax where applicable.
- Overweight starts after 15 kg.
- Overweight profit is based on the extra kilos above 15.
- COD profit is customer COD fee minus platform COD cost.

## Exceptions
- For `عبدالرحمن المطيري` and `مؤسسة اواني القيصرية التجارية`:
  - Customer shipping price stays as-is, without tax deduction.
  - Platform shipping cost stays as-is, with tax included.

## Required sheets
- `Dashboard`
- `ملخصات`
- `تفاصيل شهر 6`
- `الاسعار`
- `العمليات المالية`

## Remove
- Remove `الشحنات` from the final workbook.

## Dashboard
- Keep the dashboard as a separate local host page.
- Keep the Excel file linked from the dashboard.
- Show only the latest clean layout.
