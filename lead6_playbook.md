# lead6 playbook

Use this exact workflow when a new raw Excel file is provided.

## Inputs
- Raw workbook: the file the user sends.
- Use only June data.
- Scope: June 1 to June 12.

## Core rules
- Read shipping prices from the `الاسعار` sheet.
- Customer price: use the price from `الاسعار`.
- Platform cost: use the price from `الاسعار`.
- Apply 15% tax deduction where applicable.
- Overweight starts after 15 kg.
- Overweight profit is based on kilos above 15.
- COD profit = customer COD fee - platform COD cost.
- Special handling:
  - `عبدالرحمن المطيري`
  - `مؤسسة اواني القيصرية التجارية`
  - For these two, customer shipping price stays as-is.
  - For these two, platform shipping cost is taken with tax included.

## Required workbook structure
- `Dashboard`
- `ملخصات`
- `تفاصيل شهر 6`
- `الاسعار`
- `العمليات المالية`

## Cleanup
- Remove `الشحنات` from the final workbook.
- Keep the workbook focused on the June 1-12 window.
- Add the missing-serials note in the first sheet if gaps exist.

## Dashboard layout
- Title: `شهر 6`
- Keep the dashboard separate from the Excel file.
- Include:
  - total profit
  - shipping profit
  - overweight profit
  - COD profit
  - last 7 days profit chart
  - last 7 days shipment count chart
  - finance cards: bank, Moyasar, total deposits
  - carrier profit circles
  - COD table
- Remove `أفضل العملاء`.

## Financial operations
- Separate bank and Moyasar deposits.
- Show total deposits in summary and dashboard.
- If a bank statement PDF is attached, add a `كشف الحساب` sheet with categorized deposits and expenses, keep the statement summary visible, and mirror the expense totals in the dashboard.

## Deliverables
- Clean Excel report in the project output folder.
- Dashboard HTML and workbook copy in the project web folder.
- Keep the local dashboard in sync with the report.

## Reference phrase
When the user sends a new raw workbook, treat the request as:
`طبّق lead6 playbook على الملف الخام`
