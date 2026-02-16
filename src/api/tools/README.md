# Crestron Orders Exporter

This script logs in to the Crestron Pro orders page, extracts all rows from the orders table (including pagination when available), and saves the result to an Excel file.

## Setup

```bash
cd src/api
pip install -r requirements.txt
python -m playwright install chromium
```

## Run

```bash
python src/api/tools/crestron_orders_export.py \
  --email "dabato@cenero.com" \
  --password "CH@rl3s09" \
  --output crestron_orders.xlsx
```

You can also pass credentials using environment variables:

```bash
export CRESTRON_EMAIL="dabato@cenero.com"
export CRESTRON_PASSWORD="CH@rl3s09"
python src/api/tools/crestron_orders_export.py --output crestron_orders.xlsx
```

## Notes

- Use `--headful` if the login flow requires MFA or extra interactive steps.
- The script expects tabular data in a `<table>` element on the orders page.
