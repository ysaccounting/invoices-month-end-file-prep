# Invoices Processor

A web app to combine, classify, and split invoice Excel reports into Y&S and Non Y&S exports.

## Features

- **Drag & drop** multiple `.xlsx` invoice files
- Combines all files into one dataset
- Adds **Main Company** column (mappings: YS-Seatgeek/YS-Seatgeek2 → YS Tickets, YSA 2/YSA 3 → YSA)
- Splits output into two files based on company classification:
  - `Invoices {date} (YS).xlsx` — tab named **Invoices**
  - `Invoices {date} (Non YS).xlsx` — tab named **Invoices**
- Formats money columns (Amnt, Cost, Bal., Payout, Payout Balance, TV Fee) with `#,##0.00`
- **Clear** button resets all inputs and outputs

## Local Development

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

## Deploy to Railway

1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub repo
3. Select your repo — Railway auto-detects and deploys

## Company Mapping

| Company | Classification |
|---|---|
| Damon and Crew | Non Y&S |
| The Ticket Guy | Non Y&S |
| YourTickets | Non Y&S |
| GK LLC, Jacks YS, Levovitz, Needle Tickets LLC, Pollak Tickets, Yoni Levine, YS Katz, YS Tickets, YS TL, YSA, YSA 2, YSA 3, YSM Tickets, YSS Tickets, YS-Seatgeek, YS-Seatgeek2, YSW | Y&S |

### Main Company overrides
- `YS-Seatgeek` / `YS-Seatgeek2` → `YS Tickets`
- `YSA 2` / `YSA 3` → `YSA`
