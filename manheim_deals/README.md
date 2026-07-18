# Manheim MMR Deal Finder

Reads a dealer inventory spreadsheet, looks up the **Manheim Market Report (MMR)**
wholesale value for each vehicle by VIN + mileage, compares it to the asking
price, and ranks the best buying opportunities.

## Security first

- Your Manheim **API credentials** (`client_id` / `client_secret`) live only in a
  local `.env` file. They are never hardcoded and never committed (`.gitignore`
  covers `.env`).
- You need Manheim **API** credentials from the Manheim / Cox Automotive developer
  portal — this is *not* your website login/password.

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
# edit .env and paste your MANHEIM_CLIENT_ID / MANHEIM_CLIENT_SECRET
```

## Run

```bash
# 1) Sanity check — parse the sheet, no API calls, no billing:
python manheim_deals.py "Enterprise_DR_FM_Trade_Inventory_List_7.17.26.xlsx" --dry-run

# 2) Test on the first 5 vehicles once credentials are set:
python manheim_deals.py inventory.xlsx --limit 5

# 3) Full run — flag anything >=15% under MMR, skip damage/red-light units:
python manheim_deals.py inventory.xlsx --min-margin 0.15 --exclude-risky
```

## Output

- **Console** — top opportunities, sorted by margin.
- **deals_ranked.xlsx** — every vehicle ranked, with an `IsDeal` flag.
- **mmr_cache.json** — cached API responses so re-runs don't re-bill the API.

## Options

| Flag | Meaning |
|------|---------|
| `--min-margin 0.15` | Flag as a deal when `(MMR - ask) / MMR >= 0.15`. Default `0.12`. |
| `--exclude-risky` | Skip Damage Units and Trade Units (red-light, non-returnable). |
| `--dry-run` | Parse only; no API calls. |
| `--limit N` | Only process the first N vehicles (testing). |
| `--sleep 0.2` | Seconds between API calls (rate limiting). |

## Important note on the Manheim endpoint

The MMR base URL and valuation path can vary by account/region. If lookups fail
with 404/401, open `manheim_deals.py` and confirm `MANHEIM_TOKEN_URL` /
`MANHEIM_MMR_URL` (and the `odometer` vs `mileage` query param) against what your
Manheim developer portal documents — they're isolated at the top of the file and
can also be overridden in `.env`.
