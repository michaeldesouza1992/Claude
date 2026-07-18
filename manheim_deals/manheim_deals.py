#!/usr/bin/env python3
"""
Manheim MMR deal finder.

Reads a dealer inventory spreadsheet, looks up the Manheim Market Report (MMR)
wholesale value for each vehicle by VIN + mileage, compares it to the asking
price, and ranks the best buying opportunities.

Credentials are read from a local .env file (see .env.example). Your Manheim
API client_id / client_secret NEVER get hardcoded or shared.

Usage:
    python manheim_deals.py INVENTORY.xlsx
    python manheim_deals.py INVENTORY.xlsx --min-margin 0.15 --exclude-risky
    python manheim_deals.py INVENTORY.xlsx --dry-run        # parse only, no API

Output:
    - Console: top opportunities
    - deals_ranked.xlsx : full ranked table
    - mmr_cache.json    : cached MMR responses (so re-runs don't re-bill the API)
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import Optional

import requests

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency. Run: pip install -r requirements.txt")

# --- Load .env (no external dependency needed) --------------------------------
def load_dotenv(path: str = ".env") -> None:
    p = Path(path)
    if not p.exists():
        return
    for line in p.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


# --- Manheim API config -------------------------------------------------------
# NOTE: Confirm these against YOUR Manheim developer portal — base URLs and the
# exact valuation path can differ by account/region. They're isolated here so
# you only edit one place if your portal shows something different.
MANHEIM_TOKEN_URL = os.environ.get(
    "MANHEIM_TOKEN_URL", "https://api.manheim.com/oauth2/token"
)
MANHEIM_MMR_URL = os.environ.get(
    "MANHEIM_MMR_URL", "https://api.manheim.com/valuations/vin/{vin}"
)


class ManheimClient:
    """Thin OAuth2 client-credentials wrapper for the Manheim MMR API."""

    def __init__(self, client_id: str, client_secret: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self._token: Optional[str] = None
        self._token_expiry: float = 0.0
        self.session = requests.Session()

    def _ensure_token(self) -> str:
        if self._token and time.time() < self._token_expiry - 60:
            return self._token
        resp = self.session.post(
            MANHEIM_TOKEN_URL,
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
            },
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        self._token = data["access_token"]
        self._token_expiry = time.time() + int(data.get("expires_in", 3600))
        return self._token

    def mmr_by_vin(self, vin: str, mileage: Optional[int]) -> dict:
        token = self._ensure_token()
        url = MANHEIM_MMR_URL.format(vin=vin)
        params = {}
        if mileage:
            params["odometer"] = mileage  # some portals use 'mileage' — adjust if needed
        resp = self.session.get(
            url,
            params=params,
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()


def extract_mmr_value(payload: dict) -> Optional[int]:
    """Pull the adjusted wholesale average out of a Manheim MMR response.

    Manheim's schema nests the number in a few possible shapes depending on
    account. We probe the common ones and fall back to a recursive search for a
    'wholesale'/'average' figure.
    """
    # Common shape: {"items":[{"wholesale":{"average": 22350, ...}}]}
    def dig(obj, keys):
        cur = obj
        for k in keys:
            if isinstance(cur, list):
                cur = cur[0] if cur else None
            if not isinstance(cur, dict):
                return None
            cur = cur.get(k)
        return cur

    for path in (
        ["items", "wholesale", "average"],
        ["wholesale", "average"],
        ["mmr", "wholesale", "average"],
        ["adjustedMMR", "wholesale", "average"],
        ["baseMMR", "wholesale", "average"],
    ):
        val = dig(payload, path)
        if isinstance(val, (int, float)):
            return int(val)

    # Last resort: recursive search for a plausible 'average' under 'wholesale'.
    found = []

    def walk(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k.lower() in ("average", "adjustedaverage") and isinstance(v, (int, float)):
                    found.append(int(v))
                walk(v)
        elif isinstance(o, list):
            for x in o:
                walk(x)

    walk(payload)
    return found[0] if found else None


# --- Inventory parsing --------------------------------------------------------
# Column layout of the source spreadsheet (1-indexed):
COL = {
    "unit": 1, "year": 2, "make": 3, "model": 4, "series": 5, "equip": 6,
    "miles": 7, "color": 8, "location": 9, "vin": 12, "price": 13,
}

# Sections considered higher risk (can't be returned / damaged).
RISKY_SECTIONS = {"Damage Units", "Trade Units"}
SECTION_KEYWORDS = ("Certified Units", "Daily Rental", "FM Units", "Trade Units", "Damage Units")


@dataclass
class Vehicle:
    unit: str
    year: Optional[int]
    make: str
    model: str
    series: str
    miles: Optional[int]
    color: str
    location: str
    vin: str
    ask_price: Optional[int]
    section: str
    # filled after lookup:
    mmr: Optional[int] = None
    spread: Optional[int] = None       # mmr - ask_price (positive = under market)
    margin: Optional[float] = None     # spread / mmr
    note: str = ""


def _as_int(v) -> Optional[int]:
    try:
        if v is None or v == "":
            return None
        return int(float(str(v).replace(",", "").replace("$", "").strip()))
    except (ValueError, TypeError):
        return None


def parse_inventory(xlsx_path: str) -> list[Vehicle]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    vehicles: list[Vehicle] = []
    section = "Unknown"

    for r in range(1, ws.max_row + 1):
        def cell(name):
            return ws.cell(row=r, column=COL[name]).value

        a = cell("unit")
        vin = cell("vin")

        # Section header row: text in col A, no VIN.
        if a and not vin:
            text = str(a).strip()
            for kw in SECTION_KEYWORDS:
                if kw.lower() in text.lower():
                    section = kw
                    break
            continue

        # Repeated column-header rows.
        if vin in (None, "", "VIN"):
            continue
        if str(a).strip().lower() == "unit":
            continue

        vehicles.append(
            Vehicle(
                unit=str(a).strip() if a else "",
                year=_as_int(cell("year")),
                make=str(cell("make") or "").strip(),
                model=str(cell("model") or "").strip(),
                series=str(cell("series") or "").strip(),
                miles=_as_int(cell("miles")),
                color=str(cell("color") or "").strip(),
                location=str(cell("location") or "").strip(),
                vin=str(vin).strip(),
                ask_price=_as_int(cell("price")),
                section=section,
            )
        )
    return vehicles


# --- Cache --------------------------------------------------------------------
def load_cache(path: str) -> dict:
    p = Path(path)
    return json.loads(p.read_text()) if p.exists() else {}


def save_cache(path: str, cache: dict) -> None:
    Path(path).write_text(json.dumps(cache, indent=2))


# --- Main ---------------------------------------------------------------------
def main() -> None:
    ap = argparse.ArgumentParser(description="Manheim MMR deal finder")
    ap.add_argument("inventory", help="Path to inventory .xlsx")
    ap.add_argument("--min-margin", type=float, default=0.12,
                    help="Flag as a deal when (MMR-ask)/MMR >= this. Default 0.12 (12%%).")
    ap.add_argument("--exclude-risky", action="store_true",
                    help="Skip Damage Units and Trade Units (red-light, non-returnable).")
    ap.add_argument("--dry-run", action="store_true",
                    help="Parse inventory only; do not call the Manheim API.")
    ap.add_argument("--limit", type=int, default=0, help="Only process first N vehicles (testing).")
    ap.add_argument("--cache", default="mmr_cache.json")
    ap.add_argument("--out", default="deals_ranked.xlsx")
    ap.add_argument("--sleep", type=float, default=0.2, help="Seconds between API calls.")
    args = ap.parse_args()

    load_dotenv()

    vehicles = parse_inventory(args.inventory)
    print(f"Parsed {len(vehicles)} vehicles from {args.inventory}")

    if args.exclude_risky:
        before = len(vehicles)
        vehicles = [v for v in vehicles if v.section not in RISKY_SECTIONS]
        print(f"Excluded {before - len(vehicles)} risky (Damage/Trade) units.")

    if args.limit:
        vehicles = vehicles[: args.limit]

    if args.dry_run:
        by_section: dict[str, int] = {}
        for v in vehicles:
            by_section[v.section] = by_section.get(v.section, 0) + 1
        print("\nDry run — no API calls. Section counts:")
        for s, n in by_section.items():
            print(f"  {s:16} {n}")
        write_output(vehicles, args.out, args.min_margin)
        return

    client_id = os.environ.get("MANHEIM_CLIENT_ID")
    client_secret = os.environ.get("MANHEIM_CLIENT_SECRET")
    if not client_id or not client_secret:
        sys.exit(
            "Missing MANHEIM_CLIENT_ID / MANHEIM_CLIENT_SECRET.\n"
            "Copy .env.example to .env and fill in your Manheim API credentials."
        )

    client = ManheimClient(client_id, client_secret)
    cache = load_cache(args.cache)
    errors = 0

    for i, v in enumerate(vehicles, 1):
        key = f"{v.vin}:{v.miles or 0}"
        try:
            if key in cache:
                payload = cache[key]
            else:
                payload = client.mmr_by_vin(v.vin, v.miles)
                cache[key] = payload
                time.sleep(args.sleep)
            v.mmr = extract_mmr_value(payload)
            if v.mmr and v.ask_price:
                v.spread = v.mmr - v.ask_price
                v.margin = round(v.spread / v.mmr, 4)
            else:
                v.note = "no MMR or no price"
        except requests.HTTPError as e:
            errors += 1
            v.note = f"API error {e.response.status_code}"
        except Exception as e:  # noqa: BLE001
            errors += 1
            v.note = f"error: {e}"

        if i % 25 == 0:
            print(f"  ...{i}/{len(vehicles)} looked up")
            save_cache(args.cache, cache)

    save_cache(args.cache, cache)
    print(f"Done. {errors} lookup errors. Cache: {args.cache}")

    write_output(vehicles, args.out, args.min_margin)


def write_output(vehicles: list[Vehicle], out_path: str, min_margin: float) -> None:
    scored = [v for v in vehicles if v.margin is not None]
    scored.sort(key=lambda v: v.margin, reverse=True)
    deals = [v for v in scored if v.margin >= min_margin]

    print(f"\n=== {len(deals)} opportunities at >= {min_margin:.0%} under MMR ===")
    print(f"{'Margin':>7}  {'Spread':>8}  {'Ask':>8}  {'MMR':>8}  Year Make  Model      Miles   VIN")
    for v in deals[:40]:
        print(
            f"{v.margin:>7.0%}  {(v.spread or 0):>8,}  {(v.ask_price or 0):>8,}  "
            f"{(v.mmr or 0):>8,}  {v.year or '':>4} {v.make:<5} {v.model:<10} "
            f"{(v.miles or 0):>6,}  {v.vin}"
        )

    # Full ranked spreadsheet.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ranked"
    headers = ["Rank", "IsDeal", "Margin%", "Spread($)", "Ask($)", "MMR($)",
               "Unit", "Year", "Make", "Model", "Series", "Miles", "Color",
               "Location", "Section", "VIN", "Note"]
    ws.append(headers)
    for rank, v in enumerate(scored, 1):
        ws.append([
            rank,
            "YES" if (v.margin is not None and v.margin >= min_margin) else "",
            round(v.margin * 100, 1) if v.margin is not None else None,
            v.spread, v.ask_price, v.mmr, v.unit, v.year, v.make, v.model,
            v.series, v.miles, v.color, v.location, v.section, v.vin, v.note,
        ])
    # Append the ones we couldn't score at the bottom for visibility.
    for v in vehicles:
        if v.margin is None:
            ws.append([None, "", None, None, v.ask_price, v.mmr, v.unit, v.year,
                       v.make, v.model, v.series, v.miles, v.color, v.location,
                       v.section, v.vin, v.note or "not scored"])
    wb.save(out_path)
    print(f"\nFull ranked table written to {out_path}")


if __name__ == "__main__":
    main()
