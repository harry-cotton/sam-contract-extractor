"""
SAM.gov Search
==============
Discovery half of the BD pipeline. Where sam_fetcher.py answers "give me this
specific opportunity," sam_search.py answers "what live opportunities match
what I care about?"

Searches the SAM.gov public API by NAICS code, keyword, and posting date window,
returns a ranked list of active opportunities with notice IDs that can be piped
directly into sam_fetcher.py.

Usage:
    python sam_search.py --naics 541512
    python sam_search.py --naics 541512 --keyword analytics
    python sam_search.py --keyword "data lakehouse"
    python sam_search.py --naics 541512 --posted-from 03/01/2026 --limit 30
    python sam_search.py --naics 541512 --include-closed --include-all-types

Defaults:
    --naics:              none (no NAICS filter)
    --keyword:            none (no keyword filter)
    --posted-from:        60 days ago
    --posted-to:          today
    --limit:              20
    Active-only:          True  (response deadline in the future)
    Substantive types:    True  (excludes Sources Sought, Special Notice, Award)

Required env var:
    SAM_GOV_API_KEY -- same public API key sam_fetcher.py uses

Required libraries:
    pip install requests

After picking a candidate, download its full package with:
    python sam_fetcher.py <notice-id>

Author: Harry Cotton
"""

import os
import sys
import json
import argparse
from datetime import datetime, timedelta

import requests


SAM_API_BASE = "https://api.sam.gov/opportunities/v2/search"
REQUEST_TIMEOUT = 60


# ============================================================
# Setup
# ============================================================

def load_api_key() -> str:
    key = os.environ.get("SAM_GOV_API_KEY")
    if not key:
        print("Error: SAM_GOV_API_KEY environment variable is not set.")
        print("Get a free public API key from sam.gov -> Account Details.")
        sys.exit(1)
    return key


def default_window() -> tuple:
    """Default to the last 60 days. SAM caps the window at 1 year."""
    today = datetime.now()
    sixty_days_ago = today - timedelta(days=60)
    return sixty_days_ago.strftime("%m/%d/%Y"), today.strftime("%m/%d/%Y")


def parse_sam_date(s):
    """SAM date strings are ISO 8601 with timezone offsets."""
    if not s:
        return None
    try:
        return datetime.fromisoformat(str(s).replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return None


# ============================================================
# Filtering helpers
# ============================================================

def is_active(opp: dict) -> bool:
    """An opportunity is active if its response deadline is in the future."""
    deadline = parse_sam_date(opp.get("responseDeadLine"))
    if not deadline:
        return False
    # Strip timezone to compare against naive now()
    if deadline.tzinfo is not None:
        deadline = deadline.replace(tzinfo=None)
    return deadline > datetime.now()


def is_substantive(opp: dict) -> bool:
    """
    Substantive = something a BD team would actually respond to.
    Excludes Sources Sought, Special Notices, Award notices, Surplus, Intent-to-Bundle.
    SAM returns the type as a human-readable label like 'Solicitation' or
    'Combined Synopsis/Solicitation'.
    """
    type_text = (opp.get("type") or "").lower()
    if not type_text:
        return False
    # Exclude first
    excluded = ("sources sought", "special notice", "award", "surplus",
                "intent to bundle", "fair opportunity")
    if any(e in type_text for e in excluded):
        return False
    # Include if it's a real RFP-class notice
    included = ("solicitation", "combined synopsis", "pre-solicitation", "presolicitation",
                "request for proposal", "request for quote")
    return any(i in type_text for i in included)


# ============================================================
# Formatting helpers
# ============================================================

def fmt_deadline(s) -> str:
    d = parse_sam_date(s)
    return d.strftime("%Y-%m-%d") if d else "Not specified"


def fmt_agency(opp: dict) -> str:
    full = opp.get("fullParentPathName") or opp.get("departmentName") or "Unknown agency"
    return full[:70]


def truncate(s: str, n: int) -> str:
    s = (s or "").strip()
    return s if len(s) <= n else s[: n - 1] + "..."


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Search SAM.gov for live federal contract opportunities."
    )
    parser.add_argument("--naics", default=None, help="NAICS code (e.g., 541512 for Computer Systems Design Services)")
    parser.add_argument("--keyword", default=None, help="Keyword in opportunity title (e.g., 'analytics')")
    parser.add_argument("--posted-from", default=None, help="Earliest posted date MM/DD/YYYY (default: 60 days ago)")
    parser.add_argument("--posted-to", default=None, help="Latest posted date MM/DD/YYYY (default: today)")
    parser.add_argument("--limit", type=int, default=20, help="Max results to return after filtering (default: 20)")
    parser.add_argument("--include-closed", action="store_true",
                        help="Include opportunities past their response deadline")
    parser.add_argument("--include-all-types", action="store_true",
                        help="Include Sources Sought, Special Notices, Award notices, etc.")
    parser.add_argument("--output", default=None,
                        help="Save the full JSON result to this path (useful for downstream tooling)")
    args = parser.parse_args()

    api_key = load_api_key()
    default_from, default_to = default_window()
    posted_from = args.posted_from or default_from
    posted_to = args.posted_to or default_to

    # Over-fetch to allow client-side filtering for active + substantive
    raw_limit = min(args.limit * 5, 200)
    params = {
        "api_key": api_key,
        "limit": raw_limit,
        "postedFrom": posted_from,
        "postedTo": posted_to,
    }
    if args.naics:
        params["ncode"] = args.naics
    if args.keyword:
        params["title"] = args.keyword

    print(f"Searching SAM.gov...")
    print(f"  NAICS:                  {args.naics or 'any'}")
    print(f"  Keyword (title):        {args.keyword or 'any'}")
    print(f"  Posted window:          {posted_from} - {posted_to}")
    print(f"  Active-only:            {not args.include_closed}")
    print(f"  Substantive types only: {not args.include_all_types}")
    print()

    response = requests.get(SAM_API_BASE, params=params, timeout=REQUEST_TIMEOUT)
    if response.status_code != 200:
        print(f"Error: SAM.gov returned HTTP {response.status_code}")
        print(response.text[:500])
        sys.exit(1)

    data = response.json()
    opps = data.get("opportunitiesData", []) or []
    total_records = data.get("totalRecords", 0)

    # Client-side filtering
    pre_filter_count = len(opps)
    if not args.include_closed:
        opps = [o for o in opps if is_active(o)]
    if not args.include_all_types:
        opps = [o for o in opps if is_substantive(o)]
    post_filter_count = len(opps)

    # Sort by response deadline soonest-first; opps without a deadline go to the end
    opps.sort(key=lambda o: parse_sam_date(o.get("responseDeadLine")) or datetime(2099, 1, 1))

    # Trim to user's requested limit
    opps = opps[: args.limit]

    if not opps:
        print(f"No matching opportunities found.")
        print(f"  SAM.gov returned {total_records} total records before filtering.")
        print(f"  After active + substantive filters: 0 candidates.")
        print()
        print(f"To widen the search, try:")
        print(f"  --include-closed       (include opportunities past their deadline)")
        print(f"  --include-all-types    (include Sources Sought, Special Notices, etc.)")
        print(f"  --posted-from <date>   (widen the posting window)")
        print(f"  drop the --naics or --keyword filter")
        sys.exit(0)

    # Print human-readable list
    print(f"Found {post_filter_count} matching opportunities ({pre_filter_count} pre-filter from {total_records} total). Showing top {len(opps)}:\n")
    for i, o in enumerate(opps, start=1):
        notice_id = o.get("noticeId") or ""
        title = truncate(o.get("title") or "Untitled", 90)
        agency = fmt_agency(o)
        ptype = o.get("type") or ""
        deadline = fmt_deadline(o.get("responseDeadLine"))
        sol_num = o.get("solicitationNumber") or "-"
        print(f"  {i:>2}. {title}")
        print(f"      Notice ID:  {notice_id}")
        print(f"      Agency:     {agency}")
        print(f"      Type:       {ptype}")
        print(f"      Sol #:      {sol_num}")
        print(f"      Deadline:   {deadline}")
        print()

    print(f"To download an opportunity's full package, run:")
    print(f"  python sam_fetcher.py <notice-id>")

    # Optional JSON output for downstream tooling / agent
    if args.output:
        out = {
            "search_params": {
                "naics": args.naics,
                "keyword": args.keyword,
                "posted_from": posted_from,
                "posted_to": posted_to,
                "active_only": not args.include_closed,
                "substantive_types_only": not args.include_all_types,
            },
            "total_records_pre_filter": total_records,
            "result_count": len(opps),
            "results": opps,
        }
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2, default=str)
        print(f"\nFull JSON results saved to: {args.output}")


if __name__ == "__main__":
    main()
