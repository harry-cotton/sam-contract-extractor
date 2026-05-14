"""
SAM.gov Fetcher
================
Downloads a SAM.gov opportunity and all its attachments to a local folder
that extractor_multi.py can consume directly.

This closes the last manual step in the BD pipeline. Previously you would
visit SAM.gov, download each attachment by hand, and drop them into a
Sample Contracts folder. Now you just pass the notice ID or solicitation
number and the script fetches everything for you.

Accepts either:
  - A noticeid (the UUID-style ID shown in SAM.gov opportunity URLs,
    e.g. "73d5f8b8b7c14abea4c5b29b432c2dd1"), or
  - A solicitation number (the agency-issued number printed on the
    documents themselves, e.g. "70US0926R70093666")

The script autodetects which one you passed.

Usage:
    python sam_fetcher.py <notice_id_or_solicitation_number>
    python sam_fetcher.py <notice_id_or_solicitation_number> --output-dir <path>

After fetching, run the existing extractor:
    python extractor_multi.py "Sample Contracts/<solicitation_number>"

Required environment variable:
    SAM_GOV_API_KEY -- public API key from sam.gov (free, ~5 min setup).
                       Sign in to sam.gov, go to Account Details, generate
                       a Public API Key.

Required libraries:
    pip install requests

Author: Harry Cotton
"""

import os
import sys
import json
import re
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

import requests


SAM_API_BASE = "https://api.sam.gov/opportunities/v2/search"
DEFAULT_OUTPUT_ROOT = "Sample Contracts"
REQUEST_TIMEOUT = 60


# ============================================================
# STEP 1: API KEY AND INPUT HANDLING
# ============================================================

def load_api_key() -> str:
    """Read the SAM.gov public API key from the environment."""
    key = os.environ.get("SAM_GOV_API_KEY")
    if not key:
        print("Error: SAM_GOV_API_KEY environment variable is not set.\n")
        print("How to fix:")
        print("  1. Sign in to sam.gov")
        print("  2. Go to Account Details and generate a Public API Key")
        print("  3. Set the env var:")
        print("       Windows:    set SAM_GOV_API_KEY=your-key-here")
        print("       Mac/Linux:  export SAM_GOV_API_KEY=your-key-here")
        sys.exit(1)
    return key


def looks_like_notice_uuid(s: str) -> bool:
    """
    SAM.gov noticeids are 32-character hex strings (sometimes with dashes).
    Solicitation numbers contain letters and digits but rarely follow that
    pattern, so a simple length-plus-hex check is enough to disambiguate.
    """
    stripped = s.replace("-", "").strip()
    if len(stripped) != 32:
        return False
    return all(c in "0123456789abcdefABCDEF" for c in stripped)


# ============================================================
# STEP 2: FETCH OPPORTUNITY METADATA FROM SAM.gov
# ============================================================

def default_date_window() -> tuple:
    """
    SAM.gov caps the postedFrom/postedTo window at 1 year. We default to the
    last 12 months ending today, which catches any currently-active opportunity.
    Users can override with --posted-from / --posted-to for older opportunities.
    """
    today = datetime.now()
    one_year_ago = today - timedelta(days=364)  # 364 to stay safely under the 1-year cap
    return one_year_ago.strftime("%m/%d/%Y"), today.strftime("%m/%d/%Y")


def fetch_opportunity(identifier: str, api_key: str, posted_from: str, posted_to: str) -> dict:
    """
    Look up an opportunity by either noticeid (UUID) or solicitation number.
    SAM.gov requires postedFrom/postedTo on every query, and caps the window
    at 1 year apart.
    """
    if looks_like_notice_uuid(identifier):
        id_param = {"noticeid": identifier}
        id_label = "noticeid"
    else:
        id_param = {"solnum": identifier}
        id_label = "solnum"

    params = {
        "api_key": api_key,
        "limit": 10,
        "postedFrom": posted_from,
        "postedTo": posted_to,
        **id_param,
    }

    print(f"Querying SAM.gov by {id_label}={identifier}")
    print(f"  postedFrom={posted_from}  postedTo={posted_to}")
    response = requests.get(SAM_API_BASE, params=params, timeout=REQUEST_TIMEOUT)

    if response.status_code != 200:
        print(f"Error: SAM.gov returned HTTP {response.status_code}")
        print(response.text[:500])
        sys.exit(1)

    data = response.json()
    opps = data.get("opportunitiesData", [])

    if not opps:
        print(f"Error: No opportunity found for {id_label}={identifier} in {posted_from} - {posted_to}.")
        print(f"Total records reported by SAM.gov: {data.get('totalRecords', 0)}")
        print()
        print("If the opportunity was posted more than 12 months ago, try a custom date window:")
        print(f'  python sam_fetcher.py {identifier} --posted-from MM/DD/YYYY --posted-to MM/DD/YYYY')
        print("Note: SAM.gov caps the window at 1 year apart.")
        sys.exit(1)

    if len(opps) > 1:
        print(f"Note: {len(opps)} results returned. Using the most recent.")
        # Sort by postedDate descending if available
        opps.sort(key=lambda o: o.get("postedDate", ""), reverse=True)

    return opps[0]


# ============================================================
# STEP 3: ATTACHMENT DOWNLOAD
# ============================================================

def sanitize_filename(name: str) -> str:
    """Make a string safe to use as a filename on any OS."""
    name = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "attachment"


def extract_filename_from_response(response: requests.Response, fallback: str) -> str:
    """Pull the original filename out of Content-Disposition if SAM.gov set it."""
    cd = response.headers.get("Content-Disposition", "")
    match = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd)
    if match:
        return sanitize_filename(match.group(1))

    # Fall back to the URL path
    url_name = os.path.basename(urlparse(response.url).path)
    if url_name:
        return sanitize_filename(url_name)

    return sanitize_filename(fallback)


def download_attachment(url: str, api_key: str, output_folder: str, fallback_name: str) -> dict:
    """
    Download one attachment. SAM.gov sometimes requires the api_key on the
    resource link itself; we try with the key, and silently retry without it
    if needed.

    Returns a structured status dict (always - never None) so the caller can
    log every attempt in download_log.json.
    """
    result = {
        "url": url,
        "filename": None,
        "size_bytes": 0,
        "status": "failed",
        "http_status": None,
        "error": None,
    }

    def try_get(target_url: str):
        try:
            return requests.get(target_url, timeout=REQUEST_TIMEOUT, stream=True, allow_redirects=True)
        except requests.RequestException as e:
            result["error"] = f"network: {e}"
            print(f"  ! Network error: {e}")
            return None

    # Append api_key if not already present
    sep = "&" if "?" in url else "?"
    url_with_key = url if "api_key=" in url else f"{url}{sep}api_key={api_key}"

    response = try_get(url_with_key)
    if response is None:
        return result
    if response.status_code != 200:
        result["http_status"] = response.status_code
        # Retry without the key — some public URLs reject the extra param
        response = try_get(url)
        if response is None or response.status_code != 200:
            status = response.status_code if response else "no-response"
            result["http_status"] = status if isinstance(status, int) else result["http_status"]
            result["error"] = f"http: {status}"
            print(f"  ! Skipped (HTTP {status}): {url[:90]}")
            return result

    filename = extract_filename_from_response(response, fallback_name)
    output_path = os.path.join(output_folder, filename)

    # Avoid clobbering same-named files
    base, ext = os.path.splitext(output_path)
    counter = 1
    while os.path.exists(output_path):
        output_path = f"{base} ({counter}){ext}"
        counter += 1

    with open(output_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)

    size_bytes = os.path.getsize(output_path)
    result.update({
        "filename": os.path.basename(output_path),
        "size_bytes": size_bytes,
        "status": "downloaded",
        "http_status": 200,
    })
    print(f"  + {os.path.basename(output_path)} ({size_bytes / 1024:.1f} KB)")
    return result


def get_resource_links(opportunity: dict) -> list:
    """
    SAM.gov returns resourceLinks as either a list of URL strings or a list
    of dicts depending on which version of the response you get. Normalize
    to a flat list of URLs.
    """
    raw = opportunity.get("resourceLinks") or []
    urls = []
    for item in raw:
        if isinstance(item, str):
            urls.append(item)
        elif isinstance(item, dict):
            url = item.get("url") or item.get("link") or item.get("href")
            if url:
                urls.append(url)
    return urls


# ============================================================
# STEP 4: MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Fetch a SAM.gov opportunity and all its attachments."
    )
    parser.add_argument(
        "identifier",
        help="SAM.gov noticeid (UUID) or solicitation number",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help=f"Output folder. Defaults to '{DEFAULT_OUTPUT_ROOT}/<solicitation_number>/'.",
    )
    parser.add_argument(
        "--posted-from",
        default=None,
        help="Earliest posted date in MM/DD/YYYY (defaults to 364 days ago). SAM.gov caps the window at 1 year.",
    )
    parser.add_argument(
        "--posted-to",
        default=None,
        help="Latest posted date in MM/DD/YYYY (defaults to today). SAM.gov caps the window at 1 year.",
    )
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Force re-download even if the output folder has a prior download_log.json marking a complete run.",
    )
    args = parser.parse_args()

    api_key = load_api_key()

    # Resolve the date window
    default_from, default_to = default_date_window()
    posted_from = args.posted_from or default_from
    posted_to = args.posted_to or default_to

    opp = fetch_opportunity(args.identifier, api_key, posted_from, posted_to)

    sol_num = opp.get("solicitationNumber") or args.identifier
    title = opp.get("title", "Unknown opportunity")
    agency = opp.get("fullParentPathName") or opp.get("departmentName") or "Unknown agency"

    print()
    print("=" * 60)
    print(f"  TITLE:     {title}")
    print(f"  SOL #:     {sol_num}")
    print(f"  AGENCY:    {agency}")
    print(f"  POSTED:    {opp.get('postedDate', 'N/A')}")
    print(f"  DEADLINE:  {opp.get('responseDeadLine', 'N/A')}")
    print(f"  TYPE:      {opp.get('type', 'N/A')}")
    print("=" * 60)
    print()

    folder_name = sanitize_filename(sol_num)
    output_folder = args.output_dir or os.path.join(DEFAULT_OUTPUT_ROOT, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    print(f"Output folder: {output_folder}")

    # Save the full opportunity metadata for traceability
    meta_path = os.path.join(output_folder, "sam_metadata.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(opp, f, indent=2)
    print("  + sam_metadata.json")
    print()

    links = get_resource_links(opp)
    if not links:
        print("No resource links found on this opportunity.")
        print("This sometimes happens when the notice embeds content in its")
        print("description rather than attaching files. Inspect sam_metadata.json")
        print("or open the opportunity on sam.gov to see what's available.")
        print()
        print(f"  Description link: {opp.get('description', 'N/A')}")
        return

    # Cache check: if a prior complete run wrote a download_log.json and the
    # caller didn't pass --no-cache, skip the attachment downloads entirely.
    # Metadata was just refreshed above either way, so this still picks up any
    # amendments or Q&A updates - only the heavy attachment downloads are skipped.
    log_path = os.path.join(output_folder, "download_log.json")
    cached_log = None
    if os.path.exists(log_path) and not args.no_cache:
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                cached_log = json.load(f)
        except (json.JSONDecodeError, OSError):
            cached_log = None

    if cached_log:
        downloaded_count = cached_log.get("summary", {}).get("downloaded", 0)
        total_count = cached_log.get("summary", {}).get("total", 0)
        print(f"Cache hit: previous run downloaded {downloaded_count}/{total_count} attachments.")
        print(f"  Skipping downloads. Pass --no-cache to force re-download.")
        print(f"  Cached log: {os.path.basename(log_path)} (from {cached_log.get('fetched_at', 'unknown date')})")
        print()
        print("Next step in the pipeline:")
        print(f'  python extractor_multi.py "{output_folder}"')
        return

    print(f"Downloading {len(links)} attachment(s):")
    attachment_results = []
    downloaded = 0
    for i, url in enumerate(links, start=1):
        result = download_attachment(
            url=url,
            api_key=api_key,
            output_folder=output_folder,
            fallback_name=f"attachment_{i}",
        )
        attachment_results.append(result)
        if result.get("status") == "downloaded":
            downloaded += 1

    # Write download_log.json with per-attachment status. This is the cache
    # marker for future runs and the audit trail for any failed downloads.
    log = {
        "identifier": args.identifier,
        "solicitation_number": sol_num,
        "title": title,
        "agency": agency,
        "fetched_at": datetime.now().isoformat(timespec="seconds"),
        "attachments": attachment_results,
        "summary": {
            "total": len(links),
            "downloaded": downloaded,
            "failed": len(links) - downloaded,
        },
    }
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log, f, indent=2)

    print()
    print(f"Done. {downloaded}/{len(links)} attachment(s) saved to: {output_folder}")
    print(f"  Download log: {os.path.basename(log_path)}")
    print()
    print("Next step in the pipeline:")
    print(f'  python extractor_multi.py "{output_folder}"')


if __name__ == "__main__":
    main()
