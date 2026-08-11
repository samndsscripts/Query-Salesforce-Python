"""
find_view_id_from_url.py

You have a Tableau view open in your browser (e.g. the new product/SKU
view you just built) but need its actual view ID (LUID) to put in .env.
The browser URL doesn't show that GUID directly -- it shows a workbook
and view name slug instead. This resolves that slug back to the real ID
via the API.

Usage:
    python find_view_id_from_url.py "<paste the full browser URL here>"

Works with URLs like:
    https://tableau.adspipe.com/#/site/mysite/views/WorkbookName/SheetName
    https://tableau.adspipe.com/#/views/WorkbookName/SheetName?:iid=1
    https://tableau.adspipe.com/t/mysite/views/WorkbookName/SheetName

Or just pass the "WorkbookName/SheetName" part directly:
    python find_view_id_from_url.py "WorkbookName/SheetName"

Run from the repo root so it picks up the same .env your pipeline uses.
"""

import os
import re
import sys

from dotenv import load_dotenv
import tableauserverclient as TSC

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

load_dotenv()

TABLEAU_SERVER = os.getenv("TABLEAU_SERVER")
TABLEAU_PAT_NAME = os.getenv("TABLEAU_PAT_NAME")
TABLEAU_PAT_SECRET = os.getenv("TABLEAU_PAT_SECRET")
TABLEAU_SITE = os.getenv("TABLEAU_SITE_CONTENTURL", "")


def die(msg: str) -> None:
    print(f"\n[ERROR] {msg}\n")
    sys.exit(1)


def normalize(s: str) -> str:
    """Lowercase, strip everything but letters/digits so 'My Workbook',
    'My_Workbook', and 'My-Workbook' all compare equal."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def extract_slugs(url_or_path: str):
    """Pull the workbook and view name slugs out of a Tableau URL, or just
    pass through if given 'Workbook/View' directly."""
    # Strip query string / fragment junk like ?:iid=1
    cleaned = url_or_path.split("?")[0].strip()

    m = re.search(r"/views/([^/]+)/([^/]+)", cleaned)
    if m:
        return m.group(1), m.group(2)

    # Fallback: assume they passed "Workbook/View" directly
    parts = [p for p in cleaned.split("/") if p]
    if len(parts) >= 2:
        return parts[-2], parts[-1]

    die(f"Couldn't find a '/views/<workbook>/<view>' pattern in: {url_or_path!r}\n"
        f"Paste the full browser URL, or just 'WorkbookName/ViewName'.")


def main() -> None:
    if len(sys.argv) < 2:
        die("Usage: python find_view_id_from_url.py \"<tableau view URL or WorkbookName/ViewName>\"")

    wb_slug, view_slug = extract_slugs(sys.argv[1])
    wb_norm, view_norm = normalize(wb_slug), normalize(view_slug)
    print(f"Looking for workbook matching {wb_slug!r}, view matching {view_slug!r}...\n")

    missing = [k for k in ("TABLEAU_SERVER", "TABLEAU_PAT_NAME", "TABLEAU_PAT_SECRET")
               if not os.getenv(k)]
    if missing:
        die(f"Missing required .env values: {', '.join(missing)}.")

    print(f"Signing in to {TABLEAU_SERVER} (site: '{TABLEAU_SITE or '<default>'}') ...")
    auth = TSC.PersonalAccessTokenAuth(TABLEAU_PAT_NAME, TABLEAU_PAT_SECRET, site_id=TABLEAU_SITE)
    server = TSC.Server(TABLEAU_SERVER, use_server_version=True)

    matches = []
    with server.auth.sign_in(auth):
        print("Signed in OK. Fetching workbooks and views...\n")
        workbooks = {w.id: w for w in TSC.Pager(server.workbooks)}
        all_views = list(TSC.Pager(server.views, usage=True))

        for v in all_views:
            wb = workbooks.get(v.workbook_id)
            wb_name = wb.name if wb else ""
            # Compare both the friendly names AND the raw content_url pieces --
            # content_url is usually "Workbook/sheets/View", so check that too.
            content_url = v.content_url or ""
            content_parts = content_url.split("/")
            content_wb = content_parts[0] if content_parts else ""
            content_view = content_parts[-1] if content_parts else ""

            name_match = normalize(v.name) == view_norm
            content_match = normalize(content_view) == view_norm
            wb_match = normalize(wb_name) == wb_norm or normalize(content_wb) == wb_norm

            if (name_match or content_match) and wb_match:
                matches.append((v, wb_name))

        # If nothing matched on exact workbook+view, fall back to view-name-only
        # matches so you at least see candidates to eyeball.
        if not matches:
            print("No exact workbook+view match. Showing view-name-only matches instead:\n")
            for v in all_views:
                if normalize(v.name) == view_norm:
                    wb = workbooks.get(v.workbook_id)
                    matches.append((v, wb.name if wb else v.workbook_id))

    if not matches:
        die(f"No view found matching {view_slug!r} at all. Double check the URL/slug.")

    print(f"Found {len(matches)} match(es):\n")
    for v, wb_name in matches:
        print(f"  TABLEAU_PRODUCT_VIEW_ID={v.id}")
        print(f"    view name:  {v.name!r}")
        print(f"    workbook:   {wb_name!r}")
        print(f"    project:    {v.project_id}")
        print(f"    content_url:{v.content_url}")
        print()

    if len(matches) == 1:
        print(f"Update .env:\n  TABLEAU_PRODUCT_VIEW_ID={matches[0][0].id}")
    else:
        print("Multiple matches -- pick the one whose workbook/project you recognize.")


if __name__ == "__main__":
    main()
