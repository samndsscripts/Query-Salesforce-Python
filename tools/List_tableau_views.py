"""
list_tableau_views.py

Diagnostic helper for the Query-Salesforce-Python pipeline.

Your pipeline (src/combined_pipeline.py) calls server.views.get_by_id() using
three hardcoded view IDs (LUIDs) stored in .env:
    TABLEAU_RMA_VIEW_ID
    TABLEAU_SHIPMENT_VIEW_ID
    TABLEAU_PRODUCT_VIEW_ID

You're seeing:
    403004: Forbidden
    'Samuel.Cooper' isn't authorized to query the view
    '241e5cc1-6d7c-484a-99ad-d5700c3fd281'

That ID is your current TABLEAU_PRODUCT_VIEW_ID. View LUIDs are regenerated
whenever content is migrated/republished to a new Tableau Server/Site, even
if the workbook and view names are unchanged. So this is almost certainly
NOT a permissions problem on the new server -- it's that the old ID no
longer points to a view you have access to (it may not exist at all, or it
may now point to something else entirely).

This script signs in with the SAME credentials your pipeline already uses
(from .env) and prints every view visible to that user/PAT on the new
server/site, along with its workbook, project, and (most importantly) its
new ID. Match the workbook/view names you recognize (the ones that used to
back TABLEAU_RMA_VIEW_ID / TABLEAU_SHIPMENT_VIEW_ID / TABLEAU_PRODUCT_VIEW_ID)
to their new IDs, then update .env.

Usage:
    python list_tableau_views.py
    python list_tableau_views.py "product"     # filter by name (case-insensitive substring)

Run this from the repo root (same place you'd run combined_pipeline.py) so
it picks up the same .env file.
"""

import os
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

CURRENT_IDS = {
    "TABLEAU_RMA_VIEW_ID": os.getenv("TABLEAU_RMA_VIEW_ID"),
    "TABLEAU_SHIPMENT_VIEW_ID": os.getenv("TABLEAU_SHIPMENT_VIEW_ID"),
    "TABLEAU_PRODUCT_VIEW_ID": os.getenv("TABLEAU_PRODUCT_VIEW_ID"),
}


def die(msg: str) -> None:
    print(f"\n[ERROR] {msg}\n")
    sys.exit(1)


def main() -> None:
    missing = [k for k in ("TABLEAU_SERVER", "TABLEAU_PAT_NAME", "TABLEAU_PAT_SECRET")
               if not os.getenv(k)]
    if missing:
        die(f"Missing required .env values: {', '.join(missing)}. "
            f"Fill these in before running (same values combined_pipeline.py uses).")

    name_filter = sys.argv[1].lower() if len(sys.argv) > 1 else None

    print(f"Signing in to {TABLEAU_SERVER}  (site: '{TABLEAU_SITE or '<default>'}') ...")
    auth = TSC.PersonalAccessTokenAuth(TABLEAU_PAT_NAME, TABLEAU_PAT_SECRET, site_id=TABLEAU_SITE)
    server = TSC.Server(TABLEAU_SERVER, use_server_version=True)

    try:
        with server.auth.sign_in(auth):
            print("Signed in OK.\n")

            all_views = list(TSC.Pager(server.views, usage=True))
            if not all_views:
                die("Signed in fine, but zero views were returned. Either this "
                    "user/PAT has no view access on this site, or the site content "
                    "URL is wrong.")

            if name_filter:
                filtered = [v for v in all_views if name_filter in (v.name or "").lower()
                            or name_filter in (v.workbook_id or "")]
                print(f"Showing {len(filtered)} of {len(all_views)} views matching '{name_filter}':\n")
                rows = filtered
            else:
                print(f"Found {len(all_views)} views visible to this user/PAT:\n")
                rows = all_views

            # Column widths
            print(f"{'VIEW ID':<38} {'PROJECT':<28} {'WORKBOOK ID':<38} VIEW NAME")
            print("-" * 130)
            for v in sorted(rows, key=lambda x: (x.project_id or "", x.name or "")):
                print(f"{v.id:<38} {(v.project_id or ''):<28} {v.workbook_id:<38} {v.name}")

            print("\nCurrent .env view IDs and whether they matched something above:")
            found_ids = {v.id for v in all_views}
            for env_key, val in CURRENT_IDS.items():
                status = "FOUND (still valid/visible)" if val in found_ids else "NOT FOUND / NOT VISIBLE -- needs updating"
                print(f"  {env_key} = {val!r}  ->  {status}")

            print("\nNext step: find the row(s) above whose PROJECT/VIEW NAME match what your")
            print("old RMA / Shipment / Product views were called, copy the VIEW ID, and paste")
            print("it into the matching line in .env.")

    except TSC.ServerResponseError as e:
        die(f"Tableau server rejected the request: {e}\n\n"
            f"If this fails at sign-in (not listing views), the PAT itself may not have "
            f"carried over to the new server -- you'll need to generate a new Personal "
            f"Access Token on the new server and update TABLEAU_PAT_NAME / "
            f"TABLEAU_PAT_SECRET in .env.")


if __name__ == "__main__":
    main()
