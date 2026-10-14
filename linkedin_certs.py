"""
LinkedIn certifications -> Excel  (RHCSA RHCE EX280 EX374 AWS AZURE CKA CCNA vmware Python Anthropic)

For every row of the month sheet that has a Linkedin Id, one by one:
  - open the LinkedIn URL in Chrome with the "Arish - Chrome" Desktop shortcut's profile
  - look at the profile's "Licenses & certifications"
  - type "y" in each certification column the person has

Only empty cells are filled, a backup of the Excel file is made before saving,
and profiles already checked (with none of our certifications) are not opened
again unless you use --recheck.

Usage
  python linkedin_certs.py                         # asks the date (Enter = whole month)
  python linkedin_certs.py --month 2026-10         # every row of October 2026
  python linkedin_certs.py --date 07-10-2026       # only that day
  python linkedin_certs.py --limit 3               # only the first 3 profiles (for a test)
  python linkedin_certs.py --list                  # just show which rows would be checked
  python linkedin_certs.py --dry-run               # check everything but don't save the Excel
  python linkedin_certs.py --recheck               # also open profiles checked before
"""
import argparse
import datetime as dt
import json
import re
import time

from main import (HERE, CERTS, ask_or_parse_scope, check_excel_closed, clean_cell_text, is_blank, linkedin_profile, load_config, log, main_text, month_rows, open_chrome,
                  read_certifications, save_cells)

# sheet headers that mean the same certification (August has "Claude" instead of "Anthropic")
ALIASES = {"claude": "anthropic"}
CHECKED_FILE = HERE / "logs" / "certs_checked.json"


def cert_columns(cols):
    """{sheet column name: CERTS key} for the certification columns this sheet has."""
    found = {}
    for name in cols:
        key = ALIASES.get(name, name)
        if key in CERTS:
            found[name] = key
    return found


def load_checked():
    try:
        return json.loads(CHECKED_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_checked(checked):
    CHECKED_FILE.parent.mkdir(exist_ok=True)
    CHECKED_FILE.write_text(json.dumps(checked, indent=2), encoding="utf-8")


def profile_text(driver, url):
    """Open the profile like pasting the URL, then read its certifications list.
    '' = no certifications, None = could not be read."""
    driver.get(url)
    time.sleep(3)
    text = read_certifications(driver, url)
    if text is None:  # details page failed -> use the "Licenses & certifications" part of the profile
        driver.get(url)
        time.sleep(4)
        for _ in range(4):
            driver.execute_script("window.scrollBy(0, 900);")
            time.sleep(1)
        page = main_text(driver)
        m = re.search(r"Licenses\s*&\s*certifications(.*?)(?:\n(?:Skills|Projects|Courses|Honors|"
                      r"Languages|Recommendations|Interests|Volunteering|Publications)\n|$)", page, re.S | re.I)
        text = m.group(1) if m else ("" if page else None)
    return text


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="DD-MM-YYYY: only rows of that day")
    ap.add_argument("--month", help="YYYY-MM: every row of that month")
    ap.add_argument("--limit", type=int, help="only the first N profiles")
    ap.add_argument("--list", action="store_true", help="only show the rows that would be checked")
    ap.add_argument("--dry-run", action="store_true", help="check the profiles but don't save the Excel file")
    ap.add_argument("--recheck", action="store_true", help="also open profiles already checked before")
    args = ap.parse_args()

    cfg = load_config()
    excel = cfg.get("excel", "path")
    date, month_day = ask_or_parse_scope(args)
    log(f"=== Certifications from LinkedIn   Excel: {excel}")
    check_excel_closed(excel)

    ws, cols, rows = month_rows(excel, date, month_day, ["linkedin id"])
    certs = cert_columns(cols)
    if not certs:
        raise SystemExit(f"No certification columns (RHCSA, RHCE, ...) in sheet {ws.title}")
    log(f"Sheet {ws.title} certification columns: {', '.join(n.upper() for n in certs)}")

    checked = load_checked()
    todo, skipped = [], 0
    for r in rows:
        raw = str(ws.cell(r, cols["linkedin id"]).value or "").strip()
        url = linkedin_profile(raw)
        if not url or not all(is_blank(ws.cell(r, cols[c]).value) for c in certs):
            continue
        if url in checked and not args.recheck:
            skipped += 1
            continue
        todo.append({"row": r, "name": clean_cell_text(ws.cell(r, cols["name"]).value),
                     "linkedin": raw, "url": url, "certs": None})
    if args.limit:
        todo = todo[:args.limit]
    log(f"{len(todo)} LinkedIn profile(s) to check" + (f" on {date:%d-%b-%Y}" if date else "")
        + (f" ({skipped} checked before with none of ours - use --recheck to open them again)" if skipped else ""))
    for x in todo:
        log(f"  row {x['row']}: {x['name']} | {x['url']}")
    if not todo or args.list:
        return

    logs = HERE / "logs"
    logs.mkdir(exist_ok=True)
    progress = logs / f"certs_{dt.datetime.now():%Y%m%d_%H%M%S}.json"
    delay = cfg.getfloat("certs", "delay", fallback=5)
    try:
        driver = open_chrome(cfg, "linkedin", "https://www.linkedin.com/feed/", "linkedin.com")
        for i, x in enumerate(todo, 1):
            log(f"[{i}/{len(todo)}] {x['name']} | {x['url']}")
            try:
                text = profile_text(driver, x["url"])
            except Exception as e:
                log(f"    error: {type(e).__name__}: {str(e).splitlines()[0][:150]}")
                text = None
            if text is None:
                log("    could not open the profile - skipped")
            else:
                found = {key for key, pattern in CERTS.items() if re.search(pattern, text, re.I)}
                x["certs"] = [name for name, key in certs.items() if key in found]
                x["text"] = text[:4000]
                log(f"    certifications: {', '.join(c.upper() for c in x['certs']) or '(none of ours)'}")
                if not x["certs"]:
                    checked[x["url"]] = f"{dt.date.today()}"
            progress.write_text(json.dumps(todo, indent=2, ensure_ascii=False), encoding="utf-8")
            time.sleep(delay)
    except KeyboardInterrupt:
        log("Stopped - saving what was found so far")
    finally:
        updates = [{"row": x["row"], "key": ("linkedin id", x["linkedin"]),
                    "values": {c: "y" for c in x["certs"]},
                    "label": f"{x['name']} -> {', '.join(c.upper() for c in x['certs'])}"}
                   for x in todo if x["certs"]]
        if args.dry_run:
            log(f"Dry run - Excel file NOT saved ({len(updates)} profile(s) with certifications)")
            for u in updates:
                log(f"  row {u['row']}: {u['label']}")
        else:
            save_cells(excel, ws.title, updates, cfg.get("excel", "backup_dir"), "certifications")
            save_checked(checked)
    log("Done")


if __name__ == "__main__":
    main()
