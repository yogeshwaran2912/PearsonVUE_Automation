"""Offline check of the Excel step with sample data. Never saves the Excel file."""
import csv
import datetime as dt
import tempfile
from pathlib import Path

import pvue_automation as pv

cfg = pv.load_config()
date = dt.date(2026, 10, 6)
people = [
    {"name": "KOSAL RAJ I", "client": "Google Cloud", "code": "PR000260", "phone": "+91 9791117206"},  # already there
    {"name": "Test Person", "client": "Microsoft", "code": "AZ-900", "phone": "+91 98765 43210"},     # new
]
tmp = Path(tempfile.gettempdir()) / "fb_test.csv"
with open(tmp, "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(["Timestamp", "Name", "Phone Number", "Email address", "Your company name",
                "Your Designation", "Let others know about your experience"])
    w.writerow(["10/6/2026 18:05:11", "Kosal Raj", "9791117206", "kosal@example.com", "Acme", "Engineer", "Great centre"])
    w.writerow(["10/5/2026 11:00:00", "Old", "9000000000", "old@example.com", "X", "Y", "yesterday"])
feedback = pv.read_feedback(tmp, date, "")
pv.update_excel(cfg.get("excel", "path"), date, people, feedback, "backups", dry_run=True)
