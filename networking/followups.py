
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook
import os

WORKBOOK = Path(os.environ["NETWORKING_WORKBOOK"])

def main():
    target_date = date.fromisoformat(
        input("Enter follow-up date (YYYY-MM-DD): ").strip()
    )

    workbook = load_workbook(
        WORKBOOK,
        read_only=True,
        data_only=True
    )

    try:
        sheet = workbook["Contacts"]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows)

        columns = {
            name: headers.index(name)
            for name in ["Name", "Company", "Next Action", "Follow-up Due"]
        }

        count = 0

        for row in rows:
            due = row[columns["Follow-up Due"]]

            if isinstance(due, datetime):
                due = due.date()

            if isinstance(due, date) and due == target_date:
                count += 1
                print(f"\nName: {row[columns['Name']]}")
                print(f"Company: {row[columns['Company']]}")
                print(f"Next Action: {row[columns['Next Action']]}")

        print(f"\nTotal follow-ups: {count}")

    finally:
        workbook.close()


if __name__ == "__main__":
    main()
