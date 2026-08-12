import csv
import os
import tempfile
from datetime import date, datetime
from pathlib import Path

import requests
from openpyxl import load_workbook


FDA_URL = (
    "https://www.fda.gov/"
    "safety/recalls-market-withdrawals-safety-alerts"
)

OUTPUT_PATH = Path("data/fda_recalls.csv")


def download_workbook(destination: Path) -> None:
    token = os.environ["BROWSERLESS_TOKEN"]

    browserless_url = (
        "https://production-sfo.browserless.io/"
        f"download?token={token}"
    )

    browser_code = f"""
export default async ({{ page }}) => {{
  await page.goto({FDA_URL!r}, {{
    waitUntil: "networkidle2",
    timeout: 60000
  }});

  const link = await page.$(
    'a[href*="datatables-data"][href*="_format=xlsx"]'
  );

  if (!link) {{
    throw new Error("FDA XLSX download link was not found.");
  }}

  await link.click();

  // Allow the browser download to finish.
  await new Promise(resolve => setTimeout(resolve, 8000));
}};
"""

    response = requests.post(
        browserless_url,
        headers={"Content-Type": "application/javascript"},
        data=browser_code,
        timeout=120,
    )

    if response.status_code != 200:
        raise RuntimeError(
            "Browserless download failed: "
            f"{response.status_code}\n"
            f"{response.text[:1000]}"
        )

    content_type = response.headers.get("content-type", "")

    if (
        "spreadsheet" not in content_type
        and "excel" not in content_type
        and not response.content.startswith(b"PK")
    ):
        raise RuntimeError(
            "Browserless returned an unexpected file type: "
            f"{content_type}\n"
            f"{response.text[:500]}"
        )

    destination.write_bytes(response.content)
    print(f"Downloaded {len(response.content)} bytes.")


def normalize_cell(value):
    if value is None:
        return ""

    if isinstance(value, (datetime, date)):
        return value.isoformat()

    return value


def workbook_to_csv(
    workbook_path: Path,
    csv_path: Path,
) -> None:
    workbook = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
    )

    try:
        worksheet = workbook.active
        csv_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        with csv_path.open(
            "w",
            newline="",
            encoding="utf-8",
        ) as output:
            writer = csv.writer(output)

            for row in worksheet.iter_rows(
                values_only=True
            ):
                writer.writerow(
                    [
                        normalize_cell(value)
                        for value in row
                    ]
                )

    finally:
        workbook.close()


def main() -> None:
    OUTPUT_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with tempfile.TemporaryDirectory() as temp_dir:
        workbook_path = (
            Path(temp_dir) / "fda-recalls.xlsx"
        )
        csv_path = (
            Path(temp_dir) / "fda_recalls.csv"
        )

        download_workbook(workbook_path)
        workbook_to_csv(workbook_path, csv_path)

        new_contents = csv_path.read_bytes()

        old_contents = (
            OUTPUT_PATH.read_bytes()
            if OUTPUT_PATH.exists()
            else None
        )

        OUTPUT_PATH.write_bytes(new_contents)

    if new_contents == old_contents:
        print("CSV unchanged.")
    else:
        print(f"CSV updated: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
