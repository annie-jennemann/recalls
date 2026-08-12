import csv
import tempfile
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


FDA_URL = "https://www.fda.gov/safety/recalls-market-withdrawals-safety-alerts"
OUTPUT_PATH = Path("data/fda_recalls.csv")


def download_workbook(destination: Path) -> None:
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            page = browser.new_page(accept_downloads=True)
            page.goto(FDA_URL, wait_until="domcontentloaded", timeout=60_000)
            link = page.locator(
                'a[href*="datatables-data"][href*="_format=xlsx"]'
            ).first
            link.wait_for(state="visible", timeout=30_000)

            with page.expect_download(timeout=60_000) as download_info:
                link.click()
            download_info.value.save_as(str(destination))
        except PlaywrightTimeoutError as exc:
            raise RuntimeError("The FDA XLSX download link did not load in time.") from exc
        finally:
            browser.close()


def cell_value(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return value


def workbook_to_csv(workbook_path: Path, csv_path: Path) -> None:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with csv_path.open("w", newline="", encoding="utf-8") as output:
            writer = csv.writer(output)
            for row in worksheet.iter_rows(values_only=True):
                writer.writerow([cell_value(value) for value in row])
    finally:
        workbook.close()


def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as temp_dir:
        workbook_path = Path(temp_dir) / "fda-recalls.xlsx"
        csv_path = Path(temp_dir) / "fda_recalls.csv"
        download_workbook(workbook_path)
        workbook_to_csv(workbook_path, csv_path)

        new_contents = csv_path.read_bytes()
        old_contents = OUTPUT_PATH.read_bytes() if OUTPUT_PATH.exists() else None
        OUTPUT_PATH.write_bytes(new_contents)

    print("CSV updated." if new_contents != old_contents else "CSV unchanged.")


if __name__ == "__main__":
    main()
