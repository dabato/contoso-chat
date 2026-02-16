#!/usr/bin/env python3
"""Export Crestron Pro orders table data to an Excel file using Playwright.

Usage:
  python src/api/tools/crestron_orders_export.py \
    --email "$CRESTRON_EMAIL" \
    --password "$CRESTRON_PASSWORD" \
    --output crestron_orders.xlsx
"""

from __future__ import annotations

import argparse
import os
import time
from dataclasses import dataclass
from typing import List, Optional

import pandas as pd
from playwright.sync_api import Locator, Page, TimeoutError, sync_playwright

DEFAULT_URL = "https://www.crestron.com/pro/app/orders"


@dataclass
class ScrapeResult:
    headers: List[str]
    rows: List[List[str]]


def _first_visible(page: Page, selectors: List[str]) -> Optional[Locator]:
    for selector in selectors:
        candidate = page.locator(selector).first
        if candidate.count() and candidate.is_visible():
            return candidate
    return None


def _fill_first(page: Page, selectors: List[str], value: str) -> bool:
    for selector in selectors:
        locator = page.locator(selector).first
        if locator.count() and locator.is_visible():
            locator.click()
            locator.fill(value)
            return True
    return False


def login(page: Page, email: str, password: str, url: str) -> None:
    page.goto(url, wait_until="domcontentloaded", timeout=120000)

    email_selectors = [
        "input[type='email']",
        "input[name='email']",
        "input[name*='email' i]",
        "#email",
        "#username",
        "input[name='username']",
        "input[type='text']",
    ]
    password_selectors = [
        "input[type='password']",
        "input[name='password']",
        "#password",
        "input[name*='pass' i]",
    ]

    if not _fill_first(page, email_selectors, email):
        raise RuntimeError("Could not find the email/username field.")

    if not _fill_first(page, password_selectors, password):
        raise RuntimeError("Could not find the password field.")

    submit = _first_visible(
        page,
        [
            "button[type='submit']",
            "input[type='submit']",
            "button:has-text('Sign in')",
            "button:has-text('Log in')",
            "button:has-text('Login')",
        ],
    )
    if submit is None:
        raise RuntimeError("Could not find login submit button.")

    submit.click()
    page.wait_for_load_state("networkidle", timeout=120000)


def extract_table(page: Page) -> ScrapeResult:
    page.wait_for_selector("table", timeout=120000)

    headers = [h.inner_text().strip() for h in page.locator("table thead th").all()]
    if not headers:
        first_row_cells = page.locator("table tbody tr").first.locator("td")
        headers = [f"column_{i + 1}" for i in range(first_row_cells.count())]

    all_rows: List[List[str]] = []
    seen = set()

    while True:
        row_locators = page.locator("table tbody tr")
        row_count = row_locators.count()

        for index in range(row_count):
            cells = row_locators.nth(index).locator("td")
            row = [cells.nth(i).inner_text().strip() for i in range(cells.count())]
            key = tuple(row)
            if row and key not in seen:
                seen.add(key)
                all_rows.append(row)

        next_button = _first_visible(
            page,
            [
                "button[aria-label*='Next' i]",
                "a[aria-label*='Next' i]",
                "button:has-text('Next')",
                "a:has-text('Next')",
                "li.next:not(.disabled) a",
            ],
        )

        if next_button is None:
            break

        disabled_attr = next_button.get_attribute("disabled")
        aria_disabled = next_button.get_attribute("aria-disabled")
        class_name = next_button.get_attribute("class") or ""

        if disabled_attr is not None or aria_disabled == "true" or "disabled" in class_name.lower():
            break

        before = len(all_rows)
        next_button.click()
        page.wait_for_load_state("networkidle", timeout=120000)
        time.sleep(1)

        if len(all_rows) == before:
            pass

    return ScrapeResult(headers=headers, rows=all_rows)


def to_excel(result: ScrapeResult, output_path: str) -> None:
    width = max((len(r) for r in result.rows), default=len(result.headers))
    headers = result.headers

    if len(headers) < width:
        for i in range(len(headers), width):
            headers.append(f"column_{i + 1}")

    normalized_rows = [r + [""] * (width - len(r)) for r in result.rows]
    frame = pd.DataFrame(normalized_rows, columns=headers[:width])
    frame.to_excel(output_path, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape Crestron orders table and export to Excel.")
    parser.add_argument("--url", default=DEFAULT_URL, help=f"Orders URL (default: {DEFAULT_URL})")
    parser.add_argument("--email", default=os.getenv("CRESTRON_EMAIL"), help="Crestron account email")
    parser.add_argument("--password", default=os.getenv("CRESTRON_PASSWORD"), help="Crestron account password")
    parser.add_argument("--output", default="crestron_orders.xlsx", help="Output Excel path")
    parser.add_argument(
        "--headful",
        action="store_true",
        help="Run with a visible browser window (helpful if MFA/challenge is required).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if not args.email or not args.password:
        raise ValueError("Email and password are required. Use flags or CRESTRON_EMAIL/CRESTRON_PASSWORD.")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=not args.headful)
        context = browser.new_context()
        page = context.new_page()

        try:
            login(page, args.email, args.password, args.url)
            result = extract_table(page)
            to_excel(result, args.output)
            print(f"Exported {len(result.rows)} rows to {args.output}")
            return 0
        except TimeoutError as exc:
            raise RuntimeError(
                "Timed out while waiting for the login or table page. "
                "Try --headful in case MFA or additional prompts are shown."
            ) from exc
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    raise SystemExit(main())
