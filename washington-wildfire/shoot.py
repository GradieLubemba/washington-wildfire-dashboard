#!/usr/bin/env python3
"""
Capture real screenshots of the running dashboard for the slide deck.

Requires the local server running at http://localhost:8000/ and Google
Chrome installed. Uses Playwright driving the installed Chrome (channel).

Outputs PNGs into assets/:
    shot_map.png        - statewide fire map card
    shot_trends.png     - historical trends card
    shot_priorities.png - State Preparedness Priorities section
    shot_full.png       - full page (reference)
"""

import os
from playwright.sync_api import sync_playwright

URL = "http://localhost:8000/"
OUT = "assets"
os.makedirs(OUT, exist_ok=True)


def shoot(page, selector, path, label):
    el = page.query_selector(selector)
    if not el:
        print(f"  WARNING: selector not found: {selector} ({label})")
        return False
    el.scroll_into_view_if_needed()
    page.wait_for_timeout(600)
    el.screenshot(path=path)
    print(f"  saved {path} ({label})")
    return True


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 1600},
                                device_scale_factor=2)
        print("Loading", URL)
        page.goto(URL, wait_until="networkidle", timeout=60000)

        # Wait for data-driven content to render.
        page.wait_for_selector("#priorityCards .pcard", timeout=30000)
        page.wait_for_selector("#causeChart", timeout=30000)
        # Give Leaflet tiles + Chart.js animations time to settle.
        page.wait_for_timeout(3500)

        # State Preparedness Priorities whole section.
        shoot(page, "#priorities", f"{OUT}/shot_priorities.png", "priorities")

        # Use robust text-based locators for the cards.
        mapcard = page.locator(".card", has_text="Statewide fire locations").first
        mapcard.scroll_into_view_if_needed()
        page.wait_for_timeout(800)
        mapcard.screenshot(path=f"{OUT}/shot_map.png")
        print(f"  saved {OUT}/shot_map.png (map)")

        trends = page.locator(".card", has_text="Historical trends").first
        trends.scroll_into_view_if_needed()
        page.wait_for_timeout(600)
        trends.screenshot(path=f"{OUT}/shot_trends.png")
        print(f"  saved {OUT}/shot_trends.png (trends)")

        page.screenshot(path=f"{OUT}/shot_full.png", full_page=True)
        print(f"  saved {OUT}/shot_full.png (full page)")

        browser.close()


if __name__ == "__main__":
    main()
