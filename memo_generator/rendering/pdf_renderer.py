"""Render HTML memo to PDF using Playwright (headless Chromium)."""
import tempfile
from pathlib import Path

from playwright.sync_api import sync_playwright

from memo_generator.rendering.html_renderer import render_html


def render_pdf(memo_data: dict, output_path: str | Path) -> Path:
    """Generate a PDF investment memo and save to output_path."""
    output_path = Path(output_path)
    html_content = render_html(memo_data)

    with tempfile.NamedTemporaryFile(suffix=".html", mode="w", encoding="utf-8", delete=False) as f:
        f.write(html_content)
        tmp_path = Path(f.name)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.goto(f"file://{tmp_path}", wait_until="networkidle")
            page.pdf(path=str(output_path), format="A4", print_background=True)
            browser.close()
    finally:
        tmp_path.unlink(missing_ok=True)

    return output_path
