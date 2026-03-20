"""Render HTML memo to PDF using WeasyPrint."""
from pathlib import Path

from weasyprint import HTML

from memo_generator.rendering.html_renderer import render_html


def render_pdf(memo_data: dict, output_path: str | Path) -> Path:
    """Generate a PDF investment memo and save to output_path."""
    output_path = Path(output_path)
    html_content = render_html(memo_data)
    HTML(string=html_content).write_pdf(str(output_path))
    return output_path
