"""Render memo data to HTML using Jinja2."""
from datetime import date
from pathlib import Path

from jinja2 import Environment, FileSystemLoader

_TEMPLATES_DIR = Path(__file__).parent.parent / "templates"
_CSS_PATH = _TEMPLATES_DIR / "styles" / "memo.css"


def render_html(memo_data: dict) -> str:
    """Render the full memo as an HTML string."""
    env = Environment(loader=FileSystemLoader(str(_TEMPLATES_DIR)))
    template = env.get_template("memo_base.html")

    prop = memo_data["property"]
    metrics = memo_data["metrics"]
    sections = memo_data["sections"]

    # Flatten metrics for template convenience
    flat_metrics = {
        **metrics,
        "purchase_price": prop.purchase_price,
    }

    return template.render(
        property=prop,
        metrics=flat_metrics,
        sections=sections,
        css_path=str(_CSS_PATH),
        prepared_date=date.today().strftime("%B %Y"),
    )
