"""CLI entry point for the Real Estate Investment Memo Generator."""
import json
from pathlib import Path

import click
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn

from memo_generator.ai.generator import generate_memo_data
from memo_generator.models.property_input import PropertyInput

console = Console()


@click.group()
def cli():
    """Real Estate Investment Memo Generator — powered by Claude AI."""
    pass


@cli.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.option("--output", "-o", default="memo.pdf", help="Output file path")
@click.option(
    "--format", "-f",
    type=click.Choice(["pdf", "markdown", "excel", "pptx"]),
    default="pdf",
    help="Output format",
)
def generate(input_file: str, output: str, format: str):
    """Generate an investment memo from INPUT_FILE (JSON)."""
    try:
        raw = json.loads(Path(input_file).read_text())
        prop = PropertyInput(**raw)
    except Exception as e:
        console.print(f"[red]Input validation error:[/red] {e}")
        raise SystemExit(1)

    console.print(f"[bold blue]Generating memo for:[/bold blue] {prop.property_name}")

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
    ) as progress:
        task = progress.add_task("Calling Claude API...", total=None)
        memo_data = generate_memo_data(prop)
        progress.update(task, description="Rendering output...")

        if format == "pdf":
            from memo_generator.rendering.pdf_renderer import render_pdf
            render_pdf(memo_data, output)
        elif format == "excel":
            from memo_generator.rendering.excel_renderer import render_excel
            render_excel(memo_data, output)
        elif format == "pptx":
            from memo_generator.rendering.ppt_renderer import render_pptx
            render_pptx(memo_data, output)
        else:
            from memo_generator.rendering.markdown_renderer import render_markdown
            Path(output).write_text(render_markdown(memo_data), encoding="utf-8")

    console.print(f"[green]✓ Memo saved to:[/green] {output}")


if __name__ == "__main__":
    cli()
