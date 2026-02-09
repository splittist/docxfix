"""CLI application for docxfix using Typer."""

from typing import Annotated

import typer

app = typer.Typer(
    name="docxfix",
    help=(
        "A CLI utility for creating docx fixtures with "
        "desirable characteristics for testing."
    ),
)


@app.command()
def create(
    output: Annotated[
        str,
        typer.Argument(help="Output path for the generated docx fixture"),
    ],
    template: Annotated[
        str | None,
        typer.Option(
            "--template",
            "-t",
            help="Template docx file to use as base",
        ),
    ] = None,
    verbose: Annotated[
        bool,
        typer.Option(
            "--verbose",
            "-v",
            help="Enable verbose output",
        ),
    ] = False,
) -> None:
    """Create a new docx fixture."""
    if verbose:
        typer.echo(f"Creating docx fixture at: {output}")
        if template:
            typer.echo(f"Using template: {template}")

    # Placeholder for actual implementation
    typer.echo("✓ Docx fixture created successfully!")


@app.command()
def info() -> None:
    """Display information about docxfix."""
    from docxfix import __version__

    typer.echo(f"docxfix version {__version__}")
    typer.echo(
        "A CLI utility for creating docx fixtures with "
        "desirable characteristics for testing."
    )


def main() -> None:
    """Entry point for the CLI application."""
    app()


if __name__ == "__main__":
    main()
