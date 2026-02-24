"""
Clase base para generar documentos Markdown (.md).
Reemplaza a PDFGenerator para generar contenido compatible con Notion.
"""
import re
from pathlib import Path

from scripts.config.constants import CURSO_NOMBRE, INSTRUCTOR, ANIO


class MarkdownGenerator:
    """Base class for generating Markdown documents."""

    def __init__(self, filename: str, output_dir: Path, title: str = ""):
        self.filename = filename
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.title = title
        self.lines: list[str] = []

    @property
    def filepath(self) -> Path:
        return self.output_dir / self.filename

    @staticmethod
    def _clean_html(text: str) -> str:
        """Convert simple HTML tags to Markdown equivalents."""
        text = re.sub(r"<b>(.*?)</b>", r"**\1**", text)
        text = re.sub(r"<i>(.*?)</i>", r"*\1*", text)
        text = re.sub(r"<code>(.*?)</code>", r"`\1`", text)
        text = re.sub(r"<br\s*/?>", "\n", text)
        text = re.sub(r"<[^>]+>", "", text)
        return text

    def add_cover(self, title: str, subtitle: str = "", modulo: str = ""):
        """Add a cover section."""
        self.lines.append(f"# {title}\n")
        if modulo:
            self.lines.append(f"**{modulo}**\n")
        if subtitle:
            self.lines.append(f"*{subtitle}*\n")
        self.lines.append(f"{INSTRUCTOR} | {CURSO_NOMBRE} | {ANIO}\n")
        self.lines.append("---\n")

    def add_section(self, title: str):
        """Add a section header."""
        self.lines.append(f"## {title}\n")

    def add_subsection(self, title: str):
        """Add a subsection header."""
        self.lines.append(f"### {title}\n")

    def add_text(self, text: str):
        """Add a paragraph of text."""
        self.lines.append(f"{self._clean_html(text)}\n")

    def add_bullet(self, text: str):
        """Add a bullet point."""
        self.lines.append(f"- {self._clean_html(text)}")

    def add_code(self, text: str):
        """Add a code block."""
        self.lines.append(f"```\n{text}\n```\n")

    def add_spacer(self, height=0.2):
        """Add vertical space (blank line)."""
        self.lines.append("")

    def add_page_break(self):
        """Add a page break (horizontal rule)."""
        if self.lines and self.lines[-1] != "":
            self.lines.append("")
        self.lines.append("---\n")

    def add_table(self, data, col_widths=None, header=True):
        """Add a Markdown table. col_widths is ignored (kept for API compat)."""
        if not data:
            return

        if header and len(data) > 1:
            header_row = data[0]
            body_rows = data[1:]
        elif header and len(data) == 1:
            header_row = data[0]
            body_rows = []
        else:
            # No header: use first row values as header anyway for valid Markdown
            header_row = data[0]
            body_rows = data[1:]

        def _cell(value):
            """Sanitize a cell value for Markdown tables."""
            # Replace newlines with <br> to keep content in one table row
            return str(value).replace("\n", "<br>")

        # Build header
        self.lines.append("| " + " | ".join(_cell(c) for c in header_row) + " |")
        self.lines.append("| " + " | ".join("---" for _ in header_row) + " |")

        # Build body
        for row in body_rows:
            # Pad row to match header length
            padded = list(row) + [""] * (len(header_row) - len(row))
            self.lines.append("| " + " | ".join(_cell(c) for c in padded[:len(header_row)]) + " |")

        self.lines.append("")

    def save(self):
        """Save Markdown file."""
        content = "\n".join(self.lines)
        self.filepath.write_text(content, encoding="utf-8")
        print(f"  ✓ {self.filepath.name}")
        return self.filepath
