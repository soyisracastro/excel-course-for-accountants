"""
Clase base para generar presentaciones (.md para gamma.app) y scripts de teleprompter (.md).
"""
from pathlib import Path
from scripts.config.constants import (
    CURSO_NOMBRE, INSTRUCTOR, INSTRUCTOR_TITULO, ANIO
)


class SlidesGenerator:
    """Base class for generating Markdown slides (gamma.app) with teleprompter script."""

    def __init__(self, filename: str, output_dir: Path, script_filename: str = "",
                 script_dir: Path = None):
        self.filename = filename
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.slide_lines: list[str] = []
        self.script_lines: list[str] = []
        self.script_filename = script_filename
        self.script_dir = script_dir or output_dir
        self.script_dir.mkdir(parents=True, exist_ok=True)
        self.slide_count = 0

    @property
    def filepath(self) -> Path:
        return self.output_dir / self.filename

    @property
    def script_filepath(self) -> Path:
        return self.script_dir / self.script_filename

    def add_title_slide(self, modulo_num, modulo_nombre, subtitulo=""):
        """Add a module title slide."""
        self.slide_count += 1

        # Slide content for gamma.app
        self.slide_lines.append(f"# {'MÓDULO ' + str(modulo_num) + ': ' if modulo_num else ''}{modulo_nombre}\n")
        if subtitulo:
            self.slide_lines.append(f"*{subtitulo}*\n")
        self.slide_lines.append(f"{INSTRUCTOR} — {INSTRUCTOR_TITULO} — {CURSO_NOMBRE} {ANIO}\n")

        # Teleprompter script
        self.script_lines.append(f"# Módulo {modulo_num}: {modulo_nombre}\n")
        self.script_lines.append(f"## Slide 1 — Portada\n")

    def add_content_slide(self, title, bullets=None, body_text=None,
                          script_text="", note=""):
        """Add a standard content slide with title + bullets or body text."""
        self.slide_count += 1

        # Slide content for gamma.app
        self.slide_lines.append(f"## {title}\n")
        if bullets:
            for item in bullets:
                self.slide_lines.append(f"- {item}")
            self.slide_lines.append("")
        elif body_text:
            self.slide_lines.append(f"{body_text}\n")

        # Teleprompter script
        if script_text:
            self.script_lines.append(f"## Slide {self.slide_count} — {title}\n")
            self.script_lines.append(f"{script_text}\n")

    def add_closing_slide(self, next_module="", resources=None):
        """Add a closing/resources slide."""
        self.slide_count += 1

        # Slide content for gamma.app
        self.slide_lines.append("## Recursos y Siguiente Paso\n")
        items = resources or []
        if next_module:
            items.append(f"Siguiente: {next_module}")
        for item in items:
            self.slide_lines.append(f"- {item}")
        self.slide_lines.append("")
        self.slide_lines.append(f"*{CURSO_NOMBRE} — {INSTRUCTOR}*\n")

        # Teleprompter script
        self.script_lines.append(f"## Slide {self.slide_count} — Cierre\n")

    def save(self):
        """Save slides Markdown and teleprompter script."""
        # Save slides markdown (for gamma.app)
        slides_content = "\n".join(self.slide_lines)
        self.filepath.write_text(slides_content, encoding="utf-8")
        print(f"  ✓ {self.filepath.name}")

        # Save teleprompter script
        if self.script_filename and self.script_lines:
            script_path = self.script_filepath
            script_path.write_text("\n".join(self.script_lines), encoding="utf-8")
            print(f"  ✓ {script_path.name}")

        return self.filepath
