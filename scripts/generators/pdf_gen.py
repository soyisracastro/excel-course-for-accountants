"""
Clase base para generar PDFs con reportlab.
"""
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch, cm
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from scripts.config.constants import (
    Color, CURSO_NOMBRE, INSTRUCTOR, ANIO
)

# ── Colores reportlab ────────────────────────────────────────────
RL_AZUL = HexColor(f"#{Color.AZUL}")
RL_VERDE = HexColor(f"#{Color.VERDE}")
RL_ROJO = HexColor(f"#{Color.ROJO}")
RL_AMARILLO = HexColor(f"#{Color.AMARILLO}")
RL_TEXTO = HexColor(f"#{Color.TEXTO_OSCURO}")
RL_TEXTO_MEDIO = HexColor(f"#{Color.TEXTO_MEDIO}")
RL_FONDO = HexColor(f"#{Color.FONDO_CLARO}")
RL_GRIS_BORDE = HexColor(f"#{Color.GRIS_BORDE}")
RL_BLANCO = HexColor(f"#{Color.BLANCO}")


class PDFGenerator:
    """Base class for generating PDF documents."""

    def __init__(self, filename: str, output_dir: Path, title: str = ""):
        self.filename = filename
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.title = title
        self.elements = []
        self._setup_styles()

    @property
    def filepath(self) -> Path:
        return self.output_dir / self.filename

    def _setup_styles(self):
        """Configure paragraph styles."""
        self.styles = getSampleStyleSheet()

        self.styles.add(ParagraphStyle(
            "PDFTitle",
            parent=self.styles["Title"],
            fontSize=22,
            textColor=RL_AZUL,
            spaceAfter=6,
            fontName="Helvetica-Bold",
        ))

        self.styles.add(ParagraphStyle(
            "PDFSubtitle",
            parent=self.styles["Normal"],
            fontSize=12,
            textColor=RL_TEXTO_MEDIO,
            spaceAfter=20,
            fontName="Helvetica",
        ))

        self.styles.add(ParagraphStyle(
            "SectionHeader",
            parent=self.styles["Heading2"],
            fontSize=14,
            textColor=RL_AZUL,
            spaceBefore=16,
            spaceAfter=8,
            fontName="Helvetica-Bold",
            borderWidth=0,
            borderPadding=0,
        ))

        self.styles.add(ParagraphStyle(
            "SubSection",
            parent=self.styles["Heading3"],
            fontSize=11,
            textColor=RL_TEXTO,
            spaceBefore=10,
            spaceAfter=4,
            fontName="Helvetica-Bold",
        ))

        self.styles.add(ParagraphStyle(
            "BodyText2",
            parent=self.styles["Normal"],
            fontSize=10,
            textColor=RL_TEXTO,
            spaceAfter=6,
            fontName="Helvetica",
            leading=14,
            alignment=TA_JUSTIFY,
        ))

        self.styles.add(ParagraphStyle(
            "BulletItem",
            parent=self.styles["Normal"],
            fontSize=10,
            textColor=RL_TEXTO,
            spaceAfter=3,
            fontName="Helvetica",
            leftIndent=20,
            bulletIndent=10,
            leading=13,
        ))

        self.styles.add(ParagraphStyle(
            "CodeBlock",
            parent=self.styles["Code"],
            fontSize=9,
            textColor=RL_TEXTO,
            backColor=RL_FONDO,
            fontName="Courier",
            leftIndent=12,
            rightIndent=12,
            spaceBefore=4,
            spaceAfter=4,
            leading=12,
        ))

        self.styles.add(ParagraphStyle(
            "FooterText",
            parent=self.styles["Normal"],
            fontSize=8,
            textColor=RL_TEXTO_MEDIO,
            fontName="Helvetica",
            alignment=TA_CENTER,
        ))

    def add_cover(self, title: str, subtitle: str = "", modulo: str = ""):
        """Add a cover page."""
        self.elements.append(Spacer(1, 2 * inch))
        self.elements.append(Paragraph(title, self.styles["PDFTitle"]))
        if modulo:
            self.elements.append(Paragraph(modulo, self.styles["PDFSubtitle"]))
        if subtitle:
            self.elements.append(Paragraph(subtitle, self.styles["PDFSubtitle"]))
        self.elements.append(Spacer(1, 0.5 * inch))
        self.elements.append(HRFlowable(
            width="80%", thickness=2, color=RL_AZUL, spaceAfter=12
        ))
        self.elements.append(Paragraph(
            f"{INSTRUCTOR} | {CURSO_NOMBRE} | {ANIO}",
            self.styles["FooterText"]
        ))
        self.elements.append(PageBreak())

    def add_section(self, title: str):
        """Add a section header."""
        self.elements.append(Paragraph(title, self.styles["SectionHeader"]))
        self.elements.append(HRFlowable(
            width="100%", thickness=1, color=RL_GRIS_BORDE, spaceAfter=8
        ))

    def add_subsection(self, title: str):
        self.elements.append(Paragraph(title, self.styles["SubSection"]))

    def add_text(self, text: str):
        self.elements.append(Paragraph(text, self.styles["BodyText2"]))

    def add_bullet(self, text: str):
        self.elements.append(Paragraph(
            f"• {text}", self.styles["BulletItem"]
        ))

    def add_code(self, text: str):
        self.elements.append(Paragraph(text, self.styles["CodeBlock"]))

    def add_spacer(self, height=0.2):
        self.elements.append(Spacer(1, height * inch))

    def add_page_break(self):
        self.elements.append(PageBreak())

    def add_table(self, data, col_widths=None, header=True):
        """Add a styled table."""
        style_cmds = [
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("TEXTCOLOR", (0, 0), (-1, -1), RL_TEXTO),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, RL_GRIS_BORDE),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ]
        if header and len(data) > 0:
            style_cmds.extend([
                ("BACKGROUND", (0, 0), (-1, 0), RL_AZUL),
                ("TEXTCOLOR", (0, 0), (-1, 0), RL_BLANCO),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
            ])
            # Zebra
            for i in range(2, len(data), 2):
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), RL_FONDO))

        tbl = Table(data, colWidths=col_widths, repeatRows=1 if header else 0)
        tbl.setStyle(TableStyle(style_cmds))
        self.elements.append(tbl)
        self.elements.append(Spacer(1, 0.15 * inch))

    def _footer(self, canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(RL_TEXTO_MEDIO)
        canvas.drawCentredString(
            letter[0] / 2, 0.5 * inch,
            f"{CURSO_NOMBRE} — {INSTRUCTOR} — {ANIO}  |  Pág. {doc.page}"
        )
        canvas.restoreState()

    def save(self):
        """Build and save PDF."""
        doc = SimpleDocTemplate(
            str(self.filepath),
            pagesize=letter,
            topMargin=0.75 * inch,
            bottomMargin=0.75 * inch,
            leftMargin=0.75 * inch,
            rightMargin=0.75 * inch,
            title=self.title,
            author=INSTRUCTOR,
        )
        doc.build(self.elements, onFirstPage=self._footer, onLaterPages=self._footer)
        print(f"  ✓ {self.filepath.name}")
        return self.filepath
