"""
Generador de thumbnails para YouTube — Curso Excel para Contadores
Genera 7 thumbnails consistentes (Módulos 1-5 + 2 Bonus) en 1280x720px.

Uso:
    python3 -m scripts.generate_thumbnails
"""

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

# ── Configuración ────────────────────────────────────────────────────
WIDTH, HEIGHT = 1280, 720
OUTPUT_DIR = Path(__file__).resolve().parents[1] / "output" / "Thumbnails"

# Fuentes del sistema (macOS)
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
FONT_BLACK = "/System/Library/Fonts/Supplemental/Arial Black.ttf"
FONT_REGULAR = "/System/Library/Fonts/Supplemental/Arial.ttf"
FONT_IMPACT = "/System/Library/Fonts/Supplemental/Impact.ttf"

# ── Datos de cada thumbnail ──────────────────────────────────────────
THUMBNAILS = [
    {
        "id": "M1",
        "label": "MÓDULO 1",
        "numero": "01",
        "titulo_l1": "LÓGICA CONTABLE",
        "titulo_l2": "Y FUNCIONES",
        "subtitulo": "ISR 2026 | e.firma | RFC",
        "color_acento": (37, 99, 235),     # Azul
        "color_fondo2": (15, 50, 120),
        "icono": "fx",
    },
    {
        "id": "M2",
        "label": "MÓDULO 2",
        "numero": "02",
        "titulo_l1": "TABLAS",
        "titulo_l2": "DINÁMICAS",
        "subtitulo": "Nómina XML | Análisis masivo",
        "color_acento": (16, 185, 129),     # Verde
        "color_fondo2": (8, 90, 65),
        "icono": "TBL",
    },
    {
        "id": "M3",
        "label": "MÓDULO 3",
        "numero": "03",
        "titulo_l1": "GRÁFICAS Y",
        "titulo_l2": "REPORTES",
        "subtitulo": "Dashboard combustible | Comparativa anual",
        "color_acento": (245, 158, 11),     # Amarillo
        "color_fondo2": (120, 75, 5),
        "icono": "BAR",
    },
    {
        "id": "M4",
        "label": "MÓDULO 4",
        "numero": "04",
        "titulo_l1": "DASHBOARD",
        "titulo_l2": "INTELIGENTE",
        "subtitulo": "KPIs | Protección | Entrega profesional",
        "color_acento": (239, 68, 68),      # Rojo
        "color_fondo2": (120, 30, 30),
        "icono": "KPI",
    },
    {
        "id": "M5",
        "label": "MÓDULO 5",
        "numero": "05",
        "titulo_l1": "COPILOT &",
        "titulo_l2": "EXCEL IA",
        "subtitulo": "Microsoft 365 | Prompts contables",
        "color_acento": (139, 92, 246),     # Violeta
        "color_fondo2": (70, 40, 130),
        "icono": "AI",
    },
    {
        "id": "B1",
        "label": "BONUS",
        "numero": "B1",
        "titulo_l1": "VBA CON",
        "titulo_l2": "IA",
        "subtitulo": "Macros + ChatGPT + Claude",
        "color_acento": (236, 72, 153),     # Rosa
        "color_fondo2": (120, 30, 75),
        "icono": "VBA",
    },
    {
        "id": "B2",
        "label": "BONUS",
        "numero": "B2",
        "titulo_l1": "CLAUDE EN",
        "titulo_l2": "EXCEL",
        "subtitulo": "Tu segundo cerebro contable",
        "color_acento": (217, 119, 87),     # Terracota/Claude
        "color_fondo2": (110, 55, 40),
        "icono": "CL",
    },
]


def _lerp_color(c1, c2, t):
    """Interpolación lineal entre dos colores RGB."""
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


def _draw_rounded_rect(draw, xy, radius, fill):
    """Dibuja un rectángulo con esquinas redondeadas."""
    x0, y0, x1, y1 = xy
    draw.rounded_rectangle(xy, radius=radius, fill=fill)


def _draw_gradient_background(img, color_dark, color_accent):
    """Dibuja un fondo con gradiente diagonal oscuro."""
    draw = ImageDraw.Draw(img)
    base_dark = (18, 18, 28)
    for y in range(HEIGHT):
        t = y / HEIGHT
        # Gradiente vertical de oscuro a ligeramente más claro
        row_color = _lerp_color(base_dark, (30, 30, 45), t)
        draw.line([(0, y), (WIDTH, y)], fill=row_color)


def _draw_accent_bar(draw, color_acento):
    """Barra de acento en la parte inferior."""
    draw.rectangle([(0, HEIGHT - 8), (WIDTH, HEIGHT)], fill=color_acento)


def _draw_accent_glow(img, color_acento, color_fondo2):
    """Dibuja un resplandor sutil del color de acento en esquina superior derecha."""
    overlay = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    draw_ov = ImageDraw.Draw(overlay)
    # Círculo grande difuso en la esquina superior derecha
    cx, cy = WIDTH - 100, -50
    for r in range(400, 0, -2):
        alpha = int(35 * (1 - r / 400))
        c = color_acento + (alpha,)
        draw_ov.ellipse(
            [cx - r, cy - r, cx + r, cy + r],
            fill=c,
        )
    # Segundo resplandor más pequeño abajo-izquierda
    cx2, cy2 = 150, HEIGHT - 100
    for r in range(250, 0, -2):
        alpha = int(20 * (1 - r / 250))
        c = color_fondo2 + (alpha,)
        draw_ov.ellipse(
            [cx2 - r, cy2 - r, cx2 + r, cy2 + r],
            fill=c,
        )
    img.paste(Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB"), (0, 0))


def _draw_number_big(draw, numero, color_acento):
    """Número grande semi-transparente como elemento decorativo de fondo."""
    try:
        font_num = ImageFont.truetype(FONT_BLACK, 320)
    except OSError:
        font_num = ImageFont.truetype(FONT_IMPACT, 320)
    # Posición en la esquina derecha
    text = numero
    bbox = draw.textbbox((0, 0), text, font=font_num)
    tw = bbox[2] - bbox[0]
    x = WIDTH - tw - 40
    y = 100
    # Sombra oscura
    r, g, b = color_acento
    shadow_color = (r // 6, g // 6, b // 6)
    draw.text((x + 4, y + 4), text, fill=shadow_color, font=font_num)
    # Número con opacidad simulada (color más oscuro que el acento)
    dim_color = (r // 3, g // 3, b // 3)
    draw.text((x, y), text, fill=dim_color, font=font_num)


def _draw_module_label(draw, label, color_acento):
    """Etiqueta 'MÓDULO X' o 'BONUS' con pill de color."""
    font_label = ImageFont.truetype(FONT_BOLD, 28)
    text = label
    bbox = draw.textbbox((0, 0), text, font=font_label)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    pill_x = 70
    pill_y = 80
    padding_h = 20
    padding_v = 10
    _draw_rounded_rect(
        draw,
        (pill_x, pill_y, pill_x + tw + padding_h * 2, pill_y + th + padding_v * 2),
        radius=8,
        fill=color_acento,
    )
    draw.text(
        (pill_x + padding_h, pill_y + padding_v - 2),
        text,
        fill=(255, 255, 255),
        font=font_label,
    )
    return pill_y + th + padding_v * 2


def _draw_title(draw, titulo_l1, titulo_l2, y_start):
    """Título principal en dos líneas, grande y bold."""
    font_title = ImageFont.truetype(FONT_BLACK, 82)
    y = y_start + 30
    # Sombra
    draw.text((73, y + 3), titulo_l1, fill=(0, 0, 0), font=font_title)
    draw.text((70, y), titulo_l1, fill=(255, 255, 255), font=font_title)
    y += 90
    draw.text((73, y + 3), titulo_l2, fill=(0, 0, 0), font=font_title)
    draw.text((70, y), titulo_l2, fill=(255, 255, 255), font=font_title)
    return y + 90


def _draw_subtitle(draw, subtitulo, y_start, color_acento):
    """Subtítulo descriptivo debajo del título."""
    font_sub = ImageFont.truetype(FONT_REGULAR, 30)
    y = y_start + 20
    # Color de acento para el subtítulo
    r, g, b = color_acento
    sub_color = (min(r + 80, 255), min(g + 80, 255), min(b + 80, 255))
    draw.text((72, y), subtitulo, fill=sub_color, font=font_sub)
    return y + 40


def _draw_branding(draw, color_acento):
    """Branding inferior: nombre del curso + todoconta.com."""
    font_brand = ImageFont.truetype(FONT_BOLD, 24)
    font_url = ImageFont.truetype(FONT_BOLD, 26)

    y_brand = HEIGHT - 70

    # Nombre del curso — lado izquierdo
    draw.text(
        (70, y_brand),
        "EXCEL PARA CONTADORES Y ADMINISTRATIVOS",
        fill=(180, 180, 195),
        font=font_brand,
    )

    # URL — lado derecho
    url_text = "todoconta.com"
    bbox = draw.textbbox((0, 0), url_text, font=font_url)
    tw = bbox[2] - bbox[0]
    draw.text(
        (WIDTH - tw - 70, y_brand),
        url_text,
        fill=color_acento,
        font=font_url,
    )


def _draw_excel_badge(draw, color_acento):
    """Badge 'EXCEL' estilizado como referencia visual."""
    font_badge = ImageFont.truetype(FONT_BLACK, 36)
    badge_text = "EXCEL"
    bbox = draw.textbbox((0, 0), badge_text, font=font_badge)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    text_offset_y = bbox[1]  # Offset real del ascender
    pad_h, pad_v = 18, 14
    # Definir el rectángulo del badge
    badge_w = tw + pad_h * 2
    badge_h = th + pad_v * 2
    badge_x = WIDTH - badge_w - 80
    badge_y = 80
    # Fondo verde Excel
    _draw_rounded_rect(
        draw,
        (badge_x, badge_y, badge_x + badge_w, badge_y + badge_h),
        radius=10,
        fill=(33, 115, 70),  # Verde Excel clásico
    )
    # Centrar texto dentro del badge
    text_x = badge_x + (badge_w - tw) // 2 - bbox[0]
    text_y = badge_y + (badge_h - th) // 2 - text_offset_y
    draw.text((text_x, text_y), badge_text, fill=(255, 255, 255), font=font_badge)


def _draw_icon_element(draw, icono, color_acento):
    """Elemento decorativo con el ícono/texto representativo del módulo."""
    font_icon = ImageFont.truetype(FONT_BLACK, 48)
    bbox = draw.textbbox((0, 0), icono, font=font_icon)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    # Posicionar al lado derecho, centrado verticalmente
    ix = WIDTH - tw - 110
    iy = HEIGHT // 2 + 50
    pad = 20
    r, g, b = color_acento
    bg_color = (r // 4, g // 4, b // 4)
    _draw_rounded_rect(
        draw,
        (ix - pad, iy - pad, ix + tw + pad, iy + th + pad),
        radius=12,
        fill=bg_color,
    )
    # Borde del ícono
    draw.rounded_rectangle(
        (ix - pad, iy - pad, ix + tw + pad, iy + th + pad),
        radius=12,
        outline=color_acento,
        width=2,
    )
    draw.text((ix, iy - 2), icono, fill=color_acento, font=font_icon)


def _draw_separator_line(draw, y, color_acento):
    """Línea decorativa horizontal."""
    r, g, b = color_acento
    line_color = (r // 2, g // 2, b // 2)
    draw.line([(70, y), (500, y)], fill=line_color, width=3)


def generate_thumbnail(data: dict, output_dir: Path) -> Path:
    """Genera un thumbnail individual y lo guarda como PNG."""
    img = Image.new("RGB", (WIDTH, HEIGHT), (18, 18, 28))
    draw = ImageDraw.Draw(img)

    color_acento = data["color_acento"]
    color_fondo2 = data["color_fondo2"]

    # 1. Fondo con gradiente
    _draw_gradient_background(img, (18, 18, 28), color_acento)

    # 2. Resplandor de acento
    _draw_accent_glow(img, color_acento, color_fondo2)

    # Redibujar draw después de la composición
    draw = ImageDraw.Draw(img)

    # 3. Número grande decorativo de fondo
    _draw_number_big(draw, data["numero"], color_acento)

    # 4. Barra de acento inferior
    _draw_accent_bar(draw, color_acento)

    # 5. Etiqueta de módulo (pill)
    label_bottom = _draw_module_label(draw, data["label"], color_acento)

    # 6. Título principal
    title_bottom = _draw_title(draw, data["titulo_l1"], data["titulo_l2"], label_bottom)

    # 7. Línea separadora
    sep_y = title_bottom + 15
    _draw_separator_line(draw, sep_y, color_acento)

    # 8. Subtítulo
    _draw_subtitle(draw, data["subtitulo"], sep_y + 5, color_acento)

    # 9. Badge EXCEL
    _draw_excel_badge(draw, color_acento)

    # 10. Ícono decorativo
    _draw_icon_element(draw, data["icono"], color_acento)

    # 11. Branding inferior
    _draw_branding(draw, color_acento)

    # Guardar
    filename = f"Thumbnail_{data['id']}.png"
    filepath = output_dir / filename
    img.save(filepath, "PNG", quality=95)
    return filepath


def build():
    """Genera todos los thumbnails."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"\n{'='*60}")
    print("  GENERANDO THUMBNAILS PARA YOUTUBE")
    print(f"{'='*60}\n")

    generated = []
    for data in THUMBNAILS:
        path = generate_thumbnail(data, OUTPUT_DIR)
        generated.append(path)
        print(f"  [OK] {path.name}")

    print(f"\n  {len(generated)} thumbnails generados en:")
    print(f"  {OUTPUT_DIR}/")
    print(f"\n{'='*60}\n")
    return generated


if __name__ == "__main__":
    build()
