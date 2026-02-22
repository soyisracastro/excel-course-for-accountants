"""
Constantes globales del curso "Excel para Contadores y Administrativos"
Colores, fuentes, metadatos y rutas.
"""
from pathlib import Path

# ── Rutas base ──────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / "output"
PACK = OUTPUT / "Pack_Excel_Pro"
SLIDES_DIR = OUTPUT / "Slides"
TELEPROMPTER_DIR = OUTPUT / "Teleprompter"

# ── Metadatos del curso ────────────────────────────────────────────
CURSO_NOMBRE = "Excel para Contadores y Administrativos"
CURSO_SUBTITULO = "Automatiza tu Trabajo Contable y Administrativo"
INSTRUCTOR = "Israel Castro"
INSTRUCTOR_TITULO = "CPA & Software Engineer"
SITIO_WEB = "todoconta.com"
ANIO = 2026

# ── Paleta de colores (hex) ────────────────────────────────────────
class Color:
    FONDO_CLARO = "F8FAFC"
    FONDO_MEDIO = "F1F5F9"
    AZUL = "2563EB"
    VERDE = "10B981"
    ROJO = "EF4444"
    AMARILLO = "F59E0B"
    TEXTO_OSCURO = "1E293B"
    TEXTO_MEDIO = "475569"
    BLANCO = "FFFFFF"
    NEGRO = "000000"
    GRIS_BORDE = "CBD5E1"

# Colores como tuplas RGB (0-255) para reportlab y python-pptx
class ColorRGB:
    FONDO_CLARO = (248, 250, 252)
    FONDO_MEDIO = (241, 245, 249)
    AZUL = (37, 99, 235)
    VERDE = (16, 185, 129)
    ROJO = (239, 68, 68)
    AMARILLO = (245, 158, 11)
    TEXTO_OSCURO = (30, 41, 59)
    TEXTO_MEDIO = (71, 85, 105)
    BLANCO = (255, 255, 255)
    NEGRO = (0, 0, 0)
    GRIS_BORDE = (203, 213, 225)

# ── Fuentes ────────────────────────────────────────────────────────
FONT_TITULO = "Calibri"
FONT_CUERPO = "Calibri"
FONT_MONO = "Consolas"

# ── Combustibles (precios México 2026) ─────────────────────────────
COMBUSTIBLES = {
    "Magna": {"precio_min": 23.0, "precio_max": 25.0, "color": "10B981"},
    "Premium": {"precio_min": 25.0, "precio_max": 27.0, "color": "EF4444"},
    "Diésel": {"precio_min": 25.0, "precio_max": 27.0, "color": "64748B"},
}

# ── Salario mínimo 2026 ───────────────────────────────────────────
SALARIO_MINIMO_DIARIO = 278.80
SALARIO_MINIMO_MENSUAL = SALARIO_MINIMO_DIARIO * 30

# ── Módulos del curso ─────────────────────────────────────────────
MODULOS = {
    1: {
        "nombre": "Lógica Contable y Funciones de Control",
        "carpeta": "Modulo_1_Funciones",
        "slide_nombre": "Modulo_1_Logica_Contable.md",
        "script_nombre": "Script_Modulo_1.md",
    },
    2: {
        "nombre": "Procesamiento Masivo y Análisis con Tablas Dinámicas",
        "carpeta": "Modulo_2_Tablas_Dinamicas",
        "slide_nombre": "Modulo_2_Tablas_Dinamicas.md",
        "script_nombre": "Script_Modulo_2.md",
    },
    3: {
        "nombre": "Visualización de Impacto y Reportes Ejecutivos",
        "carpeta": "Modulo_3_Visualizacion",
        "slide_nombre": "Modulo_3_Visualizacion.md",
        "script_nombre": "Script_Modulo_3.md",
    },
    4: {
        "nombre": "El Dashboard Inteligente y Entrega Profesional",
        "carpeta": "Modulo_4_Dashboard",
        "slide_nombre": "Modulo_4_Dashboard.md",
        "script_nombre": "Script_Modulo_4.md",
    },
    5: {
        "nombre": "Automatización Nativa con Microsoft 365 Copilot",
        "carpeta": "Modulo_5_Copilot_IA",
        "slide_nombre": "Modulo_5_Copilot_IA.md",
        "script_nombre": "Script_Modulo_5.md",
    },
}
