"""
Generador: Referencia_Modulo_4.md (~5 paginas)
Modulo 4 -- El Dashboard Inteligente y Entrega Profesional

Contenido:
  - Principios de diseno de dashboards
  - Guia paso a paso de Segmentadores (Slicers)
  - Guia de proteccion de celdas
  - Checklist de distribucion
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK
from scripts.generators.md_gen import MarkdownGenerator

OUTPUT_DIR = PACK / "Modulo_4_Dashboard"


def build():
    gen = MarkdownGenerator(
        "Referencia_Modulo_4.md",
        OUTPUT_DIR,
        title="Referencia Rápida — Módulo 4"
    )

    # ---- Portada ----
    gen.add_cover(
        title="Referencia Rápida — Módulo 4",
        modulo="El Dashboard Inteligente y Entrega Profesional",
        subtitle="Diseño, Slicers, Protección y Distribución"
    )

    # ==== Sección 1: Principios de Diseño ==================================
    gen.add_section("1. Principios de Diseño de Dashboards")

    gen.add_text(
        "Un dashboard efectivo comunica información clave de un vistazo. "
        "Sigue estos principios para crear paneles profesionales en Excel."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Regla de los 5 segundos")
    gen.add_text(
        "El mensaje principal de tu dashboard debe entenderse en 5 segundos. "
        "Si necesitas más tiempo, hay demasiada información o está mal organizada."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Jerarquía visual")
    gen.add_bullet("ARRIBA: KPIs (4-6 números clave en cuadros de colores)")
    gen.add_bullet("CENTRO: Gráficos principales (barras, líneas, dona)")
    gen.add_bullet("IZQUIERDA: Segmentadores de datos (filtros interactivos)")
    gen.add_bullet("ABAJO: Tabla de detalle o gráficos secundarios")
    gen.add_spacer(0.1)

    gen.add_subsection("Paleta de colores")
    tabla_colores = [
        ["Color", "Código Hex", "Uso"],
        ["Azul", "#2563EB", "Color principal, encabezados, títulos"],
        ["Verde", "#10B981", "Positivo, vigente, aprobado"],
        ["Rojo", "#EF4444", "Alerta, vencido, rechazado"],
        ["Amarillo", "#F59E0B", "Precaución, por vencer, en proceso"],
        ["Gris claro", "#F8FAFC", "Fondo del dashboard"],
        ["Blanco", "#FFFFFF", "Fondo de cuadros de KPI"],
    ]
    gen.add_table(tabla_colores, col_widths=[70, 80, 300])

    gen.add_spacer(0.1)

    gen.add_subsection("Tipos de gráfico recomendados")
    gen.add_bullet("Barras agrupadas: comparar categorías (ej. sueldo por puesto)")
    gen.add_bullet("Líneas: mostrar tendencias en el tiempo (ej. nómina mensual)")
    gen.add_bullet("Dona/Pie: proporción de un total (ej. % ISR vs % neto)")
    gen.add_bullet("EVITAR: gráficos 3D, demasiados colores, ejes innecesarios")

    gen.add_spacer(0.1)

    gen.add_subsection("Limpieza visual")
    gen.add_bullet("Oculta líneas de cuadrícula: Vista > desmarcar 'Líneas de cuadrícula'")
    gen.add_bullet("Usa bordes solo donde agreguen claridad (no en todas las celdas)")
    gen.add_bullet("Alinea los elementos visual y consistentemente")
    gen.add_bullet("Usa una sola familia tipográfica (Calibri recomendado)")

    gen.add_page_break()

    # ==== Sección 2: Slicers Paso a Paso ===================================
    gen.add_section("2. Segmentadores de Datos (Slicers) — Paso a Paso")

    gen.add_text(
        "Los segmentadores son filtros visuales para Tablas Dinámicas. "
        "Permiten crear dashboards interactivos sin macros ni VBA."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 1: Preparar tus datos")
    gen.add_bullet("Tus datos deben estar en formato de Tabla (Ctrl+T) o Tabla Dinámica")
    gen.add_bullet("Cada columna debe tener un encabezado único y descriptivo")
    gen.add_bullet("No debe haber filas ni columnas vacías dentro de los datos")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 2: Crear la Tabla Dinámica")
    gen.add_bullet("Selecciona cualquier celda de tu tabla de datos")
    gen.add_bullet("Menú: Insertar > Tabla Dinámica > Hoja nueva o existente")
    gen.add_bullet("Arrastra campos: filas (Periodo), valores (Suma de Sueldo, Suma de ISR)")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 3: Insertar el segmentador")
    gen.add_bullet("Haz clic en cualquier celda de tu Tabla Dinámica")
    gen.add_bullet("Menú: Insertar > Segmentación de datos")
    gen.add_bullet("Selecciona los campos que quieres filtrar: Periodo, Puesto, Empleado")
    gen.add_bullet("Haz clic en Aceptar — aparecen los botones de filtro")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 4: Vincular a múltiples TDs")
    gen.add_bullet("Clic derecho sobre el segmentador > Conexiones de informe...")
    gen.add_bullet("Marca TODAS las Tablas Dinámicas que deben responder al filtro")
    gen.add_bullet("Ahora un clic filtra todas las TDs y sus gráficos asociados")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 5: Personalizar el segmentador")
    gen.add_bullet("Selecciona el slicer > pestaña Segmentación de datos (cinta)")
    gen.add_bullet("Cambia el estilo visual, número de columnas y tamaño")
    gen.add_bullet("Consejo: usa 2-3 columnas para que ocupe menos espacio horizontal")

    gen.add_spacer(0.1)

    gen.add_subsection("Atajos útiles")
    atajo_data = [
        ["Acción", "Atajo / Método"],
        ["Seleccionar múltiples ítems", "Ctrl + clic en cada botón"],
        ["Limpiar filtro del slicer", "Ícono de embudo con X (esquina del slicer)"],
        ["Mover slicer", "Arrastrar con el mouse"],
        ["Redimensionar", "Arrastrar las esquinas"],
        ["Eliminar slicer", "Seleccionar + tecla Suprimir"],
    ]
    gen.add_table(atajo_data, col_widths=[200, 250])

    gen.add_page_break()

    # ==== Sección 3: Protección de Celdas ==================================
    gen.add_section("3. Protección de Celdas — Guía Rápida")

    gen.add_text(
        "Proteger tu archivo asegura que el usuario final interactúe correctamente "
        "con el dashboard sin romper fórmulas ni la estructura."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Flujo de protección en 4 pasos")

    pasos = [
        ["Paso", "Acción", "Dónde"],
        ["1", "Selecciona las celdas de INPUT (donde el usuario escribe)", "Celdas amarillas"],
        ["2", "Clic derecho > Formato > Protección > desmarcar 'Bloqueada'", "Cada celda de input"],
        ["3", "Revisar > Proteger hoja > contraseña > permitir filtros/TDs", "Cada hoja"],
        ["4", "Revisar > Proteger libro > contraseña (estructura)", "Una vez por archivo"],
    ]
    gen.add_table(pasos, col_widths=[35, 280, 135])

    gen.add_spacer(0.1)

    gen.add_subsection("Qué proteger y qué no")
    proteccion = [
        ["Elemento", "¿Bloquear?", "Razón"],
        ["Celdas con fórmulas", "SÍ", "Evitar que borren o modifiquen cálculos"],
        ["Celdas de input del usuario", "NO", "El usuario necesita ingresar datos"],
        ["Encabezados de tabla", "SÍ", "Mantener estructura"],
        ["Celdas de KPI con SUBTOTAL", "SÍ", "Son fórmulas críticas"],
        ["Nombre de hojas", "SÍ (libro)", "Evitar renombrar o eliminar hojas"],
    ]
    gen.add_table(proteccion, col_widths=[150, 60, 240])

    gen.add_spacer(0.1)

    gen.add_subsection("Niveles de seguridad en Excel")
    niveles = [
        ["Nivel", "Método", "Seguridad"],
        ["Básico", "Proteger hoja (sin contraseña)", "Mínima — cualquiera desprotege"],
        ["Medio", "Proteger hoja + libro (con contraseña)", "Moderada — herramientas para romper"],
        ["Alto", "Contraseña de apertura de archivo (AES)", "Alta — cifrado real"],
        ["Máximo", "Exportar como PDF (no editable)", "Máxima — sin datos editables"],
    ]
    gen.add_table(niveles, col_widths=[55, 220, 175])

    gen.add_page_break()

    # ==== Sección 4: Checklist de Distribución =============================
    gen.add_section("4. Checklist de Distribución Profesional")

    gen.add_text(
        "Usa esta lista antes de enviar cualquier archivo a clientes, "
        "jefes o colegas."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Antes de compartir")
    gen.add_bullet("[ ] Las fórmulas calculan correctamente con datos de prueba")
    gen.add_bullet("[ ] Los gráficos se actualizan al cambiar filtros/slicers")
    gen.add_bullet("[ ] Las celdas de input están desbloqueadas y resaltadas en amarillo")
    gen.add_bullet("[ ] Las celdas de fórmula están bloqueadas")
    gen.add_bullet("[ ] Las hojas están protegidas con contraseña")
    gen.add_bullet("[ ] La estructura del libro está protegida")
    gen.add_bullet("[ ] No hay datos personales o confidenciales expuestos")
    gen.add_bullet("[ ] El nombre del archivo sigue la convención profesional")
    gen.add_spacer(0.1)

    gen.add_subsection("Formato del nombre de archivo")
    gen.add_code("[Empresa]_[TipoDocumento]_[Periodo]_[Version].[ext]")
    gen.add_spacer(0.05)
    gen.add_bullet("Ejemplo: GrupoTorres_Nomina_2026_Enero_v1.xlsx")
    gen.add_bullet("Ejemplo: CNO850315_ISR_Anual_2025_Final.pdf")
    gen.add_bullet("Evita espacios, caracteres especiales y acentos en nombres de archivo")
    gen.add_spacer(0.1)

    gen.add_subsection("Qué formato usar")
    formato = [
        ["Situación", "Formato", "Razón"],
        ["Reporte final a cliente", "PDF", "No editable, aspecto profesional"],
        ["Declaración o complemento SAT", "PDF + XML", "Formato oficial"],
        ["Plantilla que el cliente debe llenar", "Excel protegido", "Mantiene fórmulas"],
        ["Dashboard interactivo para equipo", "Excel protegido", "Slicers y TDs activos"],
        ["Respaldo/archivo maestro", "Excel sin proteger", "Facilita edición futura"],
    ]
    gen.add_table(formato, col_widths=[160, 90, 200])

    gen.add_spacer(0.1)

    gen.add_subsection("Después de compartir")
    gen.add_bullet("[ ] Confirma que el destinatario puede abrir el archivo")
    gen.add_bullet("[ ] Verifica que la versión de Excel del destinatario es compatible")
    gen.add_bullet("[ ] Guarda tu copia maestra sin proteger (sufijo _master.xlsx)")
    gen.add_bullet("[ ] Documenta la contraseña de protección en lugar seguro")

    gen.add_spacer(0.3)

    gen.add_subsection("Archivos de este módulo")
    gen.add_bullet("09_Layout_Dashboard_Contable.xlsx — Template de layout para dashboard")
    gen.add_bullet("10_Dashboard_Final_Integrado.xlsx — Datos + Calculadora + Dashboard")
    gen.add_bullet("11_Guia_Proteccion_y_Seguridad.md — Guía detallada de protección")
    gen.add_bullet("Referencia_Modulo_4.md — Este documento")

    gen.save()


if __name__ == "__main__":
    build()
