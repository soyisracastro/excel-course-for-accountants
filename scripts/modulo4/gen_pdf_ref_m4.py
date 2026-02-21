"""
Generador: Referencia_Modulo_4.pdf (~5 paginas)
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

from reportlab.lib.units import inch
from scripts.config.constants import PACK
from scripts.generators.pdf_gen import PDFGenerator

OUTPUT_DIR = PACK / "Modulo_4_Dashboard"


def build():
    gen = PDFGenerator(
        "Referencia_Modulo_4.pdf",
        OUTPUT_DIR,
        title="Referencia Rapida -- Modulo 4"
    )

    # ---- Portada ----
    gen.add_cover(
        title="Referencia Rapida -- Modulo 4",
        modulo="El Dashboard Inteligente y Entrega Profesional",
        subtitle="Diseno, Slicers, Proteccion y Distribucion"
    )

    # ==== Seccion 1: Principios de Diseno ==================================
    gen.add_section("1. Principios de Diseno de Dashboards")

    gen.add_text(
        "Un dashboard efectivo comunica informacion clave de un vistazo. "
        "Sigue estos principios para crear paneles profesionales en Excel."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Regla de los 5 segundos")
    gen.add_text(
        "El mensaje principal de tu dashboard debe entenderse en 5 segundos. "
        "Si necesitas mas tiempo, hay demasiada informacion o esta mal organizada."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Jerarquia visual")
    gen.add_bullet("ARRIBA: KPIs (4-6 numeros clave en cuadros de colores)")
    gen.add_bullet("CENTRO: Graficos principales (barras, lineas, dona)")
    gen.add_bullet("IZQUIERDA: Segmentadores de datos (filtros interactivos)")
    gen.add_bullet("ABAJO: Tabla de detalle o graficos secundarios")
    gen.add_spacer(0.1)

    gen.add_subsection("Paleta de colores")
    tabla_colores = [
        ["Color", "Codigo Hex", "Uso"],
        ["Azul", "#2563EB", "Color principal, encabezados, titulos"],
        ["Verde", "#10B981", "Positivo, vigente, aprobado"],
        ["Rojo", "#EF4444", "Alerta, vencido, rechazado"],
        ["Amarillo", "#F59E0B", "Precaucion, por vencer, en proceso"],
        ["Gris claro", "#F8FAFC", "Fondo del dashboard"],
        ["Blanco", "#FFFFFF", "Fondo de cuadros de KPI"],
    ]
    gen.add_table(tabla_colores, col_widths=[70, 80, 300])

    gen.add_spacer(0.1)

    gen.add_subsection("Tipos de grafico recomendados")
    gen.add_bullet("Barras agrupadas: comparar categorias (ej. sueldo por puesto)")
    gen.add_bullet("Lineas: mostrar tendencias en el tiempo (ej. nomina mensual)")
    gen.add_bullet("Dona/Pie: proporcion de un total (ej. % ISR vs % neto)")
    gen.add_bullet("EVITAR: graficos 3D, demasiados colores, ejes innecesarios")

    gen.add_spacer(0.1)

    gen.add_subsection("Limpieza visual")
    gen.add_bullet("Oculta lineas de cuadricula: Vista > desmarcar 'Lineas de cuadricula'")
    gen.add_bullet("Usa bordes solo donde agreguen claridad (no en todas las celdas)")
    gen.add_bullet("Alinea los elementos visual y consistentemente")
    gen.add_bullet("Usa una sola familia tipografica (Calibri recomendado)")

    gen.add_page_break()

    # ==== Seccion 2: Slicers Paso a Paso ===================================
    gen.add_section("2. Segmentadores de Datos (Slicers) -- Paso a Paso")

    gen.add_text(
        "Los segmentadores son filtros visuales para Tablas Dinamicas. "
        "Permiten crear dashboards interactivos sin macros ni VBA."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 1: Preparar tus datos")
    gen.add_bullet("Tus datos deben estar en formato de Tabla (Ctrl+T) o Tabla Dinamica")
    gen.add_bullet("Cada columna debe tener un encabezado unico y descriptivo")
    gen.add_bullet("No deben haber filas ni columnas vacias dentro de los datos")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 2: Crear la Tabla Dinamica")
    gen.add_bullet("Selecciona cualquier celda de tu tabla de datos")
    gen.add_bullet("Menu: Insertar > Tabla Dinamica > Hoja nueva o existente")
    gen.add_bullet("Arrastra campos: filas (Periodo), valores (Suma de Sueldo, Suma de ISR)")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 3: Insertar el segmentador")
    gen.add_bullet("Haz clic en cualquier celda de tu Tabla Dinamica")
    gen.add_bullet("Menu: Insertar > Segmentacion de datos")
    gen.add_bullet("Selecciona los campos que quieres filtrar: Periodo, Puesto, Empleado")
    gen.add_bullet("Haz clic en Aceptar -- aparecen los botones de filtro")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 4: Vincular a multiples TDs")
    gen.add_bullet("Clic derecho sobre el segmentador > Conexiones de informe...")
    gen.add_bullet("Marca TODAS las Tablas Dinamicas que deben responder al filtro")
    gen.add_bullet("Ahora un clic filtra todas las TDs y sus graficos asociados")
    gen.add_spacer(0.1)

    gen.add_subsection("Paso 5: Personalizar el segmentador")
    gen.add_bullet("Selecciona el slicer > pestana Segmentacion de datos (cinta)")
    gen.add_bullet("Cambia el estilo visual, numero de columnas y tamano")
    gen.add_bullet("Consejo: usa 2-3 columnas para que ocupe menos espacio horizontal")

    gen.add_spacer(0.1)

    gen.add_subsection("Atajos utiles")
    atajo_data = [
        ["Accion", "Atajo / Metodo"],
        ["Seleccionar multiples items", "Ctrl + clic en cada boton"],
        ["Limpiar filtro del slicer", "Icono de embudo con X (esquina del slicer)"],
        ["Mover slicer", "Arrastrar con el mouse"],
        ["Redimensionar", "Arrastrar las esquinas"],
        ["Eliminar slicer", "Seleccionar + tecla Suprimir"],
    ]
    gen.add_table(atajo_data, col_widths=[200, 250])

    gen.add_page_break()

    # ==== Seccion 3: Proteccion de Celdas ==================================
    gen.add_section("3. Proteccion de Celdas -- Guia Rapida")

    gen.add_text(
        "Proteger tu archivo asegura que el usuario final interactue correctamente "
        "con el dashboard sin romper formulas ni la estructura."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Flujo de proteccion en 4 pasos")

    pasos = [
        ["Paso", "Accion", "Donde"],
        ["1", "Selecciona las celdas de INPUT (donde el usuario escribe)", "Celdas amarillas"],
        ["2", "Clic derecho > Formato > Proteccion > desmarcar 'Bloqueada'", "Cada celda de input"],
        ["3", "Revisar > Proteger hoja > contrasena > permitir filtros/TDs", "Cada hoja"],
        ["4", "Revisar > Proteger libro > contrasena (estructura)", "Una vez por archivo"],
    ]
    gen.add_table(pasos, col_widths=[35, 280, 135])

    gen.add_spacer(0.1)

    gen.add_subsection("Que proteger y que no")
    proteccion = [
        ["Elemento", "Bloquear?", "Razon"],
        ["Celdas con formulas", "SI", "Evitar que borren o modifiquen calculos"],
        ["Celdas de input del usuario", "NO", "El usuario necesita ingresar datos"],
        ["Encabezados de tabla", "SI", "Mantener estructura"],
        ["Celdas de KPI con SUBTOTAL", "SI", "Son formulas criticas"],
        ["Nombre de hojas", "SI (libro)", "Evitar renombrar o eliminar hojas"],
    ]
    gen.add_table(proteccion, col_widths=[150, 60, 240])

    gen.add_spacer(0.1)

    gen.add_subsection("Niveles de seguridad en Excel")
    niveles = [
        ["Nivel", "Metodo", "Seguridad"],
        ["Basico", "Proteger hoja (sin contrasena)", "Minima -- cualquiera desprotege"],
        ["Medio", "Proteger hoja + libro (con contrasena)", "Moderada -- herramientas para romper"],
        ["Alto", "Contrasena de apertura de archivo (AES)", "Alta -- cifrado real"],
        ["Maximo", "Exportar como PDF (no editable)", "Maxima -- sin datos editables"],
    ]
    gen.add_table(niveles, col_widths=[55, 220, 175])

    gen.add_page_break()

    # ==== Seccion 4: Checklist de Distribucion =============================
    gen.add_section("4. Checklist de Distribucion Profesional")

    gen.add_text(
        "Usa esta lista antes de enviar cualquier archivo a clientes, "
        "jefes o colegas."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Antes de compartir")
    gen.add_bullet("[ ] Las formulas calculan correctamente con datos de prueba")
    gen.add_bullet("[ ] Los graficos se actualizan al cambiar filtros/slicers")
    gen.add_bullet("[ ] Las celdas de input estan desbloqueadas y resaltadas en amarillo")
    gen.add_bullet("[ ] Las celdas de formula estan bloqueadas")
    gen.add_bullet("[ ] Las hojas estan protegidas con contrasena")
    gen.add_bullet("[ ] La estructura del libro esta protegida")
    gen.add_bullet("[ ] No hay datos personales o confidenciales expuestos")
    gen.add_bullet("[ ] El nombre del archivo sigue la convencion profesional")
    gen.add_spacer(0.1)

    gen.add_subsection("Formato del nombre de archivo")
    gen.add_code("[Empresa]_[TipoDocumento]_[Periodo]_[Version].[ext]")
    gen.add_spacer(0.05)
    gen.add_bullet("Ejemplo: GrupoTorres_Nomina_2026_Enero_v1.xlsx")
    gen.add_bullet("Ejemplo: CNO850315_ISR_Anual_2025_Final.pdf")
    gen.add_bullet("Evita espacios, caracteres especiales y acentos en nombres de archivo")
    gen.add_spacer(0.1)

    gen.add_subsection("Que formato usar")
    formato = [
        ["Situacion", "Formato", "Razon"],
        ["Reporte final a cliente", "PDF", "No editable, aspecto profesional"],
        ["Declaracion o complemento SAT", "PDF + XML", "Formato oficial"],
        ["Plantilla que el cliente debe llenar", "Excel protegido", "Mantiene formulas"],
        ["Dashboard interactivo para equipo", "Excel protegido", "Slicers y TDs activos"],
        ["Respaldo/archivo maestro", "Excel sin proteger", "Facilita edicion futura"],
    ]
    gen.add_table(formato, col_widths=[160, 90, 200])

    gen.add_spacer(0.1)

    gen.add_subsection("Despues de compartir")
    gen.add_bullet("[ ] Confirma que el destinatario puede abrir el archivo")
    gen.add_bullet("[ ] Verifica que la version de Excel del destinatario es compatible")
    gen.add_bullet("[ ] Guarda tu copia maestra sin proteger (sufijo _master.xlsx)")
    gen.add_bullet("[ ] Documenta la contrasena de proteccion en lugar seguro")

    gen.add_spacer(0.3)

    gen.add_subsection("Archivos de este modulo")
    gen.add_bullet("09_Layout_Dashboard_Contable.xlsx -- Template de layout para dashboard")
    gen.add_bullet("10_Dashboard_Final_Integrado.xlsx -- Datos + Calculadora + Dashboard")
    gen.add_bullet("11_Guia_Proteccion_y_Seguridad.pdf -- Guia detallada de proteccion")
    gen.add_bullet("Referencia_Modulo_4.pdf -- Este documento")

    gen.save()


if __name__ == "__main__":
    build()
