"""
Generador: 11_Guia_Proteccion_y_Seguridad.md
Modulo 4 -- El Dashboard Inteligente y Entrega Profesional

Guia sobre proteccion y seguridad en Excel:
  - Proteccion de celdas (bloquear formulas, desbloquear inputs)
  - Proteccion de hojas
  - Proteccion de libro
  - Distribucion: PDF vs Excel protegido
  - Convenciones de nombres profesionales
  - Checklist de entrega
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK
from scripts.generators.md_gen import MarkdownGenerator

OUTPUT_DIR = PACK / "Modulo_4_Dashboard"


def build():
    gen = MarkdownGenerator(
        "11_Guia_Proteccion_y_Seguridad.md",
        OUTPUT_DIR,
        title="Guía de Protección y Seguridad en Excel"
    )

    # ---- Portada ----
    gen.add_cover(
        title="Guía de Protección y Seguridad en Excel",
        modulo="Módulo 4 — El Dashboard Inteligente y Entrega Profesional",
        subtitle="Protege, comparte y entrega archivos profesionales"
    )

    # ==== Sección 1: Protección de Celdas ==================================
    gen.add_section("1. Protección de Celdas")

    gen.add_text(
        "En Excel, TODAS las celdas vienen bloqueadas por defecto, pero el bloqueo "
        "solo se activa cuando proteges la hoja. La estrategia correcta es:"
    )
    gen.add_spacer(0.1)

    gen.add_subsection("1.1 Desbloquear celdas de entrada (inputs)")
    gen.add_bullet("Selecciona las celdas donde el usuario debe escribir datos (ej. B4 en Calculadora)")
    gen.add_bullet("Clic derecho > Formato de celdas > pestaña Protección")
    gen.add_bullet("Desmarca la casilla 'Bloqueada'")
    gen.add_bullet("Opcionalmente marca 'Oculta' para ocultar fórmulas en la barra de fórmulas")
    gen.add_spacer(0.1)

    gen.add_subsection("1.2 Mantener bloqueadas las celdas de fórmula")
    gen.add_bullet("Las celdas con fórmulas deben mantener la casilla 'Bloqueada' activada")
    gen.add_bullet("Esto evita que el usuario modifique o borre accidentalmente las fórmulas")
    gen.add_bullet("Ejemplo: celdas B5:B10 en la Calculadora ISR contienen BUSCARV")
    gen.add_spacer(0.1)

    gen.add_subsection("1.3 Identificar visualmente las celdas editables")
    gen.add_bullet("Usa un color de fondo distinto para celdas de entrada (ej. amarillo #F59E0B)")
    gen.add_bullet("Agrega una nota o etiqueta: 'Ingresa tu dato aquí'")
    gen.add_bullet("Esto le indica al usuario dónde puede escribir sin confusión")

    gen.add_spacer(0.2)

    # ==== Sección 2: Protección de Hojas ==================================
    gen.add_section("2. Protección de Hojas")

    gen.add_text(
        "Proteger una hoja impide que los usuarios modifiquen celdas bloqueadas, "
        "pero permite la edición en celdas desbloqueadas."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Pasos para proteger una hoja")
    gen.add_bullet("Menú: Revisar > Proteger hoja")
    gen.add_bullet("Establece una contraseña (opcional pero recomendado)")
    gen.add_bullet("Selecciona los permisos que deseas otorgar:")
    gen.add_spacer(0.05)

    permisos = [
        ["Permiso", "Descripción", "Recomendado"],
        ["Seleccionar celdas bloqueadas", "El usuario puede ver pero no editar", "Sí"],
        ["Seleccionar celdas desbloqueadas", "El usuario puede editar celdas de input", "Sí"],
        ["Formato de celdas", "Permite cambiar formato visual", "No"],
        ["Insertar filas", "Permite agregar filas nuevas", "Según caso"],
        ["Eliminar filas", "Permite borrar filas", "No"],
        ["Ordenar", "Permite ordenar datos", "Sí"],
        ["Usar Autofiltro", "Permite filtrar datos", "Sí"],
        ["Usar tablas dinámicas", "Permite interactuar con TDs", "Sí"],
    ]
    gen.add_table(permisos, col_widths=[150, 220, 80])

    gen.add_spacer(0.1)
    gen.add_subsection("Contraseñas seguras")
    gen.add_bullet("Usa al menos 8 caracteres con mayúsculas, números y símbolos")
    gen.add_bullet("IMPORTANTE: La protección de hoja NO es cifrado fuerte")
    gen.add_bullet("No confíes en ella para datos altamente confidenciales")
    gen.add_bullet("Para datos sensibles, usa contraseña de apertura de archivo (ver sección 3)")

    gen.add_page_break()

    # ==== Sección 3: Protección de Libro ==================================
    gen.add_section("3. Protección de Libro")

    gen.add_text(
        "La protección de libro evita cambios estructurales: agregar, eliminar, "
        "renombrar u ocultar hojas."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("3.1 Proteger estructura del libro")
    gen.add_bullet("Menú: Revisar > Proteger libro")
    gen.add_bullet("Marca 'Estructura' para evitar agregar/eliminar hojas")
    gen.add_bullet("Establece contraseña")
    gen.add_spacer(0.1)

    gen.add_subsection("3.2 Contraseña de apertura de archivo")
    gen.add_bullet("Menú: Archivo > Guardar como > Herramientas > Opciones generales")
    gen.add_bullet("'Contraseña de apertura': el archivo no se abre sin ella (cifrado)")
    gen.add_bullet("'Contraseña de escritura': permite abrir en solo lectura sin contraseña")
    gen.add_bullet("NOTA: La contraseña de apertura usa cifrado AES — es segura")
    gen.add_spacer(0.1)

    gen.add_subsection("3.3 Marcar como final")
    gen.add_bullet("Menú: Archivo > Información > Proteger libro > Marcar como final")
    gen.add_bullet("Pone el archivo en modo solo lectura (el usuario puede desactivarlo)")
    gen.add_bullet("Es una recomendación, no una protección fuerte")

    gen.add_spacer(0.2)

    # ==== Sección 4: Distribución ==========================================
    gen.add_section("4. Distribución: PDF vs Excel Protegido")

    gen.add_text(
        "La forma de compartir depende de lo que necesitas que haga el destinatario."
    )
    gen.add_spacer(0.1)

    comparacion = [
        ["Criterio", "PDF", "Excel Protegido"],
        ["El usuario necesita editar", "No", "Sí (celdas de input)"],
        ["Mantiene fórmulas activas", "No", "Sí"],
        ["Interactuar con slicers/TDs", "No", "Sí"],
        ["Seguridad del contenido", "Alta (no editable)", "Media (contraseña)"],
        ["Tamaño de archivo", "Ligero", "Normal"],
        ["Requiere Excel instalado", "No", "Sí"],
        ["Ideal para", "Reportes finales, firmas", "Plantillas, calculadoras"],
    ]
    gen.add_table(comparacion, col_widths=[130, 150, 170])

    gen.add_spacer(0.1)

    gen.add_subsection("Cómo guardar como PDF desde Excel")
    gen.add_bullet("Menú: Archivo > Guardar como > Tipo: PDF")
    gen.add_bullet("O bien: Archivo > Exportar > Crear documento PDF/XPS")
    gen.add_bullet("Selecciona las hojas a incluir antes de guardar")
    gen.add_bullet("Tip: Configura el área de impresión antes (Diseño de página > Área de impresión)")

    gen.add_page_break()

    # ==== Sección 5: Convenciones de Nombres ===============================
    gen.add_section("5. Convenciones de Nombres Profesionales")

    gen.add_text(
        "Un nombre de archivo profesional facilita la organización, la búsqueda "
        "y transmite seriedad al cliente."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Formato recomendado")
    gen.add_code("[Empresa]_[Tipo]_[Periodo]_[Version].[ext]")
    gen.add_spacer(0.05)

    gen.add_subsection("Ejemplos")
    gen.add_bullet("GrupoTorres_Nomina_2026_Enero_v1.xlsx")
    gen.add_bullet("CNO850315XX1_ISR_Anual_2025_Final.pdf")
    gen.add_bullet("Dashboard_Fiscal_2026_Q1_v2.xlsx")
    gen.add_bullet("Reporte_Auditoria_2025_Borrador.pdf")
    gen.add_spacer(0.1)

    gen.add_subsection("Reglas clave")
    gen.add_bullet("NO uses espacios — usa guion bajo (_) o guion medio (-)")
    gen.add_bullet("Incluye el periodo (mes, trimestre, año)")
    gen.add_bullet("Incluye versión (v1, v2, Final, Borrador)")
    gen.add_bullet("Usa el RFC cuando el archivo es para un cliente específico")
    gen.add_bullet("Mantén nombres cortos pero descriptivos (máx ~50 caracteres)")

    gen.add_spacer(0.2)

    # ==== Sección 6: Checklist de Entrega ==================================
    gen.add_section("6. Checklist de Entrega Profesional")

    gen.add_text(
        "Antes de enviar un archivo a tu cliente, jefe o colega, "
        "verifica cada punto de esta lista."
    )
    gen.add_spacer(0.1)

    checklist = [
        ["#", "Verificación", "Hecho"],
        ["1", "Las fórmulas calculan correctamente (verifica con datos de prueba)", "[ ]"],
        ["2", "Las celdas de input están DESBLOQUEADAS y resaltadas en amarillo", "[ ]"],
        ["3", "Las celdas de fórmula están BLOQUEADAS", "[ ]"],
        ["4", "La hoja está protegida con contraseña", "[ ]"],
        ["5", "La estructura del libro está protegida", "[ ]"],
        ["6", "No hay datos personales o confidenciales expuestos", "[ ]"],
        ["7", "Los gráficos se ven correctamente al cambiar filtros", "[ ]"],
        ["8", "Los segmentadores (slicers) están conectados a todas las TDs", "[ ]"],
        ["9", "El nombre del archivo sigue la convención profesional", "[ ]"],
        ["10", "Se generó una copia PDF del reporte final", "[ ]"],
        ["11", "Se verificó en otra computadora o versión de Excel", "[ ]"],
        ["12", "Se incluyen instrucciones de uso (hoja o documento adjunto)", "[ ]"],
    ]
    gen.add_table(checklist, col_widths=[25, 380, 40])

    gen.add_spacer(0.3)
    gen.add_text(
        "CONSEJO: Guarda una copia sin proteger como archivo maestro (ej. '_master.xlsx') "
        "y distribuye siempre la versión protegida."
    )

    gen.save()


if __name__ == "__main__":
    build()
