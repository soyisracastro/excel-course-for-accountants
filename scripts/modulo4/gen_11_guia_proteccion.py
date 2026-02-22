"""
Generador: 11_Guia_Proteccion_y_Seguridad.pdf
Modulo 4 -- El Dashboard Inteligente y Entrega Profesional

Guia PDF sobre proteccion y seguridad en Excel:
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
        title="Guia de Proteccion y Seguridad en Excel"
    )

    # ---- Portada ----
    gen.add_cover(
        title="Guia de Proteccion y Seguridad en Excel",
        modulo="Modulo 4 -- El Dashboard Inteligente y Entrega Profesional",
        subtitle="Protege, comparte y entrega archivos profesionales"
    )

    # ==== Seccion 1: Proteccion de Celdas ==================================
    gen.add_section("1. Proteccion de Celdas")

    gen.add_text(
        "En Excel, TODAS las celdas vienen bloqueadas por defecto, pero el bloqueo "
        "solo se activa cuando proteges la hoja. La estrategia correcta es:"
    )
    gen.add_spacer(0.1)

    gen.add_subsection("1.1 Desbloquear celdas de entrada (inputs)")
    gen.add_bullet("Selecciona las celdas donde el usuario debe escribir datos (ej. B4 en Calculadora)")
    gen.add_bullet("Clic derecho > Formato de celdas > pestana Proteccion")
    gen.add_bullet("Desmarca la casilla 'Bloqueada'")
    gen.add_bullet("Opcionalmente marca 'Oculta' para ocultar formulas en la barra de formulas")
    gen.add_spacer(0.1)

    gen.add_subsection("1.2 Mantener bloqueadas las celdas de formula")
    gen.add_bullet("Las celdas con formulas deben mantener la casilla 'Bloqueada' activada")
    gen.add_bullet("Esto evita que el usuario modifique o borre accidentalmente las formulas")
    gen.add_bullet("Ejemplo: celdas B5:B10 en la Calculadora ISR contienen BUSCARV")
    gen.add_spacer(0.1)

    gen.add_subsection("1.3 Identificar visualmente las celdas editables")
    gen.add_bullet("Usa un color de fondo distinto para celdas de entrada (ej. amarillo #F59E0B)")
    gen.add_bullet("Agrega una nota o etiqueta: 'Ingresa tu dato aqui'")
    gen.add_bullet("Esto le indica al usuario donde puede escribir sin confusion")

    gen.add_spacer(0.2)

    # ==== Seccion 2: Proteccion de Hojas ==================================
    gen.add_section("2. Proteccion de Hojas")

    gen.add_text(
        "Proteger una hoja impide que los usuarios modifiquen celdas bloqueadas, "
        "pero permite la edicion en celdas desbloqueadas."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("Pasos para proteger una hoja")
    gen.add_bullet("Menu: Revisar > Proteger hoja")
    gen.add_bullet("Establece una contrasena (opcional pero recomendado)")
    gen.add_bullet("Selecciona los permisos que deseas otorgar:")
    gen.add_spacer(0.05)

    permisos = [
        ["Permiso", "Descripcion", "Recomendado"],
        ["Seleccionar celdas bloqueadas", "El usuario puede ver pero no editar", "Si"],
        ["Seleccionar celdas desbloqueadas", "El usuario puede editar celdas de input", "Si"],
        ["Formato de celdas", "Permite cambiar formato visual", "No"],
        ["Insertar filas", "Permite agregar filas nuevas", "Segun caso"],
        ["Eliminar filas", "Permite borrar filas", "No"],
        ["Ordenar", "Permite ordenar datos", "Si"],
        ["Usar Autofiltro", "Permite filtrar datos", "Si"],
        ["Usar tablas dinamicas", "Permite interactuar con TDs", "Si"],
    ]
    gen.add_table(permisos, col_widths=[150, 220, 80])

    gen.add_spacer(0.1)
    gen.add_subsection("Contrasenas seguras")
    gen.add_bullet("Usa al menos 8 caracteres con mayusculas, numeros y simbolos")
    gen.add_bullet("IMPORTANTE: La proteccion de hoja NO es cifrado fuerte")
    gen.add_bullet("No confies en ella para datos altamente confidenciales")
    gen.add_bullet("Para datos sensibles, usa contrasena de apertura de archivo (ver seccion 3)")

    gen.add_page_break()

    # ==== Seccion 3: Proteccion de Libro ==================================
    gen.add_section("3. Proteccion de Libro")

    gen.add_text(
        "La proteccion de libro evita cambios estructurales: agregar, eliminar, "
        "renombrar u ocultar hojas."
    )
    gen.add_spacer(0.1)

    gen.add_subsection("3.1 Proteger estructura del libro")
    gen.add_bullet("Menu: Revisar > Proteger libro")
    gen.add_bullet("Marca 'Estructura' para evitar agregar/eliminar hojas")
    gen.add_bullet("Establece contrasena")
    gen.add_spacer(0.1)

    gen.add_subsection("3.2 Contrasena de apertura de archivo")
    gen.add_bullet("Menu: Archivo > Guardar como > Herramientas > Opciones generales")
    gen.add_bullet("'Contrasena de apertura': el archivo no se abre sin ella (cifrado)")
    gen.add_bullet("'Contrasena de escritura': permite abrir en solo lectura sin contrasena")
    gen.add_bullet("NOTA: La contrasena de apertura usa cifrado AES -- es segura")
    gen.add_spacer(0.1)

    gen.add_subsection("3.3 Marcar como final")
    gen.add_bullet("Menu: Archivo > Informacion > Proteger libro > Marcar como final")
    gen.add_bullet("Pone el archivo en modo solo lectura (el usuario puede desactivarlo)")
    gen.add_bullet("Es una recomendacion, no una proteccion fuerte")

    gen.add_spacer(0.2)

    # ==== Seccion 4: Distribucion ==========================================
    gen.add_section("4. Distribucion: PDF vs Excel Protegido")

    gen.add_text(
        "La forma de compartir depende de lo que necesitas que haga el destinatario."
    )
    gen.add_spacer(0.1)

    comparacion = [
        ["Criterio", "PDF", "Excel Protegido"],
        ["El usuario necesita editar", "No", "Si (celdas de input)"],
        ["Mantiene formulas activas", "No", "Si"],
        ["Interactuar con slicers/TDs", "No", "Si"],
        ["Seguridad del contenido", "Alta (no editable)", "Media (contrasena)"],
        ["Tamano de archivo", "Ligero", "Normal"],
        ["Requiere Excel instalado", "No", "Si"],
        ["Ideal para", "Reportes finales, firmas", "Plantillas, calculadoras"],
    ]
    gen.add_table(comparacion, col_widths=[130, 150, 170])

    gen.add_spacer(0.1)

    gen.add_subsection("Como guardar como PDF desde Excel")
    gen.add_bullet("Menu: Archivo > Guardar como > Tipo: PDF")
    gen.add_bullet("O bien: Archivo > Exportar > Crear documento PDF/XPS")
    gen.add_bullet("Selecciona las hojas a incluir antes de guardar")
    gen.add_bullet("Tip: Configura el area de impresion antes (Diseno de pagina > Area de impresion)")

    gen.add_page_break()

    # ==== Seccion 5: Convenciones de Nombres ===============================
    gen.add_section("5. Convenciones de Nombres Profesionales")

    gen.add_text(
        "Un nombre de archivo profesional facilita la organizacion, la busqueda "
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
    gen.add_bullet("NO uses espacios -- usa guion bajo (_) o guion medio (-)")
    gen.add_bullet("Incluye el periodo (mes, trimestre, anio)")
    gen.add_bullet("Incluye version (v1, v2, Final, Borrador)")
    gen.add_bullet("Usa el RFC cuando el archivo es para un cliente especifico")
    gen.add_bullet("Manten nombres cortos pero descriptivos (max ~50 caracteres)")

    gen.add_spacer(0.2)

    # ==== Seccion 6: Checklist de Entrega ==================================
    gen.add_section("6. Checklist de Entrega Profesional")

    gen.add_text(
        "Antes de enviar un archivo a tu cliente, jefe o colega, "
        "verifica cada punto de esta lista."
    )
    gen.add_spacer(0.1)

    checklist = [
        ["#", "Verificacion", "Hecho"],
        ["1", "Las formulas calculan correctamente (verifica con datos de prueba)", "[ ]"],
        ["2", "Las celdas de input estan DESBLOQUEADAS y resaltadas en amarillo", "[ ]"],
        ["3", "Las celdas de formula estan BLOQUEADAS", "[ ]"],
        ["4", "La hoja esta protegida con contrasena", "[ ]"],
        ["5", "La estructura del libro esta protegida", "[ ]"],
        ["6", "No hay datos personales o confidenciales expuestos", "[ ]"],
        ["7", "Los graficos se ven correctamente al cambiar filtros", "[ ]"],
        ["8", "Los segmentadores (slicers) estan conectados a todas las TDs", "[ ]"],
        ["9", "El nombre del archivo sigue la convencion profesional", "[ ]"],
        ["10", "Se genero una copia PDF del reporte final", "[ ]"],
        ["11", "Se verifico en otra computadora o version de Excel", "[ ]"],
        ["12", "Se incluyen instrucciones de uso (hoja o documento adjunto)", "[ ]"],
    ]
    gen.add_table(checklist, col_widths=[25, 380, 40])

    gen.add_spacer(0.3)
    gen.add_text(
        "CONSEJO: Guarda una copia sin proteger como archivo maestro (ej. '_master.xlsx') "
        "y distribuye siempre la version protegida."
    )

    gen.save()


if __name__ == "__main__":
    build()
