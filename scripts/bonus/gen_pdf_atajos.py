"""
Generador: Atajos_Excel_CheatSheet.md
Bonus — Cheat sheet de atajos de teclado de Excel

Categorias:
  1. Navegacion
  2. Seleccion
  3. Edicion
  4. Formato
  5. Formulas
  6. Tablas y Datos
  7. Tablas Dinamicas
  8. Graficos
  9. Atajos de Productividad
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK
from scripts.generators.md_gen import MarkdownGenerator

OUTPUT_DIR = PACK / "Bonus"


def _shortcut_table(md, title, shortcuts):
    """Add a category section with a table of shortcuts."""
    md.add_section(title)

    table_data = [["Atajo", "Accion", "Tip"]]
    for atajo, accion, tip in shortcuts:
        table_data.append([f"`{atajo}`", accion, f"*{tip}*" if tip else ""])

    md.add_table(table_data)


def build():
    pdf = MarkdownGenerator(
        filename="Atajos_Excel_CheatSheet.md",
        output_dir=OUTPUT_DIR,
        title="Atajos de Teclado Excel - Cheat Sheet",
    )

    # ── Portada ───────────────────────────────────────────────────
    pdf.add_cover(
        title="Atajos de Teclado Excel",
        subtitle="Referencia rapida organizada por categoria",
        modulo="Cheat Sheet - Bonus del Curso",
    )

    # ── 1. Navegacion ─────────────────────────────────────────────
    _shortcut_table(pdf, "1. Navegacion", [
        ("Ctrl + Home", "Ir a celda A1", "Inicio del libro"),
        ("Ctrl + End", "Ir a ultima celda con datos", "Esquina inferior derecha usada"),
        ("Ctrl + Flecha", "Saltar al borde del rango", "Funciona en las 4 direcciones"),
        ("Ctrl + *", "Seleccionar region actual", "Equivale a Ctrl+Shift+8"),
        ("Ctrl + G (o F5)", "Ir a... (cuadro de dialogo)", "Util para ir a celdas especificas"),
        ("Ctrl + Page Up", "Hoja anterior", "Navegar entre hojas rapidamente"),
        ("Ctrl + Page Down", "Hoja siguiente", "Navegar entre hojas rapidamente"),
        ("Alt + Page Up", "Pantalla a la izquierda", "Desplazamiento horizontal"),
        ("Alt + Page Down", "Pantalla a la derecha", "Desplazamiento horizontal"),
        ("Ctrl + Tab", "Siguiente libro abierto", "Cambiar entre archivos Excel"),
    ])

    # ── 2. Seleccion ─────────────────────────────────────────────
    _shortcut_table(pdf, "2. Seleccion", [
        ("Shift + Clic", "Seleccionar rango desde celda activa", "Mas rapido que arrastrar"),
        ("Ctrl + Shift + End", "Seleccionar hasta ultima celda usada", "Ideal para rangos grandes"),
        ("Ctrl + Shift + Home", "Seleccionar desde activa hasta A1", "Selecciona todo arriba"),
        ("Ctrl + Shift + Flecha", "Extender seleccion hasta borde", "Combina salto + seleccion"),
        ("Ctrl + Space", "Seleccionar columna completa", "Toda la columna de la celda activa"),
        ("Shift + Space", "Seleccionar fila completa", "Toda la fila de la celda activa"),
        ("Ctrl + A", "Seleccionar todo (tabla o hoja)", "1er clic = tabla, 2do = toda la hoja"),
        ("Ctrl + Shift + *", "Seleccionar region de datos actual", "Detecta automaticamente el rango"),
        ("Alt + ;", "Seleccionar solo celdas visibles", "Ignora filas ocultas por filtro"),
    ])

    # ── 3. Edicion ────────────────────────────────────────────────
    _shortcut_table(pdf, "3. Edicion", [
        ("F2", "Editar celda activa", "Entra en modo edicion sin borrar"),
        ("Ctrl + Z", "Deshacer ultima accion", "Hasta 100 niveles de deshacer"),
        ("Ctrl + Y", "Rehacer / Repetir ultima accion", "Tambien funciona como repetir"),
        ("Ctrl + D", "Copiar celda de arriba hacia abajo", "Rellena con contenido de arriba"),
        ("Ctrl + R", "Copiar celda de izquierda a derecha", "Rellena con contenido de la izq."),
        ("Ctrl + J", "Buscar salto de linea (en Buscar y Reemplazar)", "En el cuadro Ctrl+H, busca saltos de linea"),
        ("Delete", "Borrar contenido de celda(s)", "Solo contenido, no formato"),
        ("Ctrl + - (menos)", "Eliminar celda/fila/columna", "Muestra opciones de eliminacion"),
        ("Ctrl + + (mas)", "Insertar celda/fila/columna", "Muestra opciones de insercion"),
        ("Ctrl + H", "Buscar y reemplazar", "Reemplazo masivo de datos"),
        ("Ctrl + F", "Buscar", "Buscar texto o valores"),
        ("F3", "Pegar nombre definido", "Inserta nombres de rangos"),
    ])

    # ── 4. Formato ────────────────────────────────────────────────
    _shortcut_table(pdf, "4. Formato", [
        ("Ctrl + 1", "Formato de celdas (dialogo completo)", "Acceso a TODAS las opciones"),
        ("Ctrl + B (o Ctrl + N)", "Negrita", "Toggle on/off"),
        ("Ctrl + I (o Ctrl + K)", "Cursiva", "Toggle on/off"),
        ("Ctrl + U (o Ctrl + S)", "Subrayado", "Toggle on/off"),
        ("Ctrl + Shift + $", "Formato moneda", "Aplica formato $#,##0.00"),
        ("Ctrl + Shift + %", "Formato porcentaje", "Multiplica por 100 y agrega %"),
        ("Ctrl + Shift + #", "Formato fecha", "Formato DD-MMM-AA"),
        ("Ctrl + Shift + @", "Formato hora", "Formato HH:MM AM/PM"),
        ("Ctrl + Shift + !", "Formato numero con miles", "Separador de miles y 2 decimales"),
        ("Ctrl + Shift + ~", "Formato general", "Quita formato numerico especial"),
        ("Alt + H, O, I", "Autoajustar ancho de columna", "Ruta de cinta rapida"),
        ("Alt + H, O, A", "Autoajustar alto de fila", "Ruta de cinta rapida"),
    ])

    pdf.add_page_break()

    # ── 5. Formulas ───────────────────────────────────────────────
    _shortcut_table(pdf, "5. Formulas", [
        ("F4", "Alternar referencia absoluta/relativa", "$A$1 -> A$1 -> $A1 -> A1"),
        ("Tab", "Autocompletar funcion sugerida", "Acepta la sugerencia de IntelliSense"),
        ("Ctrl + `", "Mostrar/ocultar formulas en celdas", "Ver todas las formulas de la hoja"),
        ("Alt + =", "Autosuma rapida", "Inserta =SUMA() automaticamente"),
        ("Ctrl + Shift + U", "Expandir barra de formulas", "Ver formula completa si es larga"),
        ("F9", "Evaluar parte de formula", "Seleccionar parte y F9 para ver resultado"),
        ("Ctrl + '", "Copiar formula de celda superior", "Copia formula sin ajustar"),
        ("Ctrl + Shift + Enter", "Formula matricial (legacy)", "Para versiones pre-365"),
        ("Ctrl + Shift + A", "Insertar argumentos de funcion", "Muestra nombres de argumentos"),
        ("F4 (fuera de edicion)", "Repetir ultima accion", "Repite formato, insercion, etc."),
    ])

    # ── 6. Tablas y Datos ─────────────────────────────────────────
    _shortcut_table(pdf, "6. Tablas y Datos", [
        ("Ctrl + T", "Crear tabla desde rango", "Detecta rango automaticamente"),
        ("Alt + Flecha Abajo", "Abrir filtro de columna", "Dentro de tabla o con filtro activo"),
        ("Ctrl + Shift + L", "Activar/desactivar filtros", "Toggle filtros en rango"),
        ("Alt + D, S", "Ordenar (dialogo completo)", "Multiples niveles de orden"),
        ("Ctrl + Shift + F3", "Crear nombres desde seleccion", "Nombra rangos automaticamente"),
        ("Ctrl + T, luego Tab", "Tab para moverse en tabla", "Navega celda por celda en tabla"),
        ("Alt + A, R, A", "Quitar duplicados", "Ruta de cinta: Datos > Quitar dup."),
        ("Alt + A, V, V", "Validacion de datos", "Ruta de cinta: Datos > Validacion"),
        ("Ctrl + Shift + &", "Aplicar bordes al rango", "Borde exterior al rango seleccionado"),
    ])

    # ── 7. Tablas Dinamicas ───────────────────────────────────────
    _shortcut_table(pdf, "7. Tablas Dinamicas", [
        ("Alt + N, V", "Insertar tabla dinamica", "Ruta de cinta: Insertar > TD"),
        ("Clic derecho > Opciones", "Acceder a opciones de TD", "Configuracion detallada de la TD"),
        ("Doble clic en valor", "Drill-down (ver detalle)", "Crea hoja con datos de esa celda"),
        ("Alt + Shift + Flecha Der", "Agrupar seleccion", "Agrupa filas/columnas seleccionadas"),
        ("Alt + Shift + Flecha Izq", "Desagrupar seleccion", "Desagrupa filas/columnas"),
        ("Clic derecho > Actualizar", "Actualizar tabla dinamica", "Refresca datos de la TD"),
        ("Alt + F5", "Actualizar todas las TDs", "Refresca todas las conexiones"),
        ("Clic derecho > Formato", "Formato de numero en TD", "Formato para campo de valor"),
    ])

    # ── 8. Graficos ───────────────────────────────────────────────
    _shortcut_table(pdf, "8. Graficos", [
        ("Alt + F1", "Crear grafico en hoja actual", "Grafico incrustado instantaneo"),
        ("F11", "Crear grafico en hoja nueva", "Hoja de grafico dedicada"),
        ("Ctrl + clic en elemento", "Seleccionar elemento del grafico", "Series, ejes, leyenda"),
        ("Delete (en grafico)", "Eliminar elemento seleccionado", "Quita serie o elemento"),
        ("Ctrl + 1 (en grafico)", "Formato del elemento seleccionado", "Panel de formato detallado"),
        ("Flecha Arriba/Abajo", "Navegar entre series de datos", "Dentro del grafico"),
    ])

    # ── 9. Productividad ─────────────────────────────────────────
    _shortcut_table(pdf, "9. Atajos de Productividad", [
        ("F4", "Repetir ultima accion", "Funciona para formato, borrado, etc."),
        ("Ctrl + ;", "Insertar fecha actual", "Fecha estatica (no cambia)"),
        ("Ctrl + Shift + :", "Insertar hora actual", "Hora estatica (no cambia)"),
        ("Ctrl + Shift + +", "Insertar fila/columna", "Segun seleccion previa"),
        ("Alt + Enter", "Salto de linea en celda", "Multiples lineas en una celda"),
        ("Ctrl + E", "Relleno rapido (Flash Fill)", "Detecta patrones automaticamente"),
        ("Ctrl + Shift + V", "Pegado especial (menu)", "Elige que pegar: valores, formatos..."),
        ("Alt + E, S, V", "Pegar solo valores", "Ruta clasica de pegado especial"),
        ("Ctrl + P", "Imprimir / Vista previa", "Acceso rapido a impresion"),
        ("Ctrl + W", "Cerrar libro actual", "No cierra Excel, solo el libro"),
        ("F12", "Guardar como", "Dialogo completo de guardar"),
        ("Ctrl + U", "Nuevo libro en blanco", "Ctrl+N en Excel en ingles"),
    ])

    # ── Nota final ────────────────────────────────────────────────
    pdf.add_spacer(0.3)
    pdf.add_text(
        "**Nota:** Algunos atajos pueden variar segun la version de Excel "
        "(2019, 2021, 365) y la configuracion de idioma. Los atajos mostrados "
        "corresponden a la version en espanol de Windows. En Mac, sustituir "
        "Ctrl por Cmd en la mayoria de los casos."
    )
    pdf.add_spacer(0.1)
    pdf.add_text(
        "**Tip final:** No intenten memorizar todos los atajos de golpe. "
        "Elijan 3-5 que usen frecuentemente, practiquenlos una semana, y luego "
        "agreguen 3-5 mas. En un mes habran duplicado su velocidad en Excel."
    )

    pdf.save()


if __name__ == "__main__":
    build()
