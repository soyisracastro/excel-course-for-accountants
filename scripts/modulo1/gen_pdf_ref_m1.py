"""
Generador: Referencia_Modulo_1.pdf (~6 paginas)
Modulo 1 — Logica Contable y Funciones de Control

Contenido:
  - Portada
  - Tarjetas de funciones (SUMA, PROMEDIO, TRUNCAR, SI, BUSCARV, HOY, FECHA, EXTRAE)
  - Tarifa ISR 2026 condensada (Art. 152 LISR)
  - Guia Factor de Actualizacion paso a paso
  - 5 ejercicios con respuestas (calculos ISR)
  - Atajos del modulo
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import PACK
from scripts.config.isr_2026 import TARIFA_ANUAL, calcular_isr_anual
from scripts.generators.md_gen import MarkdownGenerator

inch = 72  # compat: MarkdownGenerator ignores col_widths

OUTPUT_DIR = PACK / "Modulo_1_Funciones"


def _fmt_money(v):
    """Formatea un numero como moneda MXN."""
    if isinstance(v, str):
        return v
    return "${:,.2f}".format(v)


def _fmt_pct(v):
    """Formatea un float como porcentaje."""
    return "{:.2f}%".format(v)


def build():
    gen = MarkdownGenerator(
        filename="Referencia_Modulo_1.md",
        output_dir=OUTPUT_DIR,
        title="Referencia - Modulo 1: Logica Contable y Funciones de Control",
    )

    # ── Portada ────────────────────────────────────────────────────
    gen.add_cover(
        title="Referencia - Modulo 1",
        subtitle="Logica Contable y Funciones de Control",
        modulo="Modulo 1 de 5",
    )

    # ── Seccion 1: Tarjetas de Funciones ───────────────────────────
    gen.add_section("Tarjetas de Funciones Esenciales")
    gen.add_text(
        "Cada tarjeta resume la sintaxis, descripcion y un ejemplo contable "
        "de las funciones cubiertas en el Modulo 1."
    )
    gen.add_spacer(0.15)

    funciones = [
        {
            "nombre": "SUMA",
            "sintaxis": "=SUMA(numero1, [numero2], ...)",
            "descripcion": (
                "Suma todos los numeros en un rango o lista de argumentos. "
                "Ignora celdas con texto o vacias."
            ),
            "ejemplo": (
                "=SUMA(B2:B13) — Suma las ventas mensuales de enero a diciembre "
                "para obtener el total anual."
            ),
        },
        {
            "nombre": "PROMEDIO",
            "sintaxis": "=PROMEDIO(numero1, [numero2], ...)",
            "descripcion": (
                "Calcula la media aritmetica de los valores. "
                "Ignora celdas vacias pero NO ignora ceros."
            ),
            "ejemplo": (
                "=PROMEDIO(C2:C13) — Gasto promedio mensual de un departamento "
                "para presupuesto del siguiente ejercicio."
            ),
        },
        {
            "nombre": "TRUNCAR",
            "sintaxis": "=TRUNCAR(numero, num_decimales)",
            "descripcion": (
                "Corta un numero al numero de decimales indicado SIN redondear. "
                "Obligatorio para Factor de Actualizacion (Art. 17-A CFF, 4 decimales)."
            ),
            "ejemplo": (
                "=TRUNCAR(141.200/136.163, 4) — Calcula el Factor de Actualizacion "
                "truncado al diezmilesimo: 1.0369."
            ),
        },
        {
            "nombre": "SI",
            "sintaxis": "=SI(prueba_logica, valor_verdadero, valor_falso)",
            "descripcion": (
                "Evalua una condicion y devuelve un valor si es verdadera y otro "
                "si es falsa. Se puede anidar hasta 7 niveles (recomendado: 3)."
            ),
            "ejemplo": (
                "=SI(C2<0, \"VENCIDO\", SI(C2<=15, \"URGENTE\", \"OK\")) — "
                "Semaforo de vencimientos fiscales."
            ),
        },
        {
            "nombre": "BUSCARV",
            "sintaxis": "=BUSCARV(valor_buscado, tabla, num_columna, [ordenado])",
            "descripcion": (
                "Busca un valor en la primera columna de una tabla y devuelve "
                "el valor de otra columna en la misma fila. Usa VERDADERO para "
                "rangos (tarifas) y FALSO para coincidencia exacta (catalogos)."
            ),
            "ejemplo": (
                "=BUSCARV(B4, Tarifa_Anual_2026, 3, VERDADERO) — Obtiene la "
                "cuota fija del rango ISR correspondiente al ingreso."
            ),
        },
        {
            "nombre": "HOY",
            "sintaxis": "=HOY()",
            "descripcion": (
                "Devuelve la fecha actual del sistema. No requiere argumentos. "
                "Se actualiza cada vez que se recalcula la hoja."
            ),
            "ejemplo": (
                "=DIAS(\"17/04/2026\", HOY()) — Calcula los dias restantes "
                "para la fecha limite de la declaracion anual."
            ),
        },
        {
            "nombre": "FECHA",
            "sintaxis": "=FECHA(anio, mes, dia)",
            "descripcion": (
                "Construye una fecha a partir de tres componentes separados. "
                "Util para armar fechas desde datos dispersos o extraidos con EXTRAE."
            ),
            "ejemplo": (
                "=FECHA(1985, 3, 15) — Construye la fecha 15/03/1985 a partir "
                "de los datos extraidos del RFC CAST850315HN7."
            ),
        },
        {
            "nombre": "EXTRAE",
            "sintaxis": "=EXTRAE(texto, posicion_inicial, num_caracteres)",
            "descripcion": (
                "Extrae un numero determinado de caracteres de una cadena de texto, "
                "empezando en la posicion indicada. Devuelve texto (multiplicar por 1 "
                "para convertir a numero)."
            ),
            "ejemplo": (
                "=EXTRAE(\"CAST850315HN7\", 5, 2) — Extrae \"85\" (anio de nacimiento) "
                "del RFC de una persona fisica."
            ),
        },
    ]

    for func in funciones:
        gen.add_subsection(func["nombre"])
        data = [
            ["Sintaxis", func["sintaxis"]],
            ["Descripcion", func["descripcion"]],
            ["Ejemplo contable", func["ejemplo"]],
        ]
        gen.add_table(data, col_widths=[1.3 * inch, 5.2 * inch], header=False)
        gen.add_spacer(0.08)

    # ── Seccion 2: Tarifa ISR 2026 ─────────────────────────────────
    gen.add_page_break()
    gen.add_section("Tarifa ISR Anual 2026 — Art. 152 LISR (Anexo 8 RMF)")
    gen.add_text(
        "Tarifa actualizada por inflacion acumulada >10% desde noviembre 2022. "
        "Publicada en el DOF el 28 de diciembre de 2025."
    )
    gen.add_spacer(0.1)

    tarifa_header = ["Limite Inferior", "Limite Superior", "Cuota Fija", "% Excedente"]
    tarifa_data = [tarifa_header]
    for r in TARIFA_ANUAL:
        sup = "En adelante" if r["lim_sup"] > 999_999_999 else _fmt_money(r["lim_sup"])
        tarifa_data.append([
            _fmt_money(r["lim_inf"]),
            sup,
            _fmt_money(r["cuota"]),
            _fmt_pct(r["pct"]),
        ])

    gen.add_table(
        tarifa_data,
        col_widths=[1.6 * inch, 1.6 * inch, 1.6 * inch, 1.2 * inch],
        header=True,
    )

    gen.add_spacer(0.1)
    gen.add_text(
        "<b>Procedimiento de calculo:</b> (1) Ubicar la base gravable en la tarifa. "
        "(2) Restar el limite inferior. (3) Multiplicar el excedente por el porcentaje. "
        "(4) Sumar la cuota fija. Resultado = ISR del ejercicio."
    )

    # ── Seccion 3: Guia Factor de Actualizacion ────────────────────
    gen.add_page_break()
    gen.add_section("Guia: Factor de Actualizacion — Art. 17-A CFF")
    gen.add_text(
        "El Factor de Actualizacion ajusta montos fiscales por el efecto de la inflacion. "
        "Se calcula con el Indice Nacional de Precios al Consumidor (INPC) publicado por el INEGI."
    )
    gen.add_spacer(0.15)

    gen.add_subsection("Paso 1: Identificar los periodos")
    gen.add_text(
        "Determina el mes mas reciente del periodo de actualizacion (INPC reciente) "
        "y el mes mas antiguo (INPC anterior). Ejemplo: actualizacion de diciembre 2024 "
        "a diciembre 2025."
    )

    gen.add_subsection("Paso 2: Obtener los valores del INPC")
    gen.add_text(
        "Consulta los valores en la pagina del INEGI o en el DOF. "
        "INPC Dic 2025 = 141.200 (estimado). INPC Dic 2024 = 136.163."
    )

    gen.add_subsection("Paso 3: Dividir INPC reciente entre INPC anterior")
    gen.add_text(
        "Factor = 141.200 / 136.163 = 1.036996..."
    )

    gen.add_subsection("Paso 4: Truncar a 4 decimales (diezmilesimo)")
    gen.add_text(
        "El Art. 17-A CFF establece que el factor se trunca, NO se redondea. "
        "En Excel: =TRUNCAR(141.200/136.163, 4) = <b>1.0369</b>"
    )

    gen.add_subsection("Paso 5: Aplicar el factor al monto original")
    gen.add_text(
        "Monto actualizado = Monto original x Factor. "
        "Ejemplo: $100,000.00 x 1.0369 = <b>$103,690.00</b>"
    )

    gen.add_spacer(0.1)

    fa_data = [
        ["Concepto", "Valor", "Formula Excel"],
        ["INPC reciente (Dic 2025)", "141.200", "Celda B4"],
        ["INPC anterior (Dic 2024)", "136.163", "Celda B5"],
        ["Factor sin truncar", "1.036996...", "=B4/B5"],
        ["Factor truncado (CFF)", "1.0369", "=TRUNCAR(B4/B5, 4)"],
        ["Monto original", "$100,000.00", "Celda B9"],
        ["Monto actualizado", "$103,690.00", "=B9*B7"],
    ]
    gen.add_table(fa_data, col_widths=[2.2 * inch, 1.5 * inch, 2.5 * inch], header=True)

    # ── Seccion 4: Ejercicios con Respuestas ───────────────────────
    gen.add_page_break()
    gen.add_section("Ejercicios de Calculo ISR 2026 con Respuestas")
    gen.add_text(
        "Para cada escenario, calcula el ISR anual 2026 usando la tarifa del Art. 152 LISR. "
        "Las respuestas incluyen el desglose paso a paso."
    )
    gen.add_spacer(0.15)

    ejercicios = [
        {
            "num": 1,
            "titulo": "Empleado con sueldo fijo",
            "ingreso": 280000,
            "deducciones": 45000,
        },
        {
            "num": 2,
            "titulo": "Freelancer con ingresos variables",
            "ingreso": 520000,
            "deducciones": 120000,
        },
        {
            "num": 3,
            "titulo": "Socio de empresa (dividendos)",
            "ingreso": 1500000,
            "deducciones": 350000,
        },
        {
            "num": 4,
            "titulo": "Trabajador zona fronteriza",
            "ingreso": 180000,
            "deducciones": 30000,
        },
        {
            "num": 5,
            "titulo": "Director general (ingreso alto)",
            "ingreso": 3200000,
            "deducciones": 600000,
        },
    ]

    for ej in ejercicios:
        base = ej["ingreso"] - ej["deducciones"]
        resultado = calcular_isr_anual(base)

        gen.add_subsection(
            "Ejercicio {}: {}".format(ej["num"], ej["titulo"])
        )

        # Tabla de datos del ejercicio
        ej_data = [
            ["Concepto", "Monto"],
            ["Ingreso anual", _fmt_money(ej["ingreso"])],
            ["Deducciones autorizadas", _fmt_money(ej["deducciones"])],
            ["Base gravable", _fmt_money(base)],
        ]

        if resultado:
            ej_data.extend([
                ["Limite inferior", _fmt_money(resultado["lim_inf"])],
                ["Excedente", _fmt_money(resultado["excedente"])],
                ["% sobre excedente", _fmt_pct(resultado["pct"])],
                ["ISR marginal", _fmt_money(resultado["isr_marginal"])],
                ["Cuota fija", _fmt_money(resultado["cuota_fija"])],
                ["ISR del ejercicio", _fmt_money(resultado["isr_total"])],
            ])
        else:
            ej_data.append(["ISR del ejercicio", "Error: fuera de rango"])

        gen.add_table(ej_data, col_widths=[2.5 * inch, 2.5 * inch], header=True)
        gen.add_spacer(0.1)

    # ── Seccion 5: Atajos del Modulo ───────────────────────────────
    gen.add_page_break()
    gen.add_section("Atajos de Teclado del Modulo 1")
    gen.add_text(
        "Dominar estos atajos te ahorrara minutos cada dia y horas cada mes."
    )
    gen.add_spacer(0.1)

    atajos_data = [
        ["Atajo", "Accion", "Cuando usarlo"],
        ["F2", "Entrar a modo edicion de celda", "Para ver y editar formulas (debuggear)"],
        ["Ctrl + *", "Seleccionar toda la region de datos", "Para seleccionar una tabla completa rapido"],
        ["Tab", "Autocompletar funcion sugerida", "Cuando escribes =SU... y aparece SUMA"],
        ["Ctrl + Z", "Deshacer la ultima accion", "Cuando cometes un error (funciona multiples veces)"],
        ["Ctrl + Y", "Rehacer (repetir ultima accion)", "Para re-aplicar algo que deshiciste"],
        ["Alt + =", "Insertar SUMA automaticamente", "Debajo de una columna de numeros"],
        ["Ctrl + F3", "Administrador de nombres de rango", "Para nombrar rangos (VentasEnero, Tarifa, etc.)"],
        ["F4", "Alternar referencia absoluta ($)", "Para fijar celdas en formulas antes de arrastrar"],
        ["Ctrl + `", "Mostrar/ocultar formulas en la hoja", "Para revisar todas las formulas de un vistazo"],
        ["Ctrl + Shift + L", "Activar/desactivar filtros", "Para filtrar datos en una tabla rapidamente"],
    ]

    gen.add_table(
        atajos_data,
        col_widths=[1.2 * inch, 2.2 * inch, 3.0 * inch],
        header=True,
    )

    gen.add_spacer(0.2)
    gen.add_text(
        "<b>Tip:</b> Practica un atajo nuevo cada dia durante una semana. "
        "Al final del curso tendras mas de 40 atajos en tu memoria muscular."
    )

    # ── Guardar ────────────────────────────────────────────────────
    gen.save()


if __name__ == "__main__":
    build()
