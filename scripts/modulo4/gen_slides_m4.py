"""
Generador: Modulo_4_Dashboard.pptx + Script_Modulo_4.md
Modulo 4 -- El Dashboard Inteligente y Entrega Profesional

12-14 slides sobre dashboards, slicers, proteccion y entrega profesional.
Script de teleprompter ~25-30 min en espanol.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.config.constants import (
    PACK, SLIDES_DIR, TELEPROMPTER_DIR, MODULOS, ColorRGB
)
from scripts.generators.pptx_gen import (
    SlidesGenerator, PPTX_AZUL, PPTX_VERDE, PPTX_ROJO,
    PPTX_TEXTO, PPTX_BLANCO, PPTX_TEXTO_MEDIO,
    PP_ALIGN, Pt, Inches, RGBColor
)

M = MODULOS[4]


def build():
    gen = SlidesGenerator(
        filename=M["slide_nombre"],
        output_dir=SLIDES_DIR,
        script_filename=M["script_nombre"],
        script_dir=TELEPROMPTER_DIR,
    )

    # ---- Slide 1: Portada ------------------------------------------------
    gen.add_title_slide(
        modulo_num=4,
        modulo_nombre=M["nombre"],
        subtitulo="ISR + Nomina + Graficos = Dashboard Profesional"
    )
    gen.script_lines.append(
        "Bienvenidos al Modulo 4, el modulo donde todo lo que hemos construido "
        "se une en un dashboard profesional. Vamos a tomar la calculadora ISR "
        "del Modulo 1, los datos de nomina del Modulo 2, los graficos del Modulo 3, "
        "y vamos a integrarlos en un panel de control que impresione. "
        "Ademas, veremos como proteger y distribuir nuestro trabajo de forma profesional.\n"
    )

    # ---- Slide 2: Que es un Dashboard? -----------------------------------
    gen.add_content_slide(
        title="Que es un Dashboard?",
        bullets=[
            "Un panel de control visual que resume informacion clave en una sola pantalla",
            "Permite tomar decisiones rapidas sin revisar datos crudos",
            "Combina: KPIs (numeros clave) + Graficos + Filtros interactivos",
            "En contabilidad: nomina, ISR, flujo de efectivo, vencimientos",
            "NO es un reporte completo -- es un resumen ejecutivo visual",
        ],
        script_text=(
            "Un dashboard es basicamente un tablero de control, como el tablero de tu auto. "
            "No necesitas abrir el motor para saber si todo esta bien -- solo ves los indicadores. "
            "En contabilidad, un dashboard te muestra de un vistazo cuanto llevas de nomina, "
            "cuanto de ISR, si alguna e.firma esta por vencer. "
            "Es un resumen ejecutivo visual, no un reporte completo. "
            "Y lo mejor: en Excel podemos construirlo sin programar, usando las herramientas "
            "que ya conocemos.\n"
        )
    )

    # ---- Slide 3: Principios de diseno -----------------------------------
    gen.add_content_slide(
        title="Principios de Diseno: Menos es Mas",
        bullets=[
            "Regla de los 5 segundos: el mensaje principal debe entenderse en 5 segundos",
            "Maximo 4-6 KPIs visibles (no satures la pantalla)",
            "Paleta de colores consistente (azul = principal, rojo = alerta, verde = ok)",
            "Graficos simples: barras para comparar, lineas para tendencias",
            "Elimina ruido visual: sin bordes de cuadricula, sin decoraciones innecesarias",
            "Jerarquia visual: KPIs arriba, graficos al centro, detalle abajo",
        ],
        script_text=(
            "Antes de construir, hablemos de diseno. La regla mas importante es: menos es mas. "
            "Tu dashboard debe comunicar el mensaje principal en 5 segundos. "
            "Si alguien necesita 30 segundos para entender que pasa, has fallado. "
            "Maximo 4 a 6 KPIs. Colores consistentes: nosotros usamos azul como color principal, "
            "rojo para alertas, verde para 'todo bien'. "
            "Los graficos deben ser simples: barras para comparar y lineas para tendencias. "
            "Y algo que muchos olvidan: oculta las lineas de cuadricula. "
            "Eso solo da un aspecto mucho mas profesional.\n"
        )
    )

    # ---- Slide 4: Segmentadores de Datos (Slicers) -----------------------
    gen.add_content_slide(
        title="Segmentadores de Datos (Slicers)",
        bullets=[
            "Filtros visuales e interactivos para Tablas Dinamicas",
            "El usuario hace clic en botones para filtrar -- sin menus complicados",
            "Como insertarlos: clic en TD > Insertar > Segmentacion de datos",
            "Selecciona los campos: Periodo, Puesto, Departamento, etc.",
            "Se pueden vincular a MULTIPLES Tablas Dinamicas simultaneamente",
            "El resultado: un dashboard interactivo sin necesidad de macros",
        ],
        script_text=(
            "Los segmentadores, o slicers en ingles, son la magia detras de un dashboard interactivo. "
            "Son filtros visuales que aparecen como botones. "
            "El usuario simplemente hace clic en 'Enero' y todo el dashboard se actualiza. "
            "Para insertarlos, primero necesitas una Tabla Dinamica. "
            "Luego vas a Insertar, Segmentacion de datos, y seleccionas los campos que quieres filtrar. "
            "Lo mas poderoso es que puedes conectar un solo slicer a multiples Tablas Dinamicas. "
            "Asi, un clic filtra todos los graficos y KPIs al mismo tiempo.\n"
        )
    )

    # ---- Slide 5: Caso - Slicers vinculados a TDs ------------------------
    gen.add_content_slide(
        title="Caso Practico: Slicers Vinculados a TDs",
        bullets=[
            "Paso 1: Crear TD1 con Periodo en filas, Sueldo/ISR en valores",
            "Paso 2: Crear TD2 con Puesto en filas, conteo de Empleados",
            "Paso 3: Insertar slicer de 'Periodo' desde TD1",
            "Paso 4: Clic derecho en slicer > Conexiones de informe > marcar TD2",
            "Paso 5: Ahora al filtrar por mes, AMBAS tablas se actualizan",
            "Paso 6: Los graficos basados en las TDs tambien se actualizan automaticamente",
        ],
        script_text=(
            "Vamos a verlo paso a paso. Primero creamos dos Tablas Dinamicas "
            "desde nuestra tabla de nomina. La primera muestra sueldos e ISR por periodo. "
            "La segunda muestra empleados por puesto. "
            "Luego insertamos un segmentador de Periodo. Por defecto solo esta conectado a TD1. "
            "Pero si hacemos clic derecho en el slicer y vamos a Conexiones de informe, "
            "podemos marcar tambien TD2. Ahora, cuando seleccionamos 'Marzo', "
            "ambas tablas y sus graficos se filtran. "
            "Esto es lo que hace que un dashboard en Excel sea realmente interactivo.\n"
        )
    )

    # ---- Slide 6: Proyecto Final - Integracion ----------------------------
    slide6 = gen.add_content_slide(
        title="Proyecto Final: La Gran Integracion",
        bullets=[
            "Modulo 1: Calculadora ISR con BUSCARV y tarifa oficial 2026",
            "Modulo 2: Datos de nomina 20 empleados x 12 meses + Tablas Dinamicas",
            "Modulo 3: Graficos de barras (comparacion) y lineas (tendencia)",
            "Modulo 4: TODO junto en un Dashboard con slicers y KPIs",
            "Archivo de trabajo: 10_Dashboard_Final_Integrado.xlsx",
            "Resultado: un panel de control de nomina completo y profesional",
        ],
        script_text=(
            "Este es el momento de la verdad. Todo lo que hemos aprendido converge aqui. "
            "Del Modulo 1 tomamos la logica de calculo ISR con BUSCARV. "
            "Del Modulo 2, los datos masivos de nomina y las Tablas Dinamicas. "
            "Del Modulo 3, los graficos y la visualizacion. "
            "Y ahora en el Modulo 4, los unimos en un dashboard interactivo. "
            "El archivo de trabajo es el Dashboard Final Integrado. "
            "Tiene 4 hojas: datos de nomina, tarifa ISR, calculadora y el dashboard. "
            "Vamos a ver como se construye.\n"
        )
    )

    # ---- Slide 7: KPIs ---------------------------------------------------
    gen.add_content_slide(
        title="KPIs: Los Numeros que Importan",
        bullets=[
            "KPI 1 -- Total Percepciones: =SUBTOTAL(109, Nomina[Sueldo])",
            "KPI 2 -- Total Deducciones: =SUBTOTAL(109, Nomina[ISR]) + SUBTOTAL(109, Nomina[IMSS])",
            "KPI 3 -- ISR del Periodo: =SUBTOTAL(109, Nomina[ISR])",
            "KPI 4 -- e.firma Status: formula SI con dias restantes",
            "SUBTOTAL(109,...) respeta filtros de Tablas Dinamicas y Slicers",
            "Alternativa: =GETPIVOTDATA() para extraer valores directamente de TDs",
        ],
        script_text=(
            "Los KPIs son los numeros clave que van en la parte superior del dashboard. "
            "Nosotros usamos 4: Total Percepciones, que es la suma de todos los sueldos. "
            "Total Deducciones, que incluye ISR e IMSS. ISR del Periodo como desglose. "
            "Y el estatus de la e.firma, que conecta con lo que vimos en el Modulo 1. "
            "Un detalle tecnico importante: usamos SUBTOTAL con el codigo 109, que es SUMA "
            "pero ignorando filas ocultas por filtros. "
            "Asi, cuando el usuario filtra con un slicer, los KPIs se recalculan. "
            "Si quieres aun mas precision, puedes usar GETPIVOTDATA para leer directamente "
            "desde la Tabla Dinamica.\n"
        )
    )

    # ---- Slide 8: Construyendo el Layout ---------------------------------
    gen.add_content_slide(
        title="Construyendo el Layout del Dashboard",
        bullets=[
            "Fila 1: Titulo del dashboard (fuente grande, color azul)",
            "Filas 2-5: Cuadros de KPI (4 cajas con colores de semaforo)",
            "Columna A-B: Zona de segmentadores (slicers)",
            "Filas 7-16: Graficos principales (2 graficos lado a lado)",
            "Filas 18-25: Tabla de detalle o grafico adicional",
            "Archivo template: 09_Layout_Dashboard_Contable.xlsx",
        ],
        script_text=(
            "El layout es la estructura visual de tu dashboard. "
            "Piensa en el como el plano de una casa antes de construir. "
            "Arriba va el titulo. Debajo, los KPIs en cuadros de colores. "
            "A la izquierda, la zona de segmentadores. "
            "En el centro, los graficos principales -- normalmente dos: uno de barras y uno de linea. "
            "Abajo, espacio para una tabla de detalle o un grafico adicional. "
            "En el archivo template que les di, ya tienen esta estructura lista. "
            "Solo necesitan llenarla con sus Tablas Dinamicas y graficos.\n"
        )
    )

    # ---- Slide 9: Proteccion de Celdas -----------------------------------
    gen.add_content_slide(
        title="Proteccion de Celdas y Hojas",
        bullets=[
            "Todas las celdas vienen bloqueadas por defecto (pero inactivo hasta proteger)",
            "Paso 1: DESBLOQUEA las celdas de entrada (Formato > Proteccion > desmarcar Bloqueada)",
            "Paso 2: Resalta las celdas editables con fondo amarillo",
            "Paso 3: Revisar > Proteger hoja > establece contrasena",
            "Permite: seleccionar celdas, usar filtros y Tablas Dinamicas",
            "Resultado: el usuario puede interactuar pero no romper las formulas",
        ],
        script_text=(
            "Ahora que tu dashboard funciona, hay que protegerlo. "
            "No quieres que alguien borre una formula por accidente. "
            "El truco es que en Excel todas las celdas ya vienen bloqueadas, "
            "pero el bloqueo no se activa hasta que proteges la hoja. "
            "Entonces el flujo es: primero, desbloqueas las celdas donde el usuario SI debe escribir. "
            "Segundo, las marcas con fondo amarillo para que sepa donde ir. "
            "Tercero, proteges la hoja con contrasena. "
            "Y en las opciones de proteccion, asegurate de permitir: seleccionar celdas, "
            "usar filtros y usar Tablas Dinamicas. Asi el dashboard sigue siendo interactivo "
            "pero las formulas quedan protegidas.\n"
        )
    )

    # ---- Slide 10: Distribucion ------------------------------------------
    gen.add_content_slide(
        title="Distribucion: PDF vs Excel Protegido",
        bullets=[
            "PDF: para reportes finales, firma digital, envio a clientes",
            "   -- No editable, no requiere Excel, archivo ligero",
            "Excel protegido: para plantillas, calculadoras interactivas",
            "   -- Mantiene formulas, slicers y TDs activas",
            "Contrasena de apertura (AES): seguridad real del archivo",
            "Contrasena de escritura: permite lectura sin edicion",
            "Marcar como final: solo una sugerencia, no proteccion real",
        ],
        script_text=(
            "La pregunta final es: como lo comparto? Hay dos opciones principales. "
            "PDF es ideal para reportes finales. No se puede editar, no requiere Excel, "
            "y puedes firmarlo digitalmente. "
            "Excel protegido es para cuando el destinatario necesita interactuar: "
            "cambiar filtros, usar slicers, ingresar datos. "
            "Si necesitas seguridad real, usa la contrasena de apertura de archivo, "
            "que usa cifrado AES. Eso es como ponerle llave al archivo. "
            "La contrasena de escritura permite que lo abran pero no lo modifiquen. "
            "Y 'Marcar como final' es solo una sugerencia, el usuario la puede quitar. "
            "Mi recomendacion: para el cliente, PDF. Para tu equipo, Excel protegido.\n"
        )
    )

    # ---- Slide 11: Demo Dashboard Final ----------------------------------
    gen.add_content_slide(
        title="Demo: Dashboard Final en Accion",
        bullets=[
            "Abre el archivo: 10_Dashboard_Final_Integrado.xlsx",
            "1. Revisa la hoja Datos_Nomina (240 registros reales)",
            "2. Ve a Calculadora y cambia el sueldo en B4 -- observa el ISR",
            "3. Crea una Tabla Dinamica desde Datos_Nomina",
            "4. Inserta Segmentadores de Periodo y Puesto",
            "5. Crea graficos de barras y lineas",
            "6. Mueve todo a la hoja Dashboard y ajusta el layout",
            "7. Protege las hojas y guarda como PDF + Excel protegido",
        ],
        script_text=(
            "Vamos a hacer la demo paso a paso. Abran el archivo Dashboard Final Integrado. "
            "Primero, revisen la hoja Datos Nomina. Son 240 registros: 20 empleados por 12 meses. "
            "Los sueldos van desde salario minimo hasta 55 mil pesos. "
            "Luego vayan a la Calculadora y cambien el sueldo en B4. "
            "Vean como BUSCARV calcula automaticamente el ISR. "
            "Ahora, creen una Tabla Dinamica con Periodo en filas y Sueldo en valores. "
            "Inserten un segmentador de Periodo. Creen un grafico de barras. "
            "Hagan lo mismo con otra TD por Puesto. "
            "Muevan todo a la hoja Dashboard. Ajusten tamanos. "
            "Finalmente, protejan las hojas y exporten un PDF. "
            "Eso es su proyecto final del modulo.\n"
        )
    )

    # ---- Slide 12: Resumen -----------------------------------------------
    gen.add_content_slide(
        title="Resumen del Modulo 4",
        bullets=[
            "Un dashboard es un resumen ejecutivo visual -- no un reporte completo",
            "Los segmentadores (slicers) hacen tu dashboard interactivo sin macros",
            "KPIs con SUBTOTAL(109,...) respetan filtros automaticamente",
            "El layout sigue la jerarquia: KPIs > Graficos > Detalle",
            "Proteccion: desbloquea inputs, bloquea formulas, protege hoja",
            "Distribuye como PDF (reportes) o Excel protegido (interactivos)",
            "Este modulo integra TODO: ISR (M1) + Nomina (M2) + Graficos (M3)",
        ],
        script_text=(
            "Hagamos un resumen rapido. Un dashboard es un resumen visual. "
            "Los slicers lo hacen interactivo. Los KPIs con SUBTOTAL se actualizan con filtros. "
            "El layout sigue una jerarquia clara. "
            "Protege tus formulas y comparte correctamente. "
            "Y lo mas importante: este modulo no es un tema aislado. "
            "Es la integracion de todo el curso. "
            "La calculadora ISR, los datos de nomina, los graficos, "
            "todo converge en un dashboard profesional. "
            "Esto es lo que puedes presentar a tu jefe o a tu cliente.\n"
        )
    )

    # ---- Slide 13: Recursos y Cierre -------------------------------------
    gen.add_closing_slide(
        next_module="Modulo 5 -- Automatizacion Nativa con Microsoft 365 Copilot",
        resources=[
            "Archivo: 09_Layout_Dashboard_Contable.xlsx (template de layout)",
            "Archivo: 10_Dashboard_Final_Integrado.xlsx (datos + calculadora + dashboard)",
            "PDF: 11_Guia_Proteccion_y_Seguridad.pdf (checklist de proteccion)",
            "PDF: Referencia_Modulo_4.pdf (guia rapida de diseno y slicers)",
            "Practica: Construye tu propio dashboard con datos reales de tu empresa",
        ]
    )
    gen.script_lines.append(
        "Estos son los recursos del modulo. Tienen el template de layout, "
        "el archivo integrado con datos, la guia de proteccion en PDF "
        "y la referencia rapida del modulo. "
        "Mi recomendacion es que practiquen con datos reales de su propia empresa. "
        "Tomen su nomina real, cambien los nombres si quieren por privacidad, "
        "y construyan su propio dashboard. "
        "En el siguiente modulo veremos como Microsoft 365 Copilot puede ayudarnos "
        "a automatizar aun mas nuestro trabajo con inteligencia artificial. "
        "Nos vemos ahi!\n"
    )

    gen.save()


if __name__ == "__main__":
    build()
