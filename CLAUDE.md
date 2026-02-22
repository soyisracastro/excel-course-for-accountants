# Excel para Contadores y Administrativos

## Sobre el proyecto

Generador automatizado de materiales para el curso **"Excel para Contadores y Administrativos: Automatiza, Analiza y Grafica"** impartido por **Israel Castro** (CPA & Software Engineer) a traves de **todoconta.com**.

El curso esta dirigido a contadores, auxiliares administrativos y profesionales que trabajan con Excel en el ambito fiscal/contable mexicano. El enfoque es 100% practico con ejemplos reales: calculo de ISR, procesamiento de XMLs de nomina del SAT, factores de actualizacion del CFF, control de vencimiento de e.firma, dashboards ejecutivos y uso de Microsoft 365 Copilot.

## Estructura del curso (5 modulos + bonus)

| Modulo | Nombre | Entregable principal |
|--------|--------|---------------------|
| 1 | Logica Contable y Funciones de Control | Calculadora ISR, Control Vencimientos e.firma, Extraccion RFC |
| 2 | Procesamiento Masivo con Tablas Dinamicas | Limpieza masiva, Analisis nomina XML, Papel de trabajo referenciado |
| 3 | Visualizacion de Impacto y Reportes Ejecutivos | Dashboard ventas combustible, Comparativa anual ventas/gastos |
| 4 | El Dashboard Inteligente y Entrega Profesional | Layout dashboard contable, Dashboard final integrado, Guia proteccion |
| 5 | Automatizacion Nativa con Microsoft 365 Copilot | Dataset para Copilot, Guia de prompts contables |
| Bonus | VBA con IA + Claude en Excel + Atajos | Guias PDF y cheat sheet |

## Arquitectura del proyecto

```
scripts/
  config/
    constants.py       # Rutas, colores, metadatos, modulos (fuente de verdad)
    isr_2026.py        # Tarifas ISR Art. 96 LISR 2026
    styles.py          # Estilos compartidos para openpyxl
  generators/
    xlsx_gen.py        # ExcelGenerator (openpyxl) — genera .xlsx
    md_gen.py          # MarkdownGenerator — genera .md (reemplazo de PDFGenerator)
    pptx_gen.py        # SlidesGenerator — genera .md slides + .md teleprompter
    pdf_gen.py         # PDFGenerator (reportlab) — LEGACY, ya no se usa
  modulo{1-5}/         # Scripts generadores por modulo
  bonus/               # Scripts generadores de material bonus
  build_all.py         # Orquestador: ejecuta todos los generadores y verifica outputs

output/                # Directorio de salida (generado, no editar manualmente)
  Pack_Excel_Pro/      # Material descargable del alumno (.xlsx + .md)
  Slides/              # Markdown para gamma.app (1 archivo por modulo)
  Teleprompter/        # Scripts de narracion para grabar videos
```

## Convenciones

- **Idioma del codigo**: Espanol para contenido, ingles para nombres de clases/metodos
- **Generadores**: Cada script tiene una funcion `build()` que es llamada por `build_all.py`
- **Salida de documentos**: Todo es Markdown (.md). Los PDFs y PPTX fueron migrados a .md
  - PDFs → .md via `MarkdownGenerator` (para subir a Notion)
  - Slides → .md via `SlidesGenerator` (para crear presentaciones en gamma.app)
  - Teleprompter → .md separado (script de narracion para grabar cada video)
- **Excel (.xlsx)**: Se siguen generando como binarios con openpyxl/xlsxwriter
- **Colores**: Definidos en `constants.py` clases `Color` (hex) y `ColorRGB` (tuplas RGB)
- **Datos fiscales**: Mexico 2026 — salario minimo $278.80/dia, tarifas ISR en `isr_2026.py`

## Comandos

```bash
# Generar todo
python3 -m scripts.build_all

# Solo verificar que los outputs existan
python3 -m scripts.build_all --verify
```

## Dependencias

- `openpyxl`, `xlsxwriter` — generacion de .xlsx
- `Faker` — datos sinteticos para dataset de Copilot
- `Pillow` — manejo de imagenes en xlsx
- `reportlab`, `python-pptx` — LEGACY, ya no se usan activamente

## Notas importantes

- Los archivos en `output/` con extension `.xlsx`, `.pdf`, `.pptx` estan en `.gitignore` (binarios)
- Los `.md` generados en `output/` SI se trackean en git
- `constants.py` es la fuente de verdad para rutas, nombres de modulos y metadatos
- El parametro `col_widths` en `add_table()` de `MarkdownGenerator` se acepta pero se ignora (compatibilidad API con el antiguo `PDFGenerator`)
