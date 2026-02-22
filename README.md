# Excel para Contadores y Administrativos

**Automatiza tu Trabajo Contable y Administrativo**

Generador automatizado de todos los materiales del curso: ejercicios en Excel, guias de referencia, presentaciones y scripts de teleprompter.

## Sobre el curso

| | |
|---|---|
| **Instructor** | Israel Castro — CPA & Software Engineer |
| **Sitio** | todoconta.com |
| **Audiencia** | Contadores, auxiliares administrativos, profesionales que usan Excel |
| **Enfoque** | Ejemplos reales del ambito fiscal/contable mexicano |

## Temario

### Modulo 1: Logica Contable y Funciones de Control
- Sintaxis profesional: diferencia entre formula y funcion, argumentos y orden
- Precision fiscal: `TRUNCAR` vs `REDONDEAR` para factores de actualizacion (CFF, diezmilesimo)
- Logica y auditoria: `SI` condicional para validacion de bancos y estatus
- Calculos automatizados: `BUSCARV` aplicado a tarifas de ISR (limite inferior, cuota fija, tasa)
- Control de vencimientos: funciones de fecha (`HOY`, `FECHA`, `EXTRAE`) para alertas de e.firma

### Modulo 2: Procesamiento Masivo con Tablas Dinamicas
- De rango a tabla oficial: `Ctrl+*`, nombrar tablas (`xml_nomina`)
- Limpieza masiva: llenado rapido (`Ctrl+J`), estandarizacion de formatos
- Tablas dinamicas: percepciones vs deducciones, agrupacion por periodo/trabajador
- Papeles de trabajo: vinculacion dinamica a formatos de devolucion de impuestos

### Modulo 3: Visualizacion de Impacto y Reportes Ejecutivos
- Graficos dinamicos vinculados a tablas dinamicas
- Personalizacion: colores con logica (Magna=verde, Premium=rojo), limpieza visual
- Comparativas anuales: ventas vs egresos entre ejercicios, estacionalidad

### Modulo 4: El Dashboard Inteligente y Entrega Profesional
- Diseno de dashboards: segmentadores (slicers), maquetacion "menos es mas"
- Proteccion: celdas bloqueadas/desbloqueadas, contrasenas, cifrado AES
- Distribucion: PDF vs Excel protegido, convenciones de nombres profesionales
- Checklist de entrega profesional

### Modulo 5: Automatizacion Nativa con Microsoft 365 Copilot
- Activacion y preparacion de datos (formato Tabla + OneDrive/SharePoint)
- Analisis con lenguaje natural, generacion de formulas, visualizacion instantanea
- 20 prompts organizados en 5 categorias para contadores
- Validacion: por que el criterio del contador sigue siendo indispensable

### Bonus
- **VBA con IA**: Guia para generar macros con ChatGPT/Claude
- **Claude en Excel**: Uso de Claude como asistente contable
- **Cheat Sheet de Atajos**: 80+ atajos organizados en 9 categorias

## Estructura del proyecto

```
scripts/                    # Codigo fuente de los generadores
  config/                   # Constantes, tarifas ISR, estilos
  generators/               # Clases base: ExcelGenerator, MarkdownGenerator, SlidesGenerator
  modulo{1-5}/              # Generadores por modulo
  bonus/                    # Generadores de material bonus
  build_all.py              # Orquestador maestro
  requirements.txt          # Dependencias Python

output/                     # Materiales generados
  Pack_Excel_Pro/           # Material descargable del alumno
    Modulo_1_Funciones/     # .xlsx ejercicios + .md referencia
    Modulo_2_Tablas_Dinamicas/
    Modulo_3_Visualizacion/
    Modulo_4_Dashboard/
    Modulo_5_Copilot_IA/
    Bonus/                  # Guias .md (VBA, Claude, Atajos)
  Slides/                   # .md para gamma.app (1 por modulo/bonus)
  Teleprompter/             # .md scripts de narracion para grabar videos
```

## Material generado (35 archivos)

| Tipo | Cantidad | Formato | Destino |
|------|----------|---------|---------|
| Ejercicios Excel | 12 | .xlsx | Pack del alumno |
| Guias de referencia | 5 | .md | Notion (documentos) |
| Guias extras | 4 | .md | Notion (documentos) |
| Slides | 7 | .md | gamma.app (presentaciones) |
| Teleprompter | 7 | .md | Script para grabar videos |

## Uso

### Requisitos

```bash
pip install -r scripts/requirements.txt
```

### Generar todo

```bash
python3 -m scripts.build_all
```

### Solo verificar outputs existentes

```bash
python3 -m scripts.build_all --verify
```

### Generar un modulo individual

```bash
python3 -c "from scripts.modulo1.gen_slides_m1 import build; build()"
```

## Flujo de trabajo

1. Los **scripts** en `scripts/` contienen toda la logica y contenido del curso
2. `build_all.py` ejecuta cada generador en orden y verifica los outputs
3. Los **.xlsx** se entregan al alumno como ejercicios interactivos
4. Los **.md de referencia** se suben a Notion como documentos del curso
5. Los **.md de slides** se importan en gamma.app para crear presentaciones
6. Los **.md de teleprompter** se usan como guion al grabar cada video

## Datos fiscales

Todos los ejemplos usan datos de Mexico 2026:
- Salario minimo diario: $278.80
- Tarifas ISR: Art. 96 LISR (definidas en `scripts/config/isr_2026.py`)
- Precios combustible: Magna $23-25, Premium $25-27, Diesel $25-27

## Licencia

Uso privado — Israel Castro / todoconta.com
