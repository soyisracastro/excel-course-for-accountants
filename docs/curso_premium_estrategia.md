# Estrategia: Curso Premium de Excel para Contadores (Mexico)

**Autor:** Israel Castro — CPA & Software Engineer
**Fecha:** Febrero 2026
**Sitio:** todoconta.com

---

## 1. Posicionamiento

El curso actual ("Excel para Contadores y Administrativos") cubre funciones, tablas
dinamicas, graficos, dashboards y una introduccion a Copilot. Es un producto solido
como **puerta de entrada**, pero su alcance lo posiciona en el rango de precio bajo
del mercado ($799-$1,199 MXN).

El curso premium ocupa un espacio que **ningun competidor cubre completo**: la
combinacion de Power Query + Power Pivot/DAX + VBA + herramientas fiscales
mexicanas + IA validada + proyecto integrador, todo en un solo programa.

| | Curso Basico (actual) | Curso Premium (propuesta) |
|---|---|---|
| **Nivel** | Introductorio | Intermedio-Avanzado |
| **Enfoque** | Funciones, Tablas Dinamicas, Dashboards | Power Query + Power Pivot + VBA + Fiscal + IA |
| **Duracion** | ~8 horas | 42-50 horas |
| **Precio sugerido** | $799-$1,199 MXN | $5,990 MXN |
| **Formato** | Video grabado + ejercicios | Video + ejercicios + sesiones en vivo + proyecto final |

---

## 2. Analisis del Mercado (Mexico, 2026)

### 2.1 Rangos de precio

| Segmento | Rango | Ejemplos |
|---|---|---|
| Commodity (Udemy, cupones) | $179-$632 MXN | Cursos genericos de Excel |
| Profesional (nicho contable) | $1,000-$5,500 MXN | ContadorMx, COFIDE, cursos especializados |
| Institucional (universidades) | $5,000-$25,500 MXN | IMECAF, Tec de Monterrey |

### 2.2 Competidores directos

| Competidor | Precio | Horas | Precio/hora | Fortalezas | Debilidades |
|---|---|---|---|---|---|
| IMECAF (presencial) | $5,046 | 20 | $252 | Reconocimiento, DC-3 STPS | Solo presencial CDMX, Excel generico |
| Tec de Monterrey | $25,500 | 40 | $637 | Marca, credencial universitaria | Precio prohibitivo, no fiscal |
| COFIDE | $632 curso + $5,490/año | Ilimitado | — | Comunidad, membresia | Contenido disperso, no profundiza |
| ContadorMx | $979-$1,149 | ~10 | $115 | Nicho contable, precio accesible | Corto, sin Power Query ni VBA |
| Udemy (varios) | $179-$299 | 5-15 | $20-$60 | Precio bajo, acceso perpetuo | Generico, sin contexto fiscal MX |

### 2.3 Brecha de mercado

**Ningun curso en Mexico combina los 5 pilares en un solo programa:**

1. Power Query / Power Pivot / DAX
2. VBA y Macros
3. Herramientas fiscales mexicanas (ISR, IVA, DIOT, CFDI, SAT)
4. IA aplicada (Copilot, ChatGPT, Claude) con ejemplos validados
5. Dashboard ejecutivo + proyecto integrador con retroalimentacion

Este es el espacio que el curso premium debe ocupar.

---

## 3. Analisis del Contenido Actual

### 3.1 Fortalezas

- **Modulo 1 (Funciones/ISR):** Fuerte y accionable. Ejemplos reales con tarifas
  Art. 96 LISR, factores de actualizacion (CFF), extraccion de RFC. El alumno
  sale haciendo cosas utiles desde la primera sesion.
- **Ejercicios generados programaticamente:** Datos realistas, consistentes y
  actualizables con un solo comando. Ventaja competitiva enorme vs crear
  ejercicios manualmente.
- **Estructura modular:** Cada modulo es independiente con sus ejercicios,
  referencia, slides y teleprompter.

### 3.2 Areas de mejora para el premium

- **Modulo 4 (Dashboard):** Conceptual pero le falta profundidad tecnica.
  El premium debe incluir dashboards financieros reales (P&L, flujo de efectivo).
- **Modulo 5 (Copilot):** Actualmente lee como material de marketing. Los prompts
  no estan probados contra datos reales. El premium debe incluir ejemplos
  validados con resultados documentados.
- **Gaps transversales:** No hay cobertura de Power Query, Power Pivot, manejo
  de errores, ni validacion de datos. Estos son pilares del premium.

---

## 4. Temario del Curso Premium (8 Modulos)

### Modulo P1: Power Query — Automatizacion de Importacion Fiscal (8 hrs)

**Objetivo:** Eliminar la captura manual de datos fiscales.

- Importar XMLs de CFDI masivamente (carpetas completas de facturas)
- Conectar a layouts bancarios (Banamex, BBVA, Banorte) y transformar a formato contable
- Consolidar multiples archivos: balanzas mensuales, estados de cuenta, auxiliares
- Limpieza automatizada: RFC duplicados, nombres inconsistentes, errores de captura del SAT
- Actualizacion con un clic: agregar el mes 13 sin rehacer nada

**Proyecto:** Pipeline que importa 12 meses de XMLs y genera el auxiliar de IVA
listo para DIOT.

### Modulo P2: Power Pivot y Modelo de Datos (6 hrs)

**Objetivo:** Conectar multiples tablas contables en un modelo relacional sin
duplicar datos.

- Relaciones entre tablas: Catalogo de cuentas <-> Polizas <-> Auxiliares <-> CFDI
- Medidas DAX para contadores: acumulados anuales, comparativos vs presupuesto, variaciones %
- Tablas de calendario fiscal (meses 1-12, ejercicio fiscal, periodos de declaracion)
- KPIs contables: razon circulante, prueba del acido, rotacion de cuentas por cobrar
- Diferencia entre columnas calculadas y medidas (y cuando usar cada una)

**Proyecto:** Modelo de datos que conecta balanza + auxiliares + CFDIs en un solo
libro con analisis interactivo.

### Modulo P3: VBA y Macros para Automatizacion Contable (8 hrs)

**Objetivo:** Automatizar tareas repetitivas del ciclo contable con macros.

- Macros grabadas: formato de polizas, impresion de balanzas, proteccion masiva
- VBA con IA: usar ChatGPT/Claude para generar macros (expansion de la guia bonus actual)
- Automatizacion real: generar 12 hojas de trabajo mensuales con un boton
- Interaccion con archivos: leer TXT del SAT (32D, opinion de cumplimiento), procesar layouts
- Formularios UserForm: captura de polizas, catalogo de proveedores, control de activos fijos
- Debugging: como leer y corregir errores en macros generadas por IA

**Proyecto:** Macro que genera el papel de trabajo de ISR anual desde la balanza
de comprobacion.

### Modulo P4: Herramientas Fiscales Avanzadas en Excel (6 hrs)

**Objetivo:** Construir papeles de trabajo fiscales auditables y vinculados.

- Calculo de pagos provisionales ISR (personas morales y fisicas) con acumulacion
- Determinacion de IVA: acreditamiento, retenciones, proporcion de actos gravados
- Factor de actualizacion y recargos (Art. 17-A CFF) con INPC desde tabla dinamica
- Papeles de trabajo para DIOT: cruce de CFDIs vs registros contables
- Conciliacion contable-fiscal: partidas no deducibles, ajuste anual por inflacion

**Proyecto:** Libro maestro de determinacion de ISR anual con 5 hojas vinculadas
(ingresos, deducciones, coeficiente de utilidad, pagos provisionales, anual).

### Modulo P5: Dashboards Ejecutivos Avanzados (4 hrs)

**Objetivo:** Presentar informacion financiera de forma ejecutiva e interactiva.

- Dashboard financiero completo: P&L interactivo, flujo de efectivo, balance condensado
- Segmentadores conectados entre multiples tablas dinamicas
- Formato condicional avanzado: semaforos para vencimientos, alertas de desviacion presupuestal
- Sparklines y mini-graficos para tendencias mensuales
- Diseno "menos es mas": que mostrar, que ocultar, como guiar la lectura

**Proyecto:** Dashboard de "salud financiera" del contribuyente para presentar
al cliente o al director.

### Modulo P6: Copilot e IA Aplicada (con ejemplos validados) (4 hrs)

**Objetivo:** Usar IA como herramienta de apoyo con criterio profesional.

- Configuracion real: licenciamiento M365, requisitos tecnicos, limitaciones honestas
- 10 casos de uso probados y validados (no marketing):
  - Analisis de nomina
  - Deteccion de anomalias en datos
  - Generacion de formulas complejas
  - Resumen ejecutivo automatico
- Cuando SI y cuando NO usar Copilot (criterio profesional)
- Alternativas: ChatGPT con Advanced Data Analysis, Claude con archivos Excel
- Comparativa honesta: que hace bien, que hace mal, donde miente

**Proyecto:** Analisis comparativo — mismo ejercicio hecho manual vs con IA,
midiendo tiempo y precision.

### Modulo P7: Proteccion, Distribucion y Estandares Profesionales (2 hrs)

**Objetivo:** Entregar trabajo con calidad de despacho profesional.

- Proteccion real con cifrado AES (no solo contrasena de hoja)
- Convenciones de nombres para archivos de trabajo (cliente, ejercicio, version)
- Distribucion profesional: PDF firmado, Excel protegido, OneDrive compartido
- Control de versiones para libros de trabajo (nomenclatura y bitacora de cambios)
- Checklist de entrega: que revisar antes de enviar al cliente/jefe

### Modulo P8: Proyecto Integrador + Evaluacion (4 hrs)

**Objetivo:** Demostrar dominio integrando todas las habilidades.

- Caso practico completo: empresa ficticia con 12 meses de operaciones
- El alumno debe entregar:
  - Balanza procesada con Power Query
  - Modelo relacional en Power Pivot
  - Dashboard ejecutivo interactivo
  - Macro de automatizacion funcional
  - Determinacion de impuestos vinculada
- Evaluacion con rubrica detallada
- Constancia de participacion (DC-3 STPS si aplica, o constancia todoconta.com)

---

## 5. Precio y Justificacion

### 5.1 Precio recomendado

**$5,990 MXN** (pago unico) o **3 pagos de $2,290 MXN**

Early-bird / lanzamiento: **$4,490 MXN** (primeras 50 inscripciones)

### 5.2 Comparativa precio/hora

| Curso | Precio | Horas | Precio/hora |
|---|---|---|---|
| IMECAF (presencial) | $5,046 | 20 | $252 |
| Tec de Monterrey | $25,500 | 40 | $637 |
| COFIDE (membresia) | $5,490/año | ilimitado | — |
| ContadorMx | $1,149 | ~10 | $115 |
| **Curso Premium todoconta** | **$5,990** | **42-50** | **$120-$143** |

El precio/hora es competitivo con ContadorMx pero con **4x mas contenido**.
Significativamente menor que IMECAF y Tec de Monterrey sin sacrificar profundidad.

### 5.3 Diferenciador de precio

Lo que justifica $5,990 vs $1,149 (ContadorMx):
- 42-50 hrs vs 10 hrs de contenido
- Power Query + Power Pivot + VBA (ellos no lo cubren)
- Herramientas fiscales mexicanas reales (ISR, IVA, DIOT, conciliacion)
- 30+ ejercicios con datos realistas generados programaticamente
- 4 sesiones en vivo de Q&A
- Proyecto integrador con retroalimentacion
- Constancia de participacion

---

## 6. Formato y Duracion

| Aspecto | Detalle |
|---|---|
| **Horas de video** | 42-50 horas |
| **Duracion del programa** | 8 semanas (1 modulo por semana) |
| **Sesiones en vivo** | 4 sesiones de Q&A (cada 2 semanas, 1.5 hrs c/u) |
| **Ejercicios descargables** | 30+ archivos Excel |
| **Proyecto final** | Caso integrador con retroalimentacion |
| **Acceso al contenido** | 12 meses desde la compra |
| **Soporte** | Grupo privado (Telegram o comunidad en plataforma) |

---

## 7. Objetivos de Aprendizaje

Al finalizar, el alumno sera capaz de:

1. **Automatizar la importacion** de datos fiscales (XMLs, layouts bancarios, TXTs
   del SAT) usando Power Query, eliminando la captura manual.
2. **Construir modelos de datos** relacionales con Power Pivot para analisis contable
   multi-tabla sin duplicar informacion.
3. **Desarrollar macros VBA** para automatizar tareas repetitivas del ciclo contable
   (papeles de trabajo, formatos, reportes).
4. **Calcular impuestos** (ISR, IVA, factores de actualizacion) con hojas de trabajo
   vinculadas y auditables.
5. **Disenar dashboards financieros** interactivos para presentacion ejecutiva a
   clientes y directivos.
6. **Aplicar herramientas de IA** (Copilot, ChatGPT, Claude) de forma critica y
   validada, sin sustituir el criterio profesional.
7. **Proteger y distribuir** libros de trabajo con estandares profesionales de
   un despacho contable.

---

## 8. Ruta de Monetizacion

```
Curso Basico ($799-$1,199)             Curso Premium ($5,990)
"Domina lo esencial"                   "Automatiza tu despacho"
         |                                      |
         |---- funnel / upsell --------------->  |
         |                                      |
         v                                      v
   Lead magnet para                     Mentoria / Consultoria
   filtrar alumnos serios               ($2,000-$5,000 por sesion)
```

### Estrategia de lanzamiento

1. **Preventa con early-bird** ($4,490) a la base de alumnos del curso basico
2. **Webinar gratuito** de 1 hora mostrando un caso real (Power Query + XMLs)
   como gancho para el premium
3. **Testimonios** de alumnos del basico que quieren mas profundidad
4. **Garantia de 15 dias** — si no le sirve, devolucion completa

---

## 9. Esfuerzo de Construccion

| Componente | Descripcion | Esfuerzo |
|---|---|---|
| **Generadores Python** | 30+ scripts nuevos (estructura identica a los actuales) | Principal |
| **Contenido de video** | 42-50 hrs de contenido terminado (grabar + editar) | Mayor tiempo |
| **Sesiones en vivo** | 4 sesiones de Q&A (preparacion + ejecucion) | Recurrente |
| **Plataforma** | Teachable, Hotmart, o Thinkific (configuracion) | Una vez |
| **Pagina de ventas** | Copy basado en el temario + testimonios | Una vez |
| **Comunidad** | Grupo de Telegram o foro en la plataforma | Continuo |

### Ventaja competitiva tecnica

Los generadores Python del proyecto actual (`scripts/build_all.py`) son la
**ventaja competitiva real**. Mientras otros instructores crean ejercicios
manualmente (y por eso tienen pocos, 3-5 archivos), este sistema puede generar
30+ archivos con datos realistas, consistentes y actualizados a las tarifas
fiscales vigentes con un solo comando:

```bash
python3 -m scripts.build_all
```

Esto permite:
- Actualizar todos los ejercicios cuando cambien tarifas (2027, 2028...)
- Generar variantes para diferentes niveles o industrias
- Mantener consistencia entre todos los materiales
- Escalar sin esfuerzo proporcional

---

## 10. Proximos Pasos

1. Validar el temario con 3-5 contadores que hayan tomado el curso basico
2. Decidir plataforma de distribucion (Hotmart tiene mayor penetracion en LATAM)
3. Crear los generadores del Modulo P1 (Power Query) como piloto
4. Grabar un modulo piloto y medir la recepcion
5. Lanzar preventa con early-bird antes de terminar todo el contenido
