# Referencia Rápida — Módulo 4

**El Dashboard Inteligente y Entrega Profesional**

*Diseño, Slicers, Protección y Distribución*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## 1. Principios de Diseño de Dashboards

Un dashboard efectivo comunica información clave de un vistazo. Sigue estos principios para crear paneles profesionales en Excel.


### Regla de los 5 segundos

El mensaje principal de tu dashboard debe entenderse en 5 segundos. Si necesitas más tiempo, hay demasiada información o está mal organizada.


### Jerarquía visual

- ARRIBA: KPIs (4-6 números clave en cuadros de colores)
- CENTRO: Gráficos principales (barras, líneas, dona)
- IZQUIERDA: Segmentadores de datos (filtros interactivos)
- ABAJO: Tabla de detalle o gráficos secundarios

### Paleta de colores

| Color | Código Hex | Uso |
| --- | --- | --- |
| Azul | #2563EB | Color principal, encabezados, títulos |
| Verde | #10B981 | Positivo, vigente, aprobado |
| Rojo | #EF4444 | Alerta, vencido, rechazado |
| Amarillo | #F59E0B | Precaución, por vencer, en proceso |
| Gris claro | #F8FAFC | Fondo del dashboard |
| Blanco | #FFFFFF | Fondo de cuadros de KPI |


### Tipos de gráfico recomendados

- Barras agrupadas: comparar categorías (ej. sueldo por puesto)
- Líneas: mostrar tendencias en el tiempo (ej. nómina mensual)
- Dona/Pie: proporción de un total (ej. % ISR vs % neto)
- EVITAR: gráficos 3D, demasiados colores, ejes innecesarios

### Limpieza visual

- Oculta líneas de cuadrícula: Vista > desmarcar 'Líneas de cuadrícula'
- Usa bordes solo donde agreguen claridad (no en todas las celdas)
- Alinea los elementos visual y consistentemente
- Usa una sola familia tipográfica (Calibri recomendado)

---

## 2. Segmentadores de Datos (Slicers) — Paso a Paso

Los segmentadores son filtros visuales para Tablas Dinámicas. Permiten crear dashboards interactivos sin macros ni VBA.


### Paso 1: Preparar tus datos

- Tus datos deben estar en formato de Tabla (Ctrl+T) o Tabla Dinámica
- Cada columna debe tener un encabezado único y descriptivo
- No debe haber filas ni columnas vacías dentro de los datos

### Paso 2: Crear la Tabla Dinámica

- Selecciona cualquier celda de tu tabla de datos
- Menú: Insertar > Tabla Dinámica > Hoja nueva o existente
- Arrastra campos: filas (Periodo), valores (Suma de Sueldo, Suma de ISR)

### Paso 3: Insertar el segmentador

- Haz clic en cualquier celda de tu Tabla Dinámica
- Menú: Insertar > Segmentación de datos
- Selecciona los campos que quieres filtrar: Periodo, Puesto, Empleado
- Haz clic en Aceptar — aparecen los botones de filtro

### Paso 4: Vincular a múltiples TDs

- Clic derecho sobre el segmentador > Conexiones de informe...
- Marca TODAS las Tablas Dinámicas que deben responder al filtro
- Ahora un clic filtra todas las TDs y sus gráficos asociados

### Paso 5: Personalizar el segmentador

- Selecciona el slicer > pestaña Segmentación de datos (cinta)
- Cambia el estilo visual, número de columnas y tamaño
- Consejo: usa 2-3 columnas para que ocupe menos espacio horizontal

### Atajos útiles

| Acción | Atajo / Método |
| --- | --- |
| Seleccionar múltiples ítems | Ctrl + clic en cada botón |
| Limpiar filtro del slicer | Ícono de embudo con X (esquina del slicer) |
| Mover slicer | Arrastrar con el mouse |
| Redimensionar | Arrastrar las esquinas |
| Eliminar slicer | Seleccionar + tecla Suprimir |

---

## 3. Protección de Celdas — Guía Rápida

Proteger tu archivo asegura que el usuario final interactúe correctamente con el dashboard sin romper fórmulas ni la estructura.


### Flujo de protección en 4 pasos

| Paso | Acción | Dónde |
| --- | --- | --- |
| 1 | Selecciona las celdas de INPUT (donde el usuario escribe) | Celdas amarillas |
| 2 | Clic derecho > Formato > Protección > desmarcar 'Bloqueada' | Cada celda de input |
| 3 | Revisar > Proteger hoja > contraseña > permitir filtros/TDs | Cada hoja |
| 4 | Revisar > Proteger libro > contraseña (estructura) | Una vez por archivo |


### Qué proteger y qué no

| Elemento | ¿Bloquear? | Razón |
| --- | --- | --- |
| Celdas con fórmulas | SÍ | Evitar que borren o modifiquen cálculos |
| Celdas de input del usuario | NO | El usuario necesita ingresar datos |
| Encabezados de tabla | SÍ | Mantener estructura |
| Celdas de KPI con SUBTOTAL | SÍ | Son fórmulas críticas |
| Nombre de hojas | SÍ (libro) | Evitar renombrar o eliminar hojas |


### Niveles de seguridad en Excel

| Nivel | Método | Seguridad |
| --- | --- | --- |
| Básico | Proteger hoja (sin contraseña) | Mínima — cualquiera desprotege |
| Medio | Proteger hoja + libro (con contraseña) | Moderada — herramientas para romper |
| Alto | Contraseña de apertura de archivo (AES) | Alta — cifrado real |
| Máximo | Exportar como PDF (no editable) | Máxima — sin datos editables |

---

## 4. Checklist de Distribución Profesional

Usa esta lista antes de enviar cualquier archivo a clientes, jefes o colegas.


### Antes de compartir

- [ ] Las fórmulas calculan correctamente con datos de prueba
- [ ] Los gráficos se actualizan al cambiar filtros/slicers
- [ ] Las celdas de input están desbloqueadas y resaltadas en amarillo
- [ ] Las celdas de fórmula están bloqueadas
- [ ] Las hojas están protegidas con contraseña
- [ ] La estructura del libro está protegida
- [ ] No hay datos personales o confidenciales expuestos
- [ ] El nombre del archivo sigue la convención profesional

### Formato del nombre de archivo

```
[Empresa]_[TipoDocumento]_[Periodo]_[Version].[ext]
```


- Ejemplo: GrupoTorres_Nomina_2026_Enero_v1.xlsx
- Ejemplo: CNO850315_ISR_Anual_2025_Final.pdf
- Evita espacios, caracteres especiales y acentos en nombres de archivo

### Qué formato usar

| Situación | Formato | Razón |
| --- | --- | --- |
| Reporte final a cliente | PDF | No editable, aspecto profesional |
| Declaración o complemento SAT | PDF + XML | Formato oficial |
| Plantilla que el cliente debe llenar | Excel protegido | Mantiene fórmulas |
| Dashboard interactivo para equipo | Excel protegido | Slicers y TDs activos |
| Respaldo/archivo maestro | Excel sin proteger | Facilita edición futura |


### Después de compartir

- [ ] Confirma que el destinatario puede abrir el archivo
- [ ] Verifica que la versión de Excel del destinatario es compatible
- [ ] Guarda tu copia maestra sin proteger (sufijo _master.xlsx)
- [ ] Documenta la contraseña de protección en lugar seguro

### Archivos de este módulo

- 09_Layout_Dashboard_Contable.xlsx — Template de layout para dashboard
- 10_Dashboard_Final_Integrado.xlsx — Datos + Calculadora + Dashboard
- 11_Guia_Proteccion_y_Seguridad.md — Guía detallada de protección
- Referencia_Modulo_4.md — Este documento