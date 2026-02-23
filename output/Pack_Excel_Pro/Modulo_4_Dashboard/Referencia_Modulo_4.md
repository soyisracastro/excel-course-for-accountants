# Referencia Rapida -- Modulo 4

**El Dashboard Inteligente y Entrega Profesional**

*Diseno, Slicers, Proteccion y Distribucion*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## 1. Principios de Diseno de Dashboards

Un dashboard efectivo comunica informacion clave de un vistazo. Sigue estos principios para crear paneles profesionales en Excel.


### Regla de los 5 segundos

El mensaje principal de tu dashboard debe entenderse en 5 segundos. Si necesitas mas tiempo, hay demasiada informacion o esta mal organizada.


### Jerarquia visual

- ARRIBA: KPIs (4-6 numeros clave en cuadros de colores)
- CENTRO: Graficos principales (barras, lineas, dona)
- IZQUIERDA: Segmentadores de datos (filtros interactivos)
- ABAJO: Tabla de detalle o graficos secundarios

### Paleta de colores

| Color | Codigo Hex | Uso |
| --- | --- | --- |
| Azul | #2563EB | Color principal, encabezados, titulos |
| Verde | #10B981 | Positivo, vigente, aprobado |
| Rojo | #EF4444 | Alerta, vencido, rechazado |
| Amarillo | #F59E0B | Precaucion, por vencer, en proceso |
| Gris claro | #F8FAFC | Fondo del dashboard |
| Blanco | #FFFFFF | Fondo de cuadros de KPI |


### Tipos de grafico recomendados

- Barras agrupadas: comparar categorias (ej. sueldo por puesto)
- Lineas: mostrar tendencias en el tiempo (ej. nomina mensual)
- Dona/Pie: proporcion de un total (ej. % ISR vs % neto)
- EVITAR: graficos 3D, demasiados colores, ejes innecesarios

### Limpieza visual

- Oculta lineas de cuadricula: Vista > desmarcar 'Lineas de cuadricula'
- Usa bordes solo donde agreguen claridad (no en todas las celdas)
- Alinea los elementos visual y consistentemente
- Usa una sola familia tipografica (Calibri recomendado)

---

## 2. Segmentadores de Datos (Slicers) -- Paso a Paso

Los segmentadores son filtros visuales para Tablas Dinamicas. Permiten crear dashboards interactivos sin macros ni VBA.


### Paso 1: Preparar tus datos

- Tus datos deben estar en formato de Tabla (Ctrl+T) o Tabla Dinamica
- Cada columna debe tener un encabezado unico y descriptivo
- No debe haber filas ni columnas vacias dentro de los datos

### Paso 2: Crear la Tabla Dinamica

- Selecciona cualquier celda de tu tabla de datos
- Menu: Insertar > Tabla Dinamica > Hoja nueva o existente
- Arrastra campos: filas (Periodo), valores (Suma de Sueldo, Suma de ISR)

### Paso 3: Insertar el segmentador

- Haz clic en cualquier celda de tu Tabla Dinamica
- Menu: Insertar > Segmentacion de datos
- Selecciona los campos que quieres filtrar: Periodo, Puesto, Empleado
- Haz clic en Aceptar -- aparecen los botones de filtro

### Paso 4: Vincular a multiples TDs

- Clic derecho sobre el segmentador > Conexiones de informe...
- Marca TODAS las Tablas Dinamicas que deben responder al filtro
- Ahora un clic filtra todas las TDs y sus graficos asociados

### Paso 5: Personalizar el segmentador

- Selecciona el slicer > pestana Segmentacion de datos (cinta)
- Cambia el estilo visual, numero de columnas y tamano
- Consejo: usa 2-3 columnas para que ocupe menos espacio horizontal

### Atajos utiles

| Accion | Atajo / Metodo |
| --- | --- |
| Seleccionar multiples items | Ctrl + clic en cada boton |
| Limpiar filtro del slicer | Icono de embudo con X (esquina del slicer) |
| Mover slicer | Arrastrar con el mouse |
| Redimensionar | Arrastrar las esquinas |
| Eliminar slicer | Seleccionar + tecla Suprimir |

---

## 3. Proteccion de Celdas -- Guia Rapida

Proteger tu archivo asegura que el usuario final interactue correctamente con el dashboard sin romper formulas ni la estructura.


### Flujo de proteccion en 4 pasos

| Paso | Accion | Donde |
| --- | --- | --- |
| 1 | Selecciona las celdas de INPUT (donde el usuario escribe) | Celdas amarillas |
| 2 | Clic derecho > Formato > Proteccion > desmarcar 'Bloqueada' | Cada celda de input |
| 3 | Revisar > Proteger hoja > contrasena > permitir filtros/TDs | Cada hoja |
| 4 | Revisar > Proteger libro > contrasena (estructura) | Una vez por archivo |


### Que proteger y que no

| Elemento | Bloquear? | Razon |
| --- | --- | --- |
| Celdas con formulas | SI | Evitar que borren o modifiquen calculos |
| Celdas de input del usuario | NO | El usuario necesita ingresar datos |
| Encabezados de tabla | SI | Mantener estructura |
| Celdas de KPI con SUBTOTAL | SI | Son formulas criticas |
| Nombre de hojas | SI (libro) | Evitar renombrar o eliminar hojas |


### Niveles de seguridad en Excel

| Nivel | Metodo | Seguridad |
| --- | --- | --- |
| Basico | Proteger hoja (sin contrasena) | Minima -- cualquiera desprotege |
| Medio | Proteger hoja + libro (con contrasena) | Moderada -- herramientas para romper |
| Alto | Contrasena de apertura de archivo (AES) | Alta -- cifrado real |
| Maximo | Exportar como PDF (no editable) | Maxima -- sin datos editables |

---

## 4. Checklist de Distribucion Profesional

Usa esta lista antes de enviar cualquier archivo a clientes, jefes o colegas.


### Antes de compartir

- [ ] Las formulas calculan correctamente con datos de prueba
- [ ] Los graficos se actualizan al cambiar filtros/slicers
- [ ] Las celdas de input estan desbloqueadas y resaltadas en amarillo
- [ ] Las celdas de formula estan bloqueadas
- [ ] Las hojas estan protegidas con contrasena
- [ ] La estructura del libro esta protegida
- [ ] No hay datos personales o confidenciales expuestos
- [ ] El nombre del archivo sigue la convencion profesional

### Formato del nombre de archivo

```
[Empresa]_[TipoDocumento]_[Periodo]_[Version].[ext]
```


- Ejemplo: GrupoTorres_Nomina_2026_Enero_v1.xlsx
- Ejemplo: CNO850315_ISR_Anual_2025_Final.pdf
- Evita espacios, caracteres especiales y acentos en nombres de archivo

### Que formato usar

| Situacion | Formato | Razon |
| --- | --- | --- |
| Reporte final a cliente | PDF | No editable, aspecto profesional |
| Declaracion o complemento SAT | PDF + XML | Formato oficial |
| Plantilla que el cliente debe llenar | Excel protegido | Mantiene formulas |
| Dashboard interactivo para equipo | Excel protegido | Slicers y TDs activos |
| Respaldo/archivo maestro | Excel sin proteger | Facilita edicion futura |


### Despues de compartir

- [ ] Confirma que el destinatario puede abrir el archivo
- [ ] Verifica que la version de Excel del destinatario es compatible
- [ ] Guarda tu copia maestra sin proteger (sufijo _master.xlsx)
- [ ] Documenta la contrasena de proteccion en lugar seguro

### Archivos de este modulo

- 09_Layout_Dashboard_Contable.xlsx -- Template de layout para dashboard
- 10_Dashboard_Final_Integrado.xlsx -- Datos + Calculadora + Dashboard
- 11_Guia_Proteccion_y_Seguridad.md -- Guia detallada de proteccion
- Referencia_Modulo_4.md -- Este documento