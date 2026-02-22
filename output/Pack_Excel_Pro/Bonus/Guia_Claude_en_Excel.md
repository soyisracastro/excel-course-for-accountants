# Guia Claude en Excel

**Bonus - Material Complementario**

*Instalacion, prompts contables y mejores practicas*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## Guia de Instalacion del Add-in Claude para Excel

Desde enero de 2026, Claude esta disponible como add-in oficial para suscriptores de Microsoft 365 Pro. La instalacion es directa desde el Marketplace de Microsoft.


### Requisitos

- Microsoft 365 Pro (suscripcion activa)
- Excel para Windows (version 2312 o superior) o Excel para Web
- Conexion a internet para la instalacion y uso del add-in
- Cuenta de Microsoft vinculada a la suscripcion M365

### Pasos de Instalacion

- **Paso 1:** Abrir Excel y crear o abrir cualquier libro
- **Paso 2:** Ir a la pestana Insertar en la cinta de opciones
- **Paso 3:** Clic en 'Obtener complementos' (o 'Get Add-ins')
- **Paso 4:** En el cuadro de busqueda, escribir 'Claude for Excel'
- **Paso 5:** Seleccionar el add-in oficial de Anthropic y clic en 'Agregar'
- **Paso 6:** Aceptar los permisos solicitados
- **Paso 7:** El panel de Claude aparecera en la pestana Inicio

### Verificacion

Para verificar que la instalacion fue exitosa, abrir el panel lateral de Claude y escribir una pregunta simple como 'Hola, que puedes hacer con esta hoja?'. Claude debe responder con una descripcion de sus capacidades.

---

## 10 Prompts Contables para Claude en Excel

Estos prompts estan disenados especificamente para tareas contables y administrativas. Copia y pega directamente en el panel de Claude.


### 1. Analisis de nomina

```
Analiza esta tabla de nomina. Identifica los empleados cuyo sueldo esta por encima del promedio de su departamento. Muestra el nombre, departamento, sueldo, y la diferencia con el promedio.
```


### 2. Deteccion de anomalias en gastos

```
Revisa esta tabla de gastos mensuales. Identifica movimientos que se desvien mas del 30% del promedio historico de su categoria. Senala cuales podrian requerir revision.
```


### 3. Generacion de formula BUSCARV + SIERROR

```
Necesito una formula que busque el RFC de la columna A en la tabla Proveedores y traiga el nombre de la columna 3. Si no lo encuentra, debe mostrar 'No encontrado'. Explicame paso a paso.
```


### 4. Explicacion de formula heredada

```
Explicame paso a paso que hace esta formula, como si yo fuera contador y no programador. Descompone cada funcion anidada: [PEGAR FORMULA AQUI]
```


### 5. Conciliacion bancaria

```
Tengo dos tablas: Estado_Banco y Registro_Contable. Ambas tienen columnas Fecha, Referencia, y Monto. Identifica las diferencias: movimientos que estan en banco pero no en contabilidad, y viceversa.
```


### 6. Proyeccion de flujo de efectivo

```
Con base en los ingresos y egresos de los ultimos 6 meses en esta tabla, proyecta el flujo de efectivo para los proximos 3 meses. Usa tendencia lineal y senala meses con posible deficit.
```


### 7. Clasificacion de cuentas contables

```
Clasifica estos movimientos en las categorias del catalogo de cuentas: Activo, Pasivo, Capital, Ingreso, Costo, Gasto. Agrega la clasificacion en una columna nueva.
```


### 8. Calculo de ISR con tarifas

```
Usando la tarifa del Art. 96 LISR vigente para 2026, calcula la retencion de ISR mensual para cada empleado de esta tabla. Muestra el desglose: base gravable, limite inferior, excedente, impuesto marginal, cuota fija, y total ISR.
```


### 9. Resumen ejecutivo de ventas

```
Genera un resumen ejecutivo de esta tabla de ventas: total por producto, por region, por vendedor. Identifica el top 5 de productos y el bottom 5. Sugiere acciones basadas en los datos.
```


### 10. Auditoria de formulas

```
Revisa todas las formulas de esta hoja. Identifica: celdas con errores (#REF!, #N/A, #VALOR!), referencias circulares, formulas que podrian simplificarse, y celdas con valores hardcodeados donde deberia haber formulas.
```


---

## Claude vs Copilot: Tabla Comparativa

Ambas herramientas son complementarias. Esta tabla resume las diferencias principales para ayudarte a decidir cual usar segun la tarea.


| Caracteristica | Claude (Anthropic) | Copilot (Microsoft) |
| --- | --- | --- |
| Integracion en Excel | Add-in desde Marketplace | Nativo en M365 |
| Analisis profundo de datos | Excelente | Bueno |
| Generacion de formulas | Excelente (con explicacion) | Muy bueno |
| Automatizacion rapida | Buena (via prompts) | Excelente (nativo) |
| Generacion de graficos | Sugiere configuracion | Crea directamente |
| Razonamiento complejo | Superior | Bueno |
| Codigo VBA | Genera y explica | Genera |
| Lenguaje natural en espaniol | Excelente | Muy bueno |
| Privacidad de datos | No almacena datos | Segun plan M365 |
| Costo | Incluido en M365 Pro | Incluido en M365 Pro/Copilot |
| Mejor para | Analisis, razonamiento, auditoria | Tareas rapidas, automatizacion |
| Conectores externos (MCP) | Si, via protocolo MCP | Si, via Microsoft Graph |


**Recomendacion:** Usa Copilot para tareas rapidas de automatizacion y creacion de graficos. Usa Claude para analisis profundo, razonamiento sobre datos, auditoria de formulas, y generacion de codigo VBA con explicaciones detalladas. Juntos cubren el 95% de las necesidades.

---

## Privacidad y Consideraciones de Datos

Al usar inteligencia artificial con datos empresariales, es fundamental entender como se manejan los datos.


### Politica de datos de Claude en Excel

- Claude **no almacena** los datos de las hojas que procesa en las sesiones del add-in
- Los datos se envian a los servidores de Anthropic para procesamiento y se descartan despues de generar la respuesta
- Anthropic **no usa datos de clientes empresariales** para entrenar sus modelos (politica vigente desde 2024)
- Para suscriptores M365 Pro, los datos transitan por la infraestructura de Microsoft Azure

### Recomendaciones de seguridad

- Consultar con el area de TI antes de usar IA con datos confidenciales
- No enviar datos sensibles como contrasenas, numeros de tarjeta, o informacion personal identificable (PII) a menos que la politica organizacional lo permita
- Para datos altamente confidenciales, considerar Claude en modo empresarial con politicas de retencion personalizadas
- Siempre validar las respuestas de la IA antes de tomar decisiones basadas en ellas
- Documentar el uso de IA en procesos contables para cumplimiento normativo y auditoria
---

## MCP (Model Context Protocol): Vision General

MCP es un protocolo abierto creado por Anthropic que permite a Claude conectarse a fuentes de datos externas de forma segura y estandarizada.


### Que es MCP?

- MCP = Model Context Protocol (Protocolo de Contexto de Modelo)
- Permite que Claude acceda a datos fuera de la hoja de Excel: bases de datos, APIs, archivos en la nube
- Protocolo abierto y estandarizado: cualquier proveedor puede crear conectores
- Los conectores se configuran a nivel organizacional por el area de TI

### Casos de uso contable

- **Base de datos SQL:** Claude consulta directamente el sistema contable y trae datos a Excel para analisis
- **SharePoint/OneDrive:** Acceso a archivos historicos para comparaciones y consolidaciones
- **APIs externas:** Conexion a servicios del SAT, tipo de cambio del Banco de Mexico, o INPC actualizado
- **Correo y calendario:** Consultar fechas limite de declaraciones y recordatorios de cierre

### Como empezar con MCP

- Visitar modelcontextprotocol.io para documentacion completa
- Coordinar con el area de TI para configurar conectores organizacionales
- Empezar con conectores de solo lectura (consulta) antes de habilitar escritura
- Claude Code (terminal) soporta MCP de forma nativa para automatizacion avanzada
---

## Recursos y Enlaces

- **Claude para Excel:** Marketplace de Microsoft 365
- **Documentacion Claude:** docs.anthropic.com
- **Claude Code:** npm install -g @anthropic-ai/claude-code
- **MCP Protocol:** modelcontextprotocol.io
- **Anthropic:** anthropic.com
- **Comunidad del curso:** todoconta.com

**Nota final:** La inteligencia artificial es una herramienta que amplifica la productividad del profesional contable. El criterio humano, la etica profesional, y el conocimiento normativo siguen siendo insustituibles. Usen la IA como su segundo cerebro, pero nunca dejen de ser la brujula.
