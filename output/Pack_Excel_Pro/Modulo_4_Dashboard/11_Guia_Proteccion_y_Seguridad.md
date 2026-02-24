# Guía de Protección y Seguridad en Excel

**Módulo 4 — El Dashboard Inteligente y Entrega Profesional**

*Protege, comparte y entrega archivos profesionales*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## 1. Protección de Celdas

En Excel, TODAS las celdas vienen bloqueadas por defecto, pero el bloqueo solo se activa cuando proteges la hoja. La estrategia correcta es:


### 1.1 Desbloquear celdas de entrada (inputs)

- Selecciona las celdas donde el usuario debe escribir datos (ej. B4 en Calculadora)
- Clic derecho > Formato de celdas > pestaña Protección
- Desmarca la casilla 'Bloqueada'
- Opcionalmente marca 'Oculta' para ocultar fórmulas en la barra de fórmulas

### 1.2 Mantener bloqueadas las celdas de fórmula

- Las celdas con fórmulas deben mantener la casilla 'Bloqueada' activada
- Esto evita que el usuario modifique o borre accidentalmente las fórmulas
- Ejemplo: celdas B5:B10 en la Calculadora ISR contienen BUSCARV

### 1.3 Identificar visualmente las celdas editables

- Usa un color de fondo distinto para celdas de entrada (ej. amarillo #F59E0B)
- Agrega una nota o etiqueta: 'Ingresa tu dato aquí'
- Esto le indica al usuario dónde puede escribir sin confusión

## 2. Protección de Hojas

Proteger una hoja impide que los usuarios modifiquen celdas bloqueadas, pero permite la edición en celdas desbloqueadas.


### Pasos para proteger una hoja

- Menú: Revisar > Proteger hoja
- Establece una contraseña (opcional pero recomendado)
- Selecciona los permisos que deseas otorgar:

| Permiso | Descripción | Recomendado |
| --- | --- | --- |
| Seleccionar celdas bloqueadas | El usuario puede ver pero no editar | Sí |
| Seleccionar celdas desbloqueadas | El usuario puede editar celdas de input | Sí |
| Formato de celdas | Permite cambiar formato visual | No |
| Insertar filas | Permite agregar filas nuevas | Según caso |
| Eliminar filas | Permite borrar filas | No |
| Ordenar | Permite ordenar datos | Sí |
| Usar Autofiltro | Permite filtrar datos | Sí |
| Usar tablas dinámicas | Permite interactuar con TDs | Sí |


### Contraseñas seguras

- Usa al menos 8 caracteres con mayúsculas, números y símbolos
- IMPORTANTE: La protección de hoja NO es cifrado fuerte
- No confíes en ella para datos altamente confidenciales
- Para datos sensibles, usa contraseña de apertura de archivo (ver sección 3)

---

## 3. Protección de Libro

La protección de libro evita cambios estructurales: agregar, eliminar, renombrar u ocultar hojas.


### 3.1 Proteger estructura del libro

- Menú: Revisar > Proteger libro
- Marca 'Estructura' para evitar agregar/eliminar hojas
- Establece contraseña

### 3.2 Contraseña de apertura de archivo

- Menú: Archivo > Guardar como > Herramientas > Opciones generales
- 'Contraseña de apertura': el archivo no se abre sin ella (cifrado)
- 'Contraseña de escritura': permite abrir en solo lectura sin contraseña
- NOTA: La contraseña de apertura usa cifrado AES — es segura

### 3.3 Marcar como final

- Menú: Archivo > Información > Proteger libro > Marcar como final
- Pone el archivo en modo solo lectura (el usuario puede desactivarlo)
- Es una recomendación, no una protección fuerte

## 4. Distribución: PDF vs Excel Protegido

La forma de compartir depende de lo que necesitas que haga el destinatario.


| Criterio | PDF | Excel Protegido |
| --- | --- | --- |
| El usuario necesita editar | No | Sí (celdas de input) |
| Mantiene fórmulas activas | No | Sí |
| Interactuar con slicers/TDs | No | Sí |
| Seguridad del contenido | Alta (no editable) | Media (contraseña) |
| Tamaño de archivo | Ligero | Normal |
| Requiere Excel instalado | No | Sí |
| Ideal para | Reportes finales, firmas | Plantillas, calculadoras |


### Cómo guardar como PDF desde Excel

- Menú: Archivo > Guardar como > Tipo: PDF
- O bien: Archivo > Exportar > Crear documento PDF/XPS
- Selecciona las hojas a incluir antes de guardar
- Tip: Configura el área de impresión antes (Diseño de página > Área de impresión)

---

## 5. Convenciones de Nombres Profesionales

Un nombre de archivo profesional facilita la organización, la búsqueda y transmite seriedad al cliente.


### Formato recomendado

```
[Empresa]_[Tipo]_[Periodo]_[Version].[ext]
```


### Ejemplos

- GrupoTorres_Nomina_2026_Enero_v1.xlsx
- CNO850315XX1_ISR_Anual_2025_Final.pdf
- Dashboard_Fiscal_2026_Q1_v2.xlsx
- Reporte_Auditoria_2025_Borrador.pdf

### Reglas clave

- NO uses espacios — usa guion bajo (_) o guion medio (-)
- Incluye el periodo (mes, trimestre, año)
- Incluye versión (v1, v2, Final, Borrador)
- Usa el RFC cuando el archivo es para un cliente específico
- Mantén nombres cortos pero descriptivos (máx ~50 caracteres)

## 6. Checklist de Entrega Profesional

Antes de enviar un archivo a tu cliente, jefe o colega, verifica cada punto de esta lista.


| # | Verificación | Hecho |
| --- | --- | --- |
| 1 | Las fórmulas calculan correctamente (verifica con datos de prueba) | [ ] |
| 2 | Las celdas de input están DESBLOQUEADAS y resaltadas en amarillo | [ ] |
| 3 | Las celdas de fórmula están BLOQUEADAS | [ ] |
| 4 | La hoja está protegida con contraseña | [ ] |
| 5 | La estructura del libro está protegida | [ ] |
| 6 | No hay datos personales o confidenciales expuestos | [ ] |
| 7 | Los gráficos se ven correctamente al cambiar filtros | [ ] |
| 8 | Los segmentadores (slicers) están conectados a todas las TDs | [ ] |
| 9 | El nombre del archivo sigue la convención profesional | [ ] |
| 10 | Se generó una copia PDF del reporte final | [ ] |
| 11 | Se verificó en otra computadora o versión de Excel | [ ] |
| 12 | Se incluyen instrucciones de uso (hoja o documento adjunto) | [ ] |


CONSEJO: Guarda una copia sin proteger como archivo maestro (ej. '_master.xlsx') y distribuye siempre la versión protegida.
