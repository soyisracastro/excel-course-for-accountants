# Guia de Proteccion y Seguridad en Excel

**Modulo 4 -- El Dashboard Inteligente y Entrega Profesional**

*Protege, comparte y entrega archivos profesionales*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## 1. Proteccion de Celdas

En Excel, TODAS las celdas vienen bloqueadas por defecto, pero el bloqueo solo se activa cuando proteges la hoja. La estrategia correcta es:


### 1.1 Desbloquear celdas de entrada (inputs)

- Selecciona las celdas donde el usuario debe escribir datos (ej. B4 en Calculadora)
- Clic derecho > Formato de celdas > pestana Proteccion
- Desmarca la casilla 'Bloqueada'
- Opcionalmente marca 'Oculta' para ocultar formulas en la barra de formulas

### 1.2 Mantener bloqueadas las celdas de formula

- Las celdas con formulas deben mantener la casilla 'Bloqueada' activada
- Esto evita que el usuario modifique o borre accidentalmente las formulas
- Ejemplo: celdas B5:B10 en la Calculadora ISR contienen BUSCARV

### 1.3 Identificar visualmente las celdas editables

- Usa un color de fondo distinto para celdas de entrada (ej. amarillo #F59E0B)
- Agrega una nota o etiqueta: 'Ingresa tu dato aqui'
- Esto le indica al usuario donde puede escribir sin confusion

## 2. Proteccion de Hojas

Proteger una hoja impide que los usuarios modifiquen celdas bloqueadas, pero permite la edicion en celdas desbloqueadas.


### Pasos para proteger una hoja

- Menu: Revisar > Proteger hoja
- Establece una contrasena (opcional pero recomendado)
- Selecciona los permisos que deseas otorgar:

| Permiso | Descripcion | Recomendado |
| --- | --- | --- |
| Seleccionar celdas bloqueadas | El usuario puede ver pero no editar | Si |
| Seleccionar celdas desbloqueadas | El usuario puede editar celdas de input | Si |
| Formato de celdas | Permite cambiar formato visual | No |
| Insertar filas | Permite agregar filas nuevas | Segun caso |
| Eliminar filas | Permite borrar filas | No |
| Ordenar | Permite ordenar datos | Si |
| Usar Autofiltro | Permite filtrar datos | Si |
| Usar tablas dinamicas | Permite interactuar con TDs | Si |


### Contrasenas seguras

- Usa al menos 8 caracteres con mayusculas, numeros y simbolos
- IMPORTANTE: La proteccion de hoja NO es cifrado fuerte
- No confies en ella para datos altamente confidenciales
- Para datos sensibles, usa contrasena de apertura de archivo (ver seccion 3)

---

## 3. Proteccion de Libro

La proteccion de libro evita cambios estructurales: agregar, eliminar, renombrar u ocultar hojas.


### 3.1 Proteger estructura del libro

- Menu: Revisar > Proteger libro
- Marca 'Estructura' para evitar agregar/eliminar hojas
- Establece contrasena

### 3.2 Contrasena de apertura de archivo

- Menu: Archivo > Guardar como > Herramientas > Opciones generales
- 'Contrasena de apertura': el archivo no se abre sin ella (cifrado)
- 'Contrasena de escritura': permite abrir en solo lectura sin contrasena
- NOTA: La contrasena de apertura usa cifrado AES -- es segura

### 3.3 Marcar como final

- Menu: Archivo > Informacion > Proteger libro > Marcar como final
- Pone el archivo en modo solo lectura (el usuario puede desactivarlo)
- Es una recomendacion, no una proteccion fuerte

## 4. Distribucion: PDF vs Excel Protegido

La forma de compartir depende de lo que necesitas que haga el destinatario.


| Criterio | PDF | Excel Protegido |
| --- | --- | --- |
| El usuario necesita editar | No | Si (celdas de input) |
| Mantiene formulas activas | No | Si |
| Interactuar con slicers/TDs | No | Si |
| Seguridad del contenido | Alta (no editable) | Media (contrasena) |
| Tamano de archivo | Ligero | Normal |
| Requiere Excel instalado | No | Si |
| Ideal para | Reportes finales, firmas | Plantillas, calculadoras |


### Como guardar como PDF desde Excel

- Menu: Archivo > Guardar como > Tipo: PDF
- O bien: Archivo > Exportar > Crear documento PDF/XPS
- Selecciona las hojas a incluir antes de guardar
- Tip: Configura el area de impresion antes (Diseno de pagina > Area de impresion)

---

## 5. Convenciones de Nombres Profesionales

Un nombre de archivo profesional facilita la organizacion, la busqueda y transmite seriedad al cliente.


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

- NO uses espacios -- usa guion bajo (_) o guion medio (-)
- Incluye el periodo (mes, trimestre, anio)
- Incluye version (v1, v2, Final, Borrador)
- Usa el RFC cuando el archivo es para un cliente especifico
- Manten nombres cortos pero descriptivos (max ~50 caracteres)

## 6. Checklist de Entrega Profesional

Antes de enviar un archivo a tu cliente, jefe o colega, verifica cada punto de esta lista.


| # | Verificacion | Hecho |
| --- | --- | --- |
| 1 | Las formulas calculan correctamente (verifica con datos de prueba) | [ ] |
| 2 | Las celdas de input estan DESBLOQUEADAS y resaltadas en amarillo | [ ] |
| 3 | Las celdas de formula estan BLOQUEADAS | [ ] |
| 4 | La hoja esta protegida con contrasena | [ ] |
| 5 | La estructura del libro esta protegida | [ ] |
| 6 | No hay datos personales o confidenciales expuestos | [ ] |
| 7 | Los graficos se ven correctamente al cambiar filtros | [ ] |
| 8 | Los segmentadores (slicers) estan conectados a todas las TDs | [ ] |
| 9 | El nombre del archivo sigue la convencion profesional | [ ] |
| 10 | Se genero una copia PDF del reporte final | [ ] |
| 11 | Se verifico en otra computadora o version de Excel | [ ] |
| 12 | Se incluyen instrucciones de uso (hoja o documento adjunto) | [ ] |


CONSEJO: Guarda una copia sin proteger como archivo maestro (ej. '_master.xlsx') y distribuye siempre la version protegida.
