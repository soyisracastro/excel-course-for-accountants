# Atajos de Teclado Excel

**Cheat Sheet - Bonus del Curso**

*Referencia rapida organizada por categoria*

Israel Castro | Excel para Contadores y Administrativos | 2026

---

## Atajos y trucos vistos en clase

Los siguientes atajos fueron demostrados en vivo durante las sesiones. Se incluye el minuto del video para que puedan regresar a ver el ejemplo en contexto.

### Modulo 1 — Logica Contable y Funciones

| Atajo / Truco | Para que lo usamos | Minuto |
| --- | --- | --- |
| `F4` (en formula) | Fijar referencia con $ (absoluta/relativa) en BUSCARV e ISR | 27:01, 51:23 |
| `F2` | Entrar a modo edicion para inspeccionar una formula | 27:52 |
| Doble clic en controlador de relleno | Copiar formula hacia abajo en toda la columna sin arrastrar | 28:57 |
| Cuadro de nombres (Name Box) | Nombrar un rango escribiendo el nombre y presionando Enter | 15:57 |
| Boton `fx` (Insertar funcion) | Abrir el catalogo de funciones por categoria | 18:58 |

### Modulo 2 — Tablas Dinamicas y Limpieza de Datos

| Atajo / Truco | Para que lo usamos | Minuto |
| --- | --- | --- |
| `Ctrl + H` | Buscar y reemplazar masivo: quitar $, puntos, guiones | 2:43 |
| `Ctrl + Espacio` | Seleccionar columna completa antes de limpiar | 6:34 |
| `Ctrl + B` | Abrir cuadro de Buscar (version en espanol) | 7:00 |
| `Ctrl + + (mas)` | Insertar columna auxiliar para limpieza | 9:03 |
| `Ctrl + C` | Copiar datos procesados | 9:50 |
| `Ctrl + Shift + V` | Pegar solo valores (sin formulas) | 10:24 |
| `Ctrl + - (menos)` | Eliminar columna auxiliar despues de usarla | 10:38 |
| `Shift + Av Pag` | Extender seleccion hacia abajo por pagina | 11:26 |
| `Ctrl + J` | Rellenar hacia abajo (Fill Down) | 19:43 |
| `Ctrl + Flecha` | Saltar al borde del rango de datos | 18:22 |
| `F2` + Enter | Forzar que Excel re-evalue una celda (ej: texto a fecha) | 21:28 |
| `Ctrl + H`: reemplazar `/` por `/` | Truco: fuerza reconocimiento de fechas en texto | 21:41 |
| `Ctrl + *` | Seleccionar toda la region de datos antes de quitar duplicados | 24:02 |
| Datos > Quitar duplicados | Eliminar registros repetidos por columnas seleccionadas | 24:15 |
| `Ctrl + T` | Convertir rango a tabla (habilita filtros, encabezados fijos, etc.) | 31:16 |
| Doble clic en controlador de relleno | Copiar formula hacia abajo (mismo truco del Modulo 1) | 26:37 |
| Diseno de tabla > Resumir con TD | Crear tabla dinamica directamente desde una tabla nombrada | 34:42 |
| Insertar > Segmentacion de datos | Agregar Slicers para filtrar visualmente la tabla dinamica | 37:54 |
| Analizar TD > Mostrar > Lista de campos | Reabrir el panel de campos si se cierra accidentalmente | 39:48 |

### Tips mencionados en clase

| Tip | Contexto | Modulo |
| --- | --- | --- |
| Usa TRUNCAR en vez de REDONDEAR para calculos fiscales | Factor de actualizacion (Art. 17-A CFF) se trunca a diezmilesimas | M1 |
| Las fechas son numeros en Excel | Puedes sumar/restar dias directamente: `=HOY() + 30` | M1 |
| HOY() no requiere argumentos | Es de las pocas funciones que van con parentesis vacios | M1 |
| El texto en formulas va entre comillas | Si no, Excel lo interpreta como nombre de rango | M1 |
| Excel no distingue mayusculas en funciones | `=suma()` y `=SUMA()` son equivalentes | M1 |
| Formato condicional con formula SI | Usa reglas personalizadas para crear semaforos de estatus | M1 |
| Triangulo verde = texto que parece numero | Usa el flag "Convertir a numero" o limpia con VALOR() | M2 |
| LARGO() para validar RFC | Verifica que el RFC tenga 12 (moral) o 13 (fisica) caracteres | M2 |
| MINUSCULA, MAYUSCULA, NOMPROPIO | Normalizar texto antes de analizar: nombres, proveedores, etc. | M2 |
| Nombra tus tablas inmediatamente | Evita confusion con Tabla1, Tabla2... cuando tengas varias | M2 |
| Aprende 1 atajo por semana | En un ano tendras 52 atajos memorizados sin esfuerzo | M2 |

---

## Referencia completa por categoria

### 1. Navegacion

| Atajo | Accion | Tip |
| --- | --- | --- |
| `Ctrl + Home` | Ir a celda A1 | *Inicio del libro* |
| `Ctrl + End` | Ir a ultima celda con datos | *Esquina inferior derecha usada* |
| `Ctrl + Flecha` | Saltar al borde del rango | *Funciona en las 4 direcciones* |
| `Ctrl + *` | Seleccionar region actual | *Equivale a Ctrl+Shift+8* |
| `Ctrl + G (o F5)` | Ir a... (cuadro de dialogo) | *Util para ir a celdas especificas* |
| `Ctrl + Page Up` | Hoja anterior | *Navegar entre hojas rapidamente* |
| `Ctrl + Page Down` | Hoja siguiente | *Navegar entre hojas rapidamente* |
| `Alt + Page Up` | Pantalla a la izquierda | *Desplazamiento horizontal* |
| `Alt + Page Down` | Pantalla a la derecha | *Desplazamiento horizontal* |
| `Ctrl + Tab` | Siguiente libro abierto | *Cambiar entre archivos Excel* |

### 2. Seleccion

| Atajo | Accion | Tip |
| --- | --- | --- |
| `Shift + Clic` | Seleccionar rango desde celda activa | *Mas rapido que arrastrar* |
| `Ctrl + Shift + End` | Seleccionar hasta ultima celda usada | *Ideal para rangos grandes* |
| `Ctrl + Shift + Home` | Seleccionar desde activa hasta A1 | *Selecciona todo arriba* |
| `Ctrl + Shift + Flecha` | Extender seleccion hasta borde | *Combina salto + seleccion* |
| `Ctrl + Space` | Seleccionar columna completa | *Toda la columna de la celda activa* |
| `Shift + Space` | Seleccionar fila completa | *Toda la fila de la celda activa* |
| `Ctrl + A` | Seleccionar todo (tabla o hoja) | *1er clic = tabla, 2do = toda la hoja* |
| `Ctrl + Shift + *` | Seleccionar region de datos actual | *Detecta automaticamente el rango* |
| `Alt + ;` | Seleccionar solo celdas visibles | *Ignora filas ocultas por filtro* |
| `Shift + Av Pag` | Extender seleccion hacia abajo por pagina | *Para seleccionar rangos grandes rapidamente* |

### 3. Edicion

| Atajo | Accion | Tip |
| --- | --- | --- |
| `F2` | Editar celda activa | *Entra en modo edicion sin borrar* |
| `Ctrl + Z` | Deshacer ultima accion | *Hasta 100 niveles de deshacer* |
| `Ctrl + Y` | Rehacer / Repetir ultima accion | *Tambien funciona como repetir* |
| `Ctrl + D` | Copiar celda de arriba hacia abajo | *Rellena con contenido de arriba* |
| `Ctrl + J` | Rellenar hacia abajo (Fill Down) | *Rellena seleccion con el valor de la primera celda* |
| `Ctrl + R` | Copiar celda de izquierda a derecha | *Rellena con contenido de la izq.* |
| `Delete` | Borrar contenido de celda(s) | *Solo contenido, no formato* |
| `Ctrl + - (menos)` | Eliminar celda/fila/columna | *Muestra opciones de eliminacion* |
| `Ctrl + + (mas)` | Insertar celda/fila/columna | *Muestra opciones de insercion* |
| `Ctrl + H` | Buscar y reemplazar | *Reemplazo masivo de datos* |
| `Ctrl + B` | Buscar (version espanol) | *Abre cuadro de Buscar; cambiar a pestana Reemplazar* |
| `Ctrl + F` | Buscar (version ingles) | *Buscar texto o valores* |
| `F3` | Pegar nombre definido | *Inserta nombres de rangos* |
| Doble clic en controlador de relleno | Copiar formula hacia abajo en toda la columna | *Detecta automaticamente hasta donde llenar* |

### 4. Formato

| Atajo | Accion | Tip |
| --- | --- | --- |
| `Ctrl + 1` | Formato de celdas (dialogo completo) | *Acceso a TODAS las opciones* |
| `Ctrl + N` | Negrita | *En Excel en espanol; Ctrl+B en ingles* |
| `Ctrl + K` | Cursiva | *En Excel en espanol; Ctrl+I en ingles* |
| `Ctrl + S` | Subrayado | *En Excel en espanol; Ctrl+U en ingles* |
| `Ctrl + Shift + $` | Formato moneda | *Aplica formato $#,##0.00* |
| `Ctrl + Shift + %` | Formato porcentaje | *Multiplica por 100 y agrega %* |
| `Ctrl + Shift + #` | Formato fecha | *Formato DD-MMM-AA* |
| `Ctrl + Shift + @` | Formato hora | *Formato HH:MM AM/PM* |
| `Ctrl + Shift + !` | Formato numero con miles | *Separador de miles y 2 decimales* |
| `Ctrl + Shift + ~` | Formato general | *Quita formato numerico especial* |
| `Alt + H, O, I` | Autoajustar ancho de columna | *Ruta de cinta rapida* |
| `Alt + H, O, A` | Autoajustar alto de fila | *Ruta de cinta rapida* |

---

### 5. Formulas

| Atajo | Accion | Tip |
| --- | --- | --- |
| `F4` | Alternar referencia absoluta/relativa | *$A$1 -> A$1 -> $A1 -> A1* |
| `Tab` | Autocompletar funcion sugerida | *Acepta la sugerencia de IntelliSense* |
| `Ctrl + ` | Mostrar/ocultar formulas en celdas | *Ver todas las formulas de la hoja* |
| `Alt + =` | Autosuma rapida | *Inserta =SUMA() automaticamente* |
| `Ctrl + Shift + U` | Expandir barra de formulas | *Ver formula completa si es larga* |
| `F9` | Evaluar parte de formula | *Seleccionar parte y F9 para ver resultado* |
| `Ctrl + '` | Copiar formula de celda superior | *Copia formula sin ajustar* |
| `Ctrl + Shift + Enter` | Formula matricial (legacy) | *Para versiones pre-365* |
| `Ctrl + Shift + A` | Insertar argumentos de funcion | *Muestra nombres de argumentos* |
| `F4 (fuera de edicion)` | Repetir ultima accion | *Repite formato, insercion, etc.* |

### 6. Tablas y Datos

| Atajo | Accion | Tip |
| --- | --- | --- |
| `Ctrl + T` | Crear tabla desde rango | *Detecta rango automaticamente* |
| `Alt + Flecha Abajo` | Abrir filtro de columna | *Dentro de tabla o con filtro activo* |
| `Ctrl + Shift + L` | Activar/desactivar filtros | *Toggle filtros en rango* |
| `Alt + D, S` | Ordenar (dialogo completo) | *Multiples niveles de orden* |
| `Ctrl + Shift + F3` | Crear nombres desde seleccion | *Nombra rangos automaticamente* |
| `Ctrl + T, luego Tab` | Tab para moverse en tabla | *Navega celda por celda en tabla* |
| `Alt + A, R, A` | Quitar duplicados | *Ruta de cinta: Datos > Quitar dup.* |
| `Alt + A, V, V` | Validacion de datos | *Ruta de cinta: Datos > Validacion* |
| `Ctrl + Shift + &` | Aplicar bordes al rango | *Borde exterior al rango seleccionado* |

### 7. Tablas Dinamicas

| Atajo | Accion | Tip |
| --- | --- | --- |
| `Alt + N, V` | Insertar tabla dinamica | *Ruta de cinta: Insertar > TD* |
| `Clic derecho > Opciones` | Acceder a opciones de TD | *Configuracion detallada de la TD* |
| `Doble clic en valor` | Drill-down (ver detalle) | *Crea hoja con datos de esa celda* |
| `Alt + Shift + Flecha Der` | Agrupar seleccion | *Agrupa filas/columnas seleccionadas* |
| `Alt + Shift + Flecha Izq` | Desagrupar seleccion | *Desagrupa filas/columnas* |
| `Clic derecho > Actualizar` | Actualizar tabla dinamica | *Refresca datos de la TD* |
| `Alt + F5` | Actualizar todas las TDs | *Refresca todas las conexiones* |
| `Clic derecho > Formato` | Formato de numero en TD | *Formato para campo de valor* |

### 8. Graficos

| Atajo | Accion | Tip |
| --- | --- | --- |
| `Alt + F1` | Crear grafico en hoja actual | *Grafico incrustado instantaneo* |
| `F11` | Crear grafico en hoja nueva | *Hoja de grafico dedicada* |
| `Ctrl + clic en elemento` | Seleccionar elemento del grafico | *Series, ejes, leyenda* |
| `Delete (en grafico)` | Eliminar elemento seleccionado | *Quita serie o elemento* |
| `Ctrl + 1 (en grafico)` | Formato del elemento seleccionado | *Panel de formato detallado* |
| `Flecha Arriba/Abajo` | Navegar entre series de datos | *Dentro del grafico* |

### 9. Atajos de Productividad

| Atajo | Accion | Tip |
| --- | --- | --- |
| `F4` | Repetir ultima accion | *Funciona para formato, borrado, etc.* |
| `Ctrl + ;` | Insertar fecha actual | *Fecha estatica (no cambia)* |
| `Ctrl + Shift + :` | Insertar hora actual | *Hora estatica (no cambia)* |
| `Ctrl + Shift + +` | Insertar fila/columna | *Segun seleccion previa* |
| `Alt + Enter` | Salto de linea en celda | *Multiples lineas en una celda* |
| `Ctrl + E` | Relleno rapido (Flash Fill) | *Detecta patrones automaticamente* |
| `Ctrl + Shift + V` | Pegado especial (menu) | *Elige que pegar: valores, formatos...* |
| `Alt + E, S, V` | Pegar solo valores | *Ruta clasica de pegado especial* |
| `Ctrl + P` | Imprimir / Vista previa | *Acceso rapido a impresion* |
| `Ctrl + W` | Cerrar libro actual | *No cierra Excel, solo el libro* |
| `F12` | Guardar como | *Dialogo completo de guardar* |
| `Ctrl + U` | Nuevo libro en blanco | *Ctrl+N en Excel en ingles* |


**Nota:** Algunos atajos pueden variar segun la version de Excel (2019, 2021, 365) y la configuracion de idioma. Los atajos mostrados corresponden a la version en espanol de Windows. En Mac, sustituir Ctrl por Cmd en la mayoria de los casos.


**Tip final:** No intenten memorizar todos los atajos de golpe. Elijan 3-5 que usen frecuentemente, practiquenlos una semana, y luego agreguen 3-5 mas. En un mes habran duplicado su velocidad en Excel.
