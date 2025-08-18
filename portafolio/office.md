#
#
# Avanzado
## 🔹 1. Funciones avanzadas
BUSCARX (o BUSCARV + COINCIDIR en versiones anteriores).
ÍNDICE + COINCIDIR para búsquedas dinámicas.
SI.CONJUNTO para condiciones múltiples.
FILTRO (en Excel 365) para extraer datos dinámicamente.
UNIRCADENAS para concatenar texto de varias celdas.
📌 Ejemplo: buscar el sueldo de un empleado en una tabla grande usando ÍNDICE + COINCIDIR.
## 🔹 2. Tablas dinámicas avanzadas
Agrupar datos por meses, trimestres o años automáticamente.
Crear campos calculados (ej: calcular % de participación).
Usar segmentaciones y cronologías para filtrar de forma interactiva.
📌 Ejemplo: analizar las ventas de un año completo y ver la evolución mensual con un par de clics.
## 🔹 3. Gráficos dinámicos e interactivos
Combinar gráficos (ej: columna + línea).
Gráficos vinculados a segmentaciones de una tabla dinámica.
Gráficos con barra de progreso o termómetros para KPIs.
📌 Ejemplo: un gráfico de ventas por región que cambia al elegir el año con un filtro.
## 🔹 4. Validación de datos + listas dinámicas
Crear listas desplegables dependientes (ej: seleccionar “Región” y que la segunda lista muestre solo “Ciudades de esa región”).
📌 Ejemplo: seleccionar un país y que en la siguiente celda aparezcan solo sus ciudades.
## 🔹 5. Automatización con macros (VBA)
Grabar una macro para tareas repetitivas.
Crear un botón que al hacer clic genere un reporte.
📌 Ejemplo: un botón que copie datos filtrados y los pegue en una nueva hoja con formato.
## 🔹 6. Análisis de datos avanzado
Uso de Tablas de datos y Escenarios para proyecciones.
Buscar objetivo y Solver para optimizar decisiones (ej: ¿cuántas unidades debo vender para alcanzar X utilidad?).
📌 Ejemplo: calcular automáticamente la cantidad mínima de ventas necesarias para cubrir costos.

### BUSCARX
La función BUSCARX en Excel es una de las más potentes y modernas, y básicamente reemplaza a BUSCARV, BUSCARH y hasta algunas combinaciones de ÍNDICE + COINCIDIR.

📖 Parámetros:

valor_buscado → lo que quieres encontrar (ej: un código, nombre, etc.).
matriz_buscar → la columna o fila donde Excel va a buscar ese valor.
matriz_resultado → la columna o fila desde donde quieres devolver el resultado.
[si_no_se_encuentra] → opcional, valor que quieres que aparezca si no encuentra coincidencia (ej: "No existe").
[modo_coincidencia] → cómo debe coincidir:

0 → Coincidencia exacta (por defecto).
-1 → Coincidencia exacta o el siguiente más pequeño.
1 → Coincidencia exacta o el siguiente más grande.
2 → Permite usar comodines (* y ?).

[modo_busqueda] → dirección de búsqueda:

1 → De arriba a abajo o izquierda a derecha (por defecto).
-1 → De abajo hacia arriba o de derecha a izquierda.
2 → Búsqueda binaria ascendente.
-2 → Búsqueda binaria descendente.

### 1. COINCIDIR

Busca la posición (número de fila o columna) donde se encuentra un valor dentro de un rango.
```bash
=COINCIDIR(valor_buscado; rango; [tipo])
```
valor_buscado → lo que quieres encontrar.
rango → dónde buscar.
tipo → 0 (coincidencia exacta), 1 (menor más cercano), -1 (mayor más cercano).

Ventajas sobre BUSCARV:

Puedes buscar a la izquierda o derecha (no está limitado como BUSCARV).
Es más rápido en tablas grandes.
No se rompe si insertas o eliminas columnas.
### 2. ÍNDICE
Devuelve el valor de una celda en una posición específica de un rango.
```bash
=ÍNDICE(rango; número_fila; [número_columna])
```
### SI.CONJUNTO
La función SI.CONJUNTO en Excel es muy útil cuando necesitas evaluar varias condiciones a la vez, sin tener que anidar múltiples funciones SI.
```bash
=SI.CONJUNTO(condición1; valor_si_verdadero1; condición2; valor_si_verdadero2; ... )

=SI.CONJUNTO(
 A2>=60;"Aprobado";
 A2<60;"Reprobado"
)
```
📌 Ventaja: Es más limpio que usar muchos SI( ... ) anidados.
📌 Ojo: No tiene opción de "si no cumple ninguna condición", a menos que tú la pongas al final como VERDADERO; "Otro".
### FILTRO
matriz: el rango de celdas que quieres filtrar.
incluir: la condición o condición(es) que deben cumplirse (puede ser lógica, con =, >, <, etc.).
si_vacío (opcional): el valor que se devuelve si no hay datos que cumplan la condición. Por defecto, da error #CALC!.
```bash
=FILTRO(matriz, incluir, [si_vacío])
```
Y: todas las condiciones deben cumplirse → multiplicar las condiciones
O: alguna de las condiciones debe cumplirse → sumar las condiciones

