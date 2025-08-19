# BUSCARX: Evaluar distintos rangos de datos
La función BUSCARX en Excel funciona siempre de la forma:
```bash
=BUSCARX(valor_buscado, matriz_buscada, matriz_devolución, [si_no_encontrado], [modo_coincidencia], [modo_busqueda])
```
📌 Puntos clave:
matriz_buscada y matriz_devolución deben ser rangos unidimensionales (una sola fila o una sola columna).
No puedes pasarle directamente una matriz bidimensional (ej: A:C) y esperar que busque en varias columnas a la vez.
# COINCIDIR: Revisar como realizar busqueda de abajo hacia arriba, de derecha a izquierda.
Por defecto, COINCIDIR busca siempre de arriba hacia abajo (primera coincidencia).
No tiene un argumento nativo para hacerlo “de abajo hacia arriba” o “de derecha a izquierda” (porque solo trabaja en 1D).
```bash
=FILAS(C:C) - COINCIDIR("Ana", ÍNDICE(C:C, FILAS(C:C)):C1, 0) + 1
```
# ÍNDICE: Revisar error de procesamiento
La función INDICE devuelve un valor o la referencia a un valor desde una tabla o rango.

Hay dos formas de utilizar la función INDICE:
- Si desea devolver el valor de una celda especificada o de una matriz de celdas, consulte Forma de matriz.
[LINK](https://support.microsoft.com/es-es/office/indice-funci%C3%B3n-indice-a5dcf0dd-996d-40a4-a822-b56b061328bd#bmarray_form)
- Si desea devolver una referencia a las celdas especificadas, consulte Forma de referencia.
[LINK](https://support.microsoft.com/es-es/office/indice-funci%C3%B3n-indice-a5dcf0dd-996d-40a4-a822-b56b061328bd#__reference_form)

```bash
=INDICE(matriz, num_fila, num_columna)
```
- matriz(Obligatorio). Es un rango de celdas o una constante de matriz.
Si matriz contiene solo una fila o columna, el argumento núm_fila o núm_columna correspondiente es opcional.
Si matriz tiene varias filas y columnas, y solo usa núm_fila o núm_columna, INDICE devuelve una matriz de dicha fila o columna completa.
- fila(Obligatorio, a menos que núm_columna esté presente). Selecciona la fila de la matriz desde la cual devolverá un valor. Si se omite núm_fila, núm_columna es obligatorio.
- núm_columna(Opcional). Selecciona la columna de la matriz desde la cual devolverá un valor. Si se omite núm_columna, núm_fila es obligatorio.
