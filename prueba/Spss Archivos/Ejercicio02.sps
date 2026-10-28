* Sintaxis SPSS para Ejercicio 2: Comparación de 2 Algoritmos Independientes.
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.

* 1. Descriptivos (Informe).
MEANS TABLES=Tiempo_ms BY Algoritmo
  /CELLS=MEAN COUNT STDDEV.

* 2. Prueba de normalidad Shapiro-Wilk.
EXAMINE VARIABLES=Tiempo_ms BY Algoritmo
  /PLOT NPPLOT
  /STATISTICS DESCRIPTIVES
  /CINTERVAL 95
  /MISSING LISTWISE
  /NOTOTAL.

* 3. Prueba T para muestras independientes (con Levene).
T-TEST GROUPS=Algoritmo(1 2)
  /MISSING=ANALYSIS
  /VARIABLES=Tiempo_ms
  /CRITERIA=CI(.99).

* 4. Prueba no paramétrica U de Mann-Whitney (Alternativa).
NPAR TESTS
  /M-W= Tiempo_ms BY Algoritmo(1 2)
  /STATISTICS=DESCRIPTIVES
  /MISSING ANALYSIS.
