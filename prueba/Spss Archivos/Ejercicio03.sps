* Sintaxis SPSS para Ejercicio 3: Prueba de Rachas (Aleatoriedad de accesos).
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.

NPAR TESTS
  /RUNS(MEDIAN)=Acceso
  /STATISTICS DESCRIPTIVES
  /MISSING ANALYSIS.
