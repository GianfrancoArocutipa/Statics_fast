* Sintaxis SPSS para Ejercicio 4: Prueba Binomial (Proporción de éxito).
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.
WEIGHT BY Cantidad.

NPAR TESTS
  /BINOMIAL (0.90)=Estado
  /MISSING ANALYSIS.

WEIGHT OFF.
