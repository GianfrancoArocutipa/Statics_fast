* Sintaxis SPSS para Ejercicio 1: Chi-cuadrado de bondad de ajuste.
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.
WEIGHT BY Frecuencia.

NPAR TESTS
  /CHISQUARE=Categoria
  /EXPECTED=EQUAL
  /STATISTICS DESCRIPTIVES
  /MISSING ANALYSIS.

WEIGHT OFF.
