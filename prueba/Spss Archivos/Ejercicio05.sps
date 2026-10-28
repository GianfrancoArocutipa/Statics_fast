* Sintaxis SPSS para Ejercicio 5: ANOVA de un factor y Post Hoc Tukey.
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.

* 1. Descriptivos (Informe).
MEANS TABLES=Tiempo_ms BY Servicio
  /CELLS=MEAN COUNT STDDEV.

* 2. Normalidad Shapiro-Wilk.
EXAMINE VARIABLES=Tiempo_ms BY Servicio
  /PLOT NPPLOT
  /STATISTICS DESCRIPTIVES
  /CINTERVAL 97
  /MISSING LISTWISE
  /NOTOTAL.

* 3. Homogeneidad de varianzas Levene, ANOVA y Post Hoc Tukey.
ONEWAY Tiempo_ms BY Servicio
  /STATISTICS DESCRIPTIVES HOMOGENEITY 
  /MISSING ANALYSIS
  /POSTHOC=TUKEY ALPHA(0.05).

* 4. Kruskal-Wallis (Prueba no paramétrica alternativa).
NPAR TESTS
  /K-W=Tiempo_ms BY Servicio(1 3)
  /STATISTICS DESCRIPTIVES
  /MISSING ANALYSIS.
