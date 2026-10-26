# Metodología de Resolución de Ejercicios de Estadística Inferencial

Esta guía define las reglas estrictas de formato y metodología para la resolución de guías de ejercicios de Estadística Inferencial (ANOVA, Kruskal-Wallis, Chi-cuadrado, Mann-Whitney, etc.) dictadas por el docente (Dr. Luis A. Fernández Vizcarra - Universidad Privada de Tacna).

## 1. Estructura General del Documento
- **Página y Márgenes:** Formato A4 (210 x 297 mm), márgenes superior e inferior de 0.7 pulgadas, laterales de 0.85 pulgadas.
- **Encabezado institucional:**
  ```text
  Universidad Privada de Tacna
  Facultad de Ingeniería
  Escuela Profesional de Ingeniería de Sistemas
  ```
- **Tipografía:** Calibri 10.5 pt para cuerpo de texto, interlineado 1.12 a 1.15, títulos en negrita 12-13 pt.

## 2. Pasos Obligatorios para Pruebas de 3 o más grupos (Kruskal-Wallis / ANOVA)
Para cada ejercicio, seguir estrictamente la siguiente secuencia:

1. **Paso 1: Formación de hipótesis**
   - H0: Las Promedios de los grupos son iguales ($\bar{x}_i = \bar{x}_y$)
   - H1: Por lo menos los Promedios de 2 grupos son diferentes ($\bar{x}_i \neq \bar{x}_y$)
2. **Paso 2: Determinar el nivel de significancia**
   - N.S = X% (ejemplo: 5% (0.05), 4% (0.04), 3% (0.03))
3. **Paso 3: Hallar estadísticos descriptivos**
   - Insertar tabla auténtica de SPSS "Informe" (Media, N, Desv. estándar, Total).
4. **Paso 4: Prueba de Distribución normal**
   - H0: Los datos de los grupos presentan distribución normal
   - H1: Por lo menos los datos de un grupo no presentan distribución normal
   - Insertar tabla auténtica de SPSS "Pruebas de normalidad" (Shapiro-Wilk: Estadístico, gl, Sig.).
   - Justificación con p-valores: "Dado que los p-valores de los grupos [...] son mayores que el nivel de significancia, pero el grupo [...] es menor [o menor o igual] al nivel de significancia por lo que es probable que se rechace la hipótesis nula"
   - Conclusión: "A un margen de error de X % se concluye que , los datos de los grupos [...], no presentan Distribución normal" (o sí presentan).
5. **Paso 5: Elegir la prueba estadística**
   - Si no hay normalidad: "Dado que los datos del grupo [...] no presentan distribución normal se aplicara la prueba Kruskal-Wallis (Prueba no paramétrica)"
     - Insertar tabla SPSS "Rangos" (N, Rango promedio).
     - Insertar tabla SPSS "Estadísticos de prueba" (H de Kruskal-Wallis, gl, Sig. asin.).
   - Si sí hay normalidad: Evaluar homogeneidad de varianzas con Levene (Paso 5.A) y aplicar ANOVA (Prueba paramétrica).
6. **Paso 6: Contrastación de Hipótesis**
   - "Dado que el p-valor de Kruskal-Wallis (X.XXX) es menor [o mayor] que el nivel de significancia (X.XX), se rechaza [o no se rechaza] la hipótesis nula."
7. **Paso 7: Conclusión** (REGLA CRÍTICA: NO EXPANDIRSE, MAXIMO 1-2 LÍNEAS DIRECTAS)
   - Formato exacto:  
     "A un margen de error del X% se concluye que, por lo menos los [nombres de variable/métrica] de 2 [grupos] son diferentes."
     (o si no se rechaza H0: "A un margen de error del X% se concluye que, Las Promedios de los grupos son iguales.")

## 3. Generación de Tablas SPSS
- Utilizar Playwright y CSS fiel a IBM SPSS Statistics 26-29 (bordes azul marino `#103B60`, cabecera gris azulada `#E8ECEF`, flecha roja de navegación `►`, fuente Arial 11px) guardadas en PNG e insertadas en el Word.

## 4. Scripts y Herramientas del Repositorio
- `generate_kruskal_images.py`: Genera capturas de tablas SPSS nítidas y auténticas.
- `build_kruskal_doc.py`: Ensambla el documento Word `.docx` con espaciado óptimo y exporta automáticamente a `.pdf` vía Word COM automation.
