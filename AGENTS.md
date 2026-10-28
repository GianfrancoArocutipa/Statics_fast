# AGENTS.md - Instrucciones de Comportamiento para Asistentes de IA

Este repositorio está configurado para la resolución automatizada y generación de informes de ejercicios prácticos de **Estadística Inferencial** de la **Universidad Privada de Tacna (Facultad de Ingeniería, Escuela Profesional de Ingeniería de Sistemas)**, dictados por el docente **Dr. Luis A. Fernández Vizcarra**.

Cualquier agente que trabaje en este repositorio debe seguir **estrictamente** las directrices descritas a continuación.

---

## 1. Disparador / Comando Obligatorio: "preparate" o "prepárate"

Cuando el usuario escriba **"preparate"**, **"prepárate"**, **"prepara el entorno"** o similar inmediatamente después de clonar el repositorio:
El agente DEBE ejecutar de forma autónoma y sin pedir confirmación adicional los siguientes comandos:
```bash
pip install -r requirements.txt
playwright install chromium
```
Y verificar con una prueba rápida de importación (`python -c "import scipy, docx, playwright; print('OK')"`), notificando al usuario que el entorno está completamente listo para recibir cualquier archivo Word (.docx) y resolverlo al instante.

---

## 2. Flujo de Trabajo Obligatorio al Recibir Ejercicios

> [!IMPORTANT]
> **REGLA CRÍTICA DE PRESERVACIÓN DE FORMATO:**
> Siempre que el usuario envíe o proporcione un documento (`.docx`), se DEBE mantener **TODO el formato** original (colores de texto, sombreados, paleta cromática, tipografías, tamaños de fuente, márgenes, encabezados, pies de página, tablas, bordes y estilos de párrafos). **Únicamente se debe modificar/actualizar la información y los datos** (cálculos, tablas SPSS, pasos de hipótesis y conclusiones), sin alterar la estética ni la estructura visual del documento base.

Cuando el usuario proporcione un documento Word (.docx), PDF o texto con ejercicios a resolver:

1. **Extracción y Análisis de Datos:**
   - Extraer con precisión las muestras numéricas de cada grupo.
   - Calcular mediante Python (`scipy.stats`, `numpy`):
     - Estadísticos descriptivos (Media, N, Desviación estándar).
     - Prueba de normalidad **Shapiro-Wilk** ($W$ y $p$-valor para cada grupo).
     - Prueba de homogeneidad de varianzas **Levene** (basada en la media, gl1, gl2, $p$-valor).
     - Prueba no paramétrica **Kruskal-Wallis** ($H$, gl, $p$-valor asintótico, rangos promedio) si falla la normalidad en algún grupo ($p < \alpha$).
     - Prueba paramétrica **ANOVA** ($F$, gl, $p$-valor, medias cuadráticas) y post hoc **Tukey** si se cumple la normalidad en todos los grupos ($p \ge \alpha$).

2. **Generación de Tablas con Estilo SPSS Auténtico:**
   - No usar tablas genéricas de Word.
   - Utilizar el motor de renderizado HTML + Playwright (como en `generate_kruskal_images.py` o `generate_spss_images.py`) para producir imágenes `.png` idénticas a las salidas de IBM SPSS Statistics 26-29:
     - Cabecera con bordes azul marino (`#103B60`), fondo gris azulado (`#E8ECEF`), flecha roja de navegación (`►`), fuente Arial.
   - Generar para cada ejercicio:
     - `Informe` (Descriptivos)
     - `Pruebas de normalidad` (Shapiro-Wilk)
     - `Rangos` y `Estadísticos de prueba` (Kruskal-Wallis) O `ANOVA` y `Comparaciones múltiples` (Tukey).

3. **Estructura Estricta de Resolución (Paso a Paso):**
   - **Paso 1: Formación de hipótesis**
     - $H_0$: Las Promedios de los grupos son iguales ($\bar{x}_i = \bar{x}_y$)
     - $H_1$: Por lo menos los Promedios de 2 grupos son diferentes ($\bar{x}_i \neq \bar{x}_y$)
   - **Paso 2: Determinar el nivel de significancia**
     - N.S = X% (ej: 5% (0.05), 4% (0.04), 3% (0.03))
   - **Paso 3: Hallar estadísticos descriptivos**
     - Insertar la imagen de la tabla SPSS `Informe`.
   - **Paso 4: Prueba de Distribución normal**
     - $H_0$: Los datos de los grupos [...] presentan distribución normal
     - $H_1$: Por lo menos los datos de un grupo [...] no presentan distribución normal
     - Insertar imagen de la tabla SPSS `Pruebas de normalidad`.
     - Texto de decisión: *"Dado que los p-valores de los grupos [...] son mayores que el nivel de significancia, pero el grupo [...] es menor [o igual] al nivel de significancia por lo que es probable que se rechace la hipótesis nula"*
     - Conclusión: *"Conclusión: A un margen de error de X % se concluye que , los datos de los grupos [...], no presentan Distribución normal"*
   - **Paso 5: Elegir la prueba estadística**
     - *"Dado que los datos del grupo [...] no presentan distribución normal se aplicara la prueba Kruskal-Wallis (Prueba no paramétrica)"*
     - Insertar imágenes de las tablas SPSS `Rangos` y `Estadísticos de prueba`.
   - **Paso 6: Contrastación de Hipótesis**
     - *"Dado que el p -valor de Kruskal -Wallis (X.XXX) es menor que el nivel de significancia (X.XX), se rechaza la hipótesis nula."*
   - **Paso 7: Conclusión (REGLA CRÍTICA: NO EXPANDIRSE)**
     - Mantener la conclusión en **una sola oración directa**, sin añadir párrafos descriptivos adicionales de quién tuvo mayor o menor puntaje, siguiendo el estilo exacto de los ejercicios 1 y 2 resueltos por el docente:
     - Formato:  
       `A un margen de error del X% se concluye que, por lo menos los [métrica] de 2 [unidades/grupos] son diferentes.`

4. **Compilación de Documentos Finales:**
   - Usar `python-docx` para compilar el archivo `.docx` en formato A4 con encabezado institucional:
     ```text
     Universidad Privada de Tacna
     Facultad de Ingeniería
     Escuela Profesional de Ingeniería de Sistemas
     ```
   - Exportar a `.pdf` nativo de Microsoft Word mediante `win32com.client` para un resultado limpio y uniforme.
   - Proporcionar siempre enlaces markdown clickeables tanto al `.docx` como al `.pdf`.
