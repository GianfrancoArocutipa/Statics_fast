# Solucionario y Generador de Informes de Estadística Inferencial

Repositorio para la resolución automatizada, cálculo estadístico y maquetación de informes académicos de **Estadística Inferencial y Análisis de Datos** para la **Universidad Privada de Tacna (UPT)** - Escuela Profesional de Ingeniería de Sistemas (Docente: Dr. Luis A. Fernández Vizcarra).

---

## ⚡ Activación Inmediata con la IA

Cuando clones este repositorio y abras el chat de la IA (Antigravity / Gemini), solo escribe:

> **`preparate`** (o **`prepárate`**)

La IA automáticamente ejecutará:
```bash
pip install -r requirements.txt
playwright install chromium
```
Y dejará el entorno 100% operativo sin que tengas que teclear ningún comando manualmente.

---

## 📋 Estructura del Repositorio

- **`AGENTS.md` / `GEMINI.md`:** Instrucciones automáticas para agentes de IA para que al clonar el repositorio reconozcan el comando `preparate` y resuelvan cualquier guía Word (.docx) o PDF siguiendo la metodología y formato exacto del docente.
- **`.agents/rules/`:** Reglas permanentes del espacio de trabajo sobre el formato de los pasos (Paso 1 a Paso 7/8), redacción estricta y concisa de conclusiones, y criterios estadísticos.
- **`setup.bat`:** Script por lotes para Windows que ejecuta la instalación en un solo clic si prefieres no usar el chat.
- **`generate_kruskal_images.py`:** Script generador de tablas auténticas de IBM SPSS Statistics 26-29 en formato PNG (Informe, Shapiro-Wilk, Rangos, Estadísticos de prueba).
- **`build_kruskal_doc.py`:** Generador del documento Word A4 con membrete institucional y exportador a PDF nativo.
- **`generated_kruskal_images/`:** Directorio con las capturas de SPSS de alta definición generadas para cada ejercicio.

---

## 🛠️ Cómo Resolver Nuevos Ejercicios

1. Escribe a la IA: `preparate`
2. Pasa el archivo `.docx` o enunciado de los ejercicios:
   > *"Resuelve este documento: @[Guia_Ejercicios.docx]"*
3. La IA automáticamente:
   - Extraerá los datos numéricos.
   - Verificará supuestos (Shapiro-Wilk para normalidad y Levene para varianzas).
   - Elegirá la prueba correcta (Kruskal-Wallis para no paramétricas o ANOVA + Tukey para paramétricas).
   - Generará las tablas SPSS en PNG.
   - Compilará y entregará el informe en `.docx` y `.pdf` con conclusiones directas.
