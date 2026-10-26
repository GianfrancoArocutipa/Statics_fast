# Solucionario y Generador de Informes de Estadística Inferencial

Repositorio para la resolución automatizada, cálculo estadístico y maquetación de informes académicos de **Estadística Inferencial y Análisis de Datos** para la **Universidad Privada de Tacna (UPT)** - Escuela Profesional de Ingeniería de Sistemas (Docente: Dr. Luis A. Fernández Vizcarra).

---

## 📋 Estructura del Repositorio

- **`AGENTS.md` / `GEMINI.md`:** Instrucciones automáticas para agentes de IA para que al clonar el repositorio resuelvan cualquier guía Word (.docx) o PDF siguiendo la metodología y formato exacto del docente.
- **`.agents/rules/`:** Reglas permanentes del espacio de trabajo sobre el formato de los pasos (Paso 1 a Paso 7/8), redacción estricta y concisa de conclusiones, y criterios estadísticos.
- **`generate_kruskal_images.py`:** Script generador de tablas auténticas de IBM SPSS Statistics 26-29 en formato PNG (Informe, Shapiro-Wilk, Rangos, Estadísticos de prueba).
- **`build_kruskal_doc.py`:** Generador del documento Word A4 con membrete institucional y exportador a PDF nativo.
- **`generated_kruskal_images/`:** Directorio con las capturas de SPSS de alta definición generadas para cada ejercicio.

---

## 🚀 Instalación y Puesta en Marcha

Para cualquier persona o entorno que clone este repositorio:

```bash
# 1. Clonar el repositorio
git clone <URL_DEL_REPOSITORIO>
cd Estats

# 2. Instalar dependencias de Python
pip install -r requirements.txt

# 3. Instalar navegadores para el generador de tablas SPSS (Playwright)
playwright install chromium msedge
```

---

## 🛠️ Cómo Resolver Nuevos Ejercicios

Cualquier agente de IA conectado a este repositorio leerá automáticamente `AGENTS.md` y sabrá cómo proceder. Para resolver una nueva guía de ejercicios:

1. Proporciona el archivo `.docx` o enunciado de los ejercicios.
2. El asistente o script:
   - Calculará las pruebas de normalidad (**Shapiro-Wilk**).
   - Determinará la prueba adecuada (**Kruskal-Wallis** para muestras no normales, o **ANOVA** + **Tukey** para muestras normales con varianzas homogéneas).
   - Renderizará las tablas con estética oficial de SPSS.
   - Compilará el informe en formato Word `.docx` y `.pdf` manteniendo las **conclusiones breves y directas** en el Paso 7/8.
