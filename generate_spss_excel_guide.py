# -*- coding: utf-8 -*-
"""
Generador de Guía Completa de SPSS en Excel (.xlsx)
Contiene el paso a paso exacto con botones, menús, casillas y ventanas de diálogo
para la resolución de ejercicios de Estadística Inferencial en IBM SPSS Statistics.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_spss_excel_guide(filename="Guia_Pasos_SPSS_Estadistica_Inferencial.xlsx"):
    wb = openpyxl.Workbook()
    
    # Fuentes y estilos
    font_title = Font(name="Calibri", size=14, bold=True, color="103B60")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="4A5568")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10, color="1A202C")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1A202C")
    font_badge = Font(name="Calibri", size=9, bold=True, color="0D47A1")

    fill_header = PatternFill(start_color="103B60", end_color="103B60", fill_type="solid")
    fill_header_sub = PatternFill(start_color="1E5B94", end_color="1E5B94", fill_type="solid")
    fill_zebra = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_highlight = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    fill_btn = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CBD5E0")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom = Border(bottom=Side(border_style="medium", color="103B60"))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = [
        "Paso Nº",
        "Fase / Objetivo",
        "Ruta en Menú SPSS",
        "Ventana / Cuadro de Diálogo",
        "Botón o Pestaña Específica",
        "Acción Específica (Casillas, Parámetros y Campos)",
        "Resultado Esperado / Observaciones"
    ]

    col_widths = [8, 22, 28, 26, 25, 45, 30]

    # ==========================================
    # HOJA 1: RESUMEN Y DECISIÓN DE PRUEBAS
    # ==========================================
    ws_index = wb.active
    ws_index.title = "Flujo de Decisión"
    ws_index.views.sheetView[0].showGridLines = True

    ws_index.merge_cells("A1:G1")
    ws_index["A1"] = "GUÍA RÁPIDA: FLUJO DE DECISIÓN ESTADÍSTICA EN SPSS"
    ws_index["A1"].font = font_title
    ws_index["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws_index.merge_cells("A2:G2")
    ws_index["A2"] = "Universidad Privada de Tacna | Escuela Profesional de Ingeniería de Sistemas | Dr. Luis A. Fernández Vizcarra"
    ws_index["A2"].font = font_subtitle
    ws_index["A2"].alignment = Alignment(horizontal="left", vertical="center")

    flow_headers = ["Paso / Fase", "¿Qué se evalúa?", "Condición / Pregunta", "Decisión Estadística", "Prueba a Ejecutar", "Hoja con Pasos", "Archivos a Generar"]
    for c_idx, h_text in enumerate(flow_headers, 1):
        cell = ws_index.cell(row=4, column=c_idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = thin_border

    flow_data = [
        ("1. Carga de Datos", "Creación de matriz en SPSS", "Datos en bruto por grupos", "Estructurar 2 variables (Grupo y Valor)", "Vista de Variables y Datos", "01_Config_Variables", "EjercicioXX_GF.sav"),
        ("2. Descriptivos", "Informe de medias", "Siempre se calcula", "Presentar Media, N y Desv. Estándar", "Medias (Informe)", "02_Descriptivos_Informe", "Tabla Informe"),
        ("3. Normalidad", "Shapiro-Wilk (por grupo)", "Muestra por grupo n <= 50", "Evaluar Sig. (p-valor) vs alpha (0.05)", "Explorar (Shapiro-Wilk)", "03_Normalidad_Shapiro", "Tabla Normalidad"),
        ("4.A Decisión 3+ grupos", "Normalidad en TODOS los grupos", "Si al menos UN grupo tiene p < alpha", "No se cumple supuesto paramétrico", "Kruskal-Wallis (No paramétrica)", "04_Kruskal_Wallis", "Rangos + Est. Prueba"),
        ("4.B Decisión 3+ grupos", "Normalidad en TODOS los grupos", "Si TODOS los grupos tienen p >= alpha", "Se cumple normalidad -> Probar Levene", "Homogeneidad (Levene) y ANOVA", "05_ANOVA_PostHoc", "ANOVA + Tukey / Games"),
        ("5. Decisión 2 grupos", "Comparación de 2 muestras indep.", "Si no cumple normalidad", "Prueba no paramétrica de 2 grupos", "U de Mann-Whitney", "06_Mann_Whitney", "Rangos + Est. Prueba"),
        ("6. Decisión 2 grupos", "Comparación de 2 muestras indep.", "Si cumple normalidad", "Prueba paramétrica de 2 grupos", "T de Student / Welch", "07_T_Student", "Tabla T Muestras Indep."),
        ("7. Muestras Pareadas", "Mismo sujeto Antes - Después", "Evaluar normalidad de diferencias", "Si diferencias son normales", "T de Student para Muestras Emparejadas", "08_Muestras_Pareadas", "Tabla T Emparejadas"),
        ("8. Frecuencias / Cualitativo", "Independencia o Bondad de Ajuste", "Variables cualitativas / categorías", "Comparar observadas vs esperadas", "Chi-cuadrado (Tablas cruzadas)", "09_Chi_Cuadrado", "Tabla Chi-cuadrado"),
        ("9. Guardado Final", "Entrega formal al docente", "Finalización de corrida", "Guardar datos (.sav) y visor (.spv)", "Guardar como...", "10_Guardar_SAV_SPV", "XX_GF.sav y XX.spv")
    ]

    for r_idx, row_items in enumerate(flow_data, 5):
        fill_curr = fill_zebra if r_idx % 2 == 1 else fill_white
        for c_idx, val in enumerate(row_items, 1):
            cell = ws_index.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if c_idx in [1, 5] else font_body
            cell.fill = fill_curr
            cell.alignment = align_center if c_idx in [1, 6] else align_left
            cell.border = thin_border

    flow_widths = [18, 26, 32, 32, 28, 22, 22]
    for i, w in enumerate(flow_widths, 1):
        ws_index.column_dimensions[get_column_letter(i)].width = w

    # ==========================================
    # DEFINICIÓN DE HOJAS DETALLADAS
    # ==========================================
    sheets_info = [
        {
            "code": "01_Config_Variables",
            "title": "CONFIGURACIÓN DE VARIABLES Y DATOS EN SPSS (.SAV)",
            "desc": "Procedimiento exacto para definir variables en Vista de Variables e introducir datos en Vista de Datos.",
            "steps": [
                ("1", "Abrir SPSS", "Archivo > Nuevo > Datos", "Ventana Principal SPSS", "Pestaña 'Vista de variables'", "Hacer clic en la pestaña inferior izquierda 'Vista de variables'.", "Se muestra la grilla con columnas: Nombre, Tipo, Anchura, Decimales, Etiqueta, Valores, Medida, etc."),
                ("2", "Crear Variable Factor", "Vista de variables (Fila 1)", "Columna: Nombre", "Teclado", "Escribir el nombre sin espacios: 'Grupo' o 'Sistema' o 'Tratamiento'.", "Nombre asignado."),
                ("3", "Ajustar Decimales Factor", "Vista de variables (Fila 1)", "Columna: Decimales", "Flechas arriba/abajo", "Cambiar de 2 a 0 (las categorías son enteros: 1, 2, 3...).", "Decimales en 0."),
                ("4", "Asignar Etiqueta Factor", "Vista de variables (Fila 1)", "Columna: Etiqueta", "Teclado", "Escribir descripción completa, ej: 'Sistema Operativo evaluado'.", "Etiqueta descriptiva lista."),
                ("5", "Codificar Valores Factor", "Vista de variables (Fila 1)", "Columna: Valores", "Botón [...] (puntos suspensivos)", "Clic en [...]. En 'Valor' escribir '1', en 'Etiqueta' escribir 'Sistema X' -> Clic en botón [Añadir]. Repetir para '2' = 'Sistema Y' [Añadir] y '3' = 'Sistema Z' [Añadir]. Clic en [Aceptar].", "Valores codificados: {1, Sistema X; 2, Sistema Y; 3, Sistema Z}."),
                ("6", "Definir Escala Factor", "Vista de variables (Fila 1)", "Columna: Medida", "Menú desplegable", "Seleccionar 'Nominal' (icono de tres círculos de colores).", "Medida establecida como Nominal."),
                ("7", "Crear Variable Numérica", "Vista de variables (Fila 2)", "Columna: Nombre", "Teclado", "Escribir el nombre: 'Tiempo' o 'Rendimiento' o 'Puntaje'.", "Segunda variable creada."),
                ("8", "Ajustar Decimales Numérica", "Vista de variables (Fila 2)", "Columna: Decimales", "Flechas arriba/abajo", "Colocar '2' si tiene decimales o '0' si son enteros.", "Decimales configurados."),
                ("9", "Asignar Etiqueta Numérica", "Vista de variables (Fila 2)", "Columna: Etiqueta", "Teclado", "Escribir: 'Tiempo de respuesta en milisegundos (ms)'.", "Etiqueta lista."),
                ("10", "Definir Escala Numérica", "Vista de variables (Fila 2)", "Columna: Medida", "Menú desplegable", "Seleccionar 'Escala' (icono de regla amarilla).", "Medida establecida como Escala."),
                ("11", "Digitar o Pegar Datos", "Barra inferior de SPSS", "Pestaña 'Vista de datos'", "Clic en 'Vista de datos'", "En la Columna 'Grupo' colocar el código (1, 2 o 3) y en 'Tiempo' el valor medido correspondiente a cada observación.", "Matriz de datos completa lista para el análisis.")
            ]
        },
        {
            "code": "02_Descriptivos_Informe",
            "title": "PASO 3: ESTADÍSTICOS DESCRIPTIVOS (TABLA INFORME DE MEDIAS)",
            "desc": "Generación de la tabla auténtica SPSS 'Informe' con Media, N y Desviación estándar por grupo y Total.",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Comparar medias > Medias...", "Colocar el cursor sobre 'Comparar medias' y hacer clic en 'Medias...'.", "Se abre el cuadro de diálogo 'Medias'."),
                ("2", "Asignar Var. Dependiente", "Cuadro: Medias", "Panel izquierdo (Lista de variables)", "Botón Flecha [ ► ] superior", "Seleccionar la variable numérica 'Tiempo' y hacer clic en la flecha [ ► ] apuntando a 'Lista de dependientes'.", "'Tiempo' pasa a 'Lista de dependientes'."),
                ("3", "Asignar Var. Independiente", "Cuadro: Medias", "Panel izquierdo (Lista de variables)", "Botón Flecha [ ► ] inferior", "Seleccionar la variable de grupo 'Grupo' y hacer clic en la flecha [ ► ] apuntando a 'Capa 1 de 1' (Independientes).", "'Grupo' pasa a 'Capa 1 de 1'."),
                ("4", "Configurar Estadísticos", "Cuadro: Medias", "Panel derecho", "Botón [Opciones...]", "Hacer clic en el botón [Opciones...].", "Se abre el cuadro 'Medias: Opciones'."),
                ("5", "Seleccionar Métricas", "Cuadro: Medias: Opciones", "Listas 'Estadísticos' y 'Estadísticos de celda'", "Botón Flecha [ ► ]", "Verificar que en 'Estadísticos de celda' aparezcan: 1. Media, 2. Número de casos (N), 3. Desviación típica. Si falta alguno, seleccionarlo de la lista izquierda y mover con [ ► ].", "Las 3 métricas oficiales del docente quedan en el orden exacto."),
                ("6", "Confirmar Opciones", "Cuadro: Medias: Opciones", "Panel inferior derecho", "Botón [Continuar]", "Hacer clic en el botón [Continuar].", "Regresa al cuadro principal 'Medias'."),
                ("7", "Ejecutar Análisis", "Cuadro: Medias", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "SPSS procesa y abre el Visor de Resultados con la tabla 'Informe' conteniendo Media, N y Desv. típica por grupo.")
            ]
        },
        {
            "code": "03_Normalidad_Shapiro",
            "title": "PASO 4: PRUEBA DE DISTRIBUCIÓN NORMAL (SHAPIRO-WILK)",
            "desc": "Evaluación del supuesto de normalidad en cada grupo mediante Shapiro-Wilk (n <= 50) usando la función Explorar.",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Estadísticos descriptivos > Explorar...", "Hacer clic en 'Analizar' > 'Estadísticos descriptivos' > 'Explorar...'.", "Se abre el cuadro de diálogo 'Explorar'."),
                ("2", "Asignar Dependiente", "Cuadro: Explorar", "Panel izquierdo de variables", "Botón Flecha [ ► ] superior", "Seleccionar 'Tiempo' y hacer clic en la flecha [ ► ] apuntando a 'Lista de dependientes'.", "'Tiempo' pasa a 'Lista de dependientes'."),
                ("3", "Asignar Factor de Grupo", "Cuadro: Explorar", "Panel izquierdo de variables", "Botón Flecha [ ► ] medio", "Seleccionar 'Grupo' y hacer clic en la flecha [ ► ] apuntando a 'Lista de factores'.", "'Grupo' pasa a 'Lista de factores'."),
                ("4", "Configurar Gráficos", "Cuadro: Explorar", "Panel derecho", "Botón [Gráficos...]", "Hacer clic en el botón [Gráficos...].", "Se abre el cuadro 'Explorar: Gráficos'."),
                ("5", "Activar Prueba de Normalidad", "Cuadro: Explorar: Gráficos", "Sección Central", "Casilla [X] Gráficos de normalidad con pruebas", "MARCAR obligatoriamente la casilla: '[X] Gráficos de normalidad con pruebas'.", "Se activa el cálculo de Shapiro-Wilk y Kolmogorov-Smirnov."),
                ("6", "Desactivar Gráficos Innecesarios", "Cuadro: Explorar: Gráficos", "Diagramas de caja / Descriptivos", "Botones de opción", "En 'Diagramas de caja' seleccionar 'Ninguno'. En 'Descriptivos' desmarcar 'Tallo y hojas' (ahorra espacio en el visor).", "Solo se generarán las tablas numéricas solicitadas."),
                ("7", "Confirmar Gráficos", "Cuadro: Explorar: Gráficos", "Panel inferior derecho", "Botón [Continuar]", "Hacer clic en el botón [Continuar].", "Regresa al cuadro principal 'Explorar'."),
                ("8", "Configurar Visualización", "Cuadro: Explorar", "Panel inferior izquierdo 'Mostrar'", "Botón de opción (Radio button)", "Seleccionar 'Ambos' o 'Gráficos'.", "Asegura la generación de las tablas de contraste."),
                ("9", "Ejecutar Prueba", "Cuadro: Explorar", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "Aparece la tabla 'Pruebas de normalidad' con Estadístico, gl y Sig. (p-valor) de Shapiro-Wilk para cada grupo.")
            ]
        },
        {
            "code": "04_Kruskal_Wallis",
            "title": "PASO 5: PRUEBA KRUSKAL-WALLIS (NO PARAMÉTRICA - 3 O MÁS GRUPOS)",
            "desc": "Se aplica cuando al menos un grupo no presenta distribución normal (p-valor < 0.05).",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Pruebas no paramétricas > Cuadros de diálogo antiguos > K muestras independientes...", "Ir a 'Analizar' > 'Pruebas no paramétricas' > 'Cuadros de diálogo antiguos' > 'K muestras independientes...'.", "Se abre el cuadro 'Pruebas para varias muestras independientes'."),
                ("2", "Asignar Variable a Contrastar", "Cuadro: Pruebas K muestras", "Panel izquierdo", "Botón Flecha [ ► ] superior", "Seleccionar 'Tiempo' y hacer clic en la flecha [ ► ] hacia 'Lista contrastar variables'.", "'Tiempo' queda en la lista de prueba."),
                ("3", "Asignar Variable de Agrupación", "Cuadro: Pruebas K muestras", "Panel izquierdo", "Botón Flecha [ ► ] inferior", "Seleccionar 'Grupo' y hacer clic en la flecha [ ► ] hacia 'Variable de agrupación'.", "'Grupo(? ?)' queda asignado."),
                ("4", "Abrir Definición de Rango", "Cuadro: Pruebas K muestras", "Debajo de Variable de agrupación", "Botón [Definir rango...]", "Hacer clic en el botón [Definir rango...].", "Se abre la ventana emergente 'Varias muestras independientes: Definir rango'."),
                ("5", "Ingresar Rango de Grupos", "Cuadro: Definir rango", "Campos numéricos", "Teclado", "En el campo 'Mínimo' escribir '1'. En el campo 'Máximo' escribir '3' (o el número total de grupos: 4, 5, etc.).", "Rango definido: 1 a k."),
                ("6", "Confirmar Rango", "Cuadro: Definir rango", "Panel inferior", "Botón [Continuar]", "Hacer clic en el botón [Continuar].", "Regresa al cuadro principal mostrando 'Grupo(1 3)'."),
                ("7", "Verificar Tipo de Prueba", "Cuadro: Pruebas K muestras", "Sección 'Tipo de prueba'", "Casilla [X] Kruskal-Wallis H", "Verificar que esté MARCADAS las casilla: '[X] Kruskal-Wallis H'.", "Prueba Kruskal-Wallis seleccionada."),
                ("8", "Ejecutar Análisis", "Cuadro: Pruebas K muestras", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "Se generan 2 tablas: 1) 'Rangos' (N y Rango promedio) y 2) 'Estadísticos de prueba' (H de Kruskal-Wallis, gl y Sig. asintót.).")
            ]
        },
        {
            "code": "05_ANOVA_PostHoc",
            "title": "PASO 5: ANOVA DE UN FACTOR Y POST HOC (PARAMÉTRICA - 3 O MÁS GRUPOS)",
            "desc": "Se aplica cuando TODOS los grupos presentan distribución normal (p-valor >= 0.05). Incluye Levene y Tukey.",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Comparar medias > ANOVA de un factor...", "Hacer clic en 'Analizar' > 'Comparar medias' > 'ANOVA de un factor...'.", "Se abre el cuadro de diálogo 'ANOVA de un factor'."),
                ("2", "Asignar Variable Dependiente", "Cuadro: ANOVA de un factor", "Panel izquierdo", "Botón Flecha [ ► ] superior", "Seleccionar 'Tiempo' y mover con [ ► ] hacia 'Lista de dependientes'.", "'Tiempo' pasa a la lista dependiente."),
                ("3", "Asignar Factor", "Cuadro: ANOVA de un factor", "Panel izquierdo", "Botón Flecha [ ► ] inferior", "Seleccionar 'Grupo' y mover con [ ► ] hacia 'Factor'.", "'Grupo' queda asignado como factor."),
                ("4", "Configurar Opciones (Levene)", "Cuadro: ANOVA de un factor", "Panel derecho", "Botón [Opciones...]", "Hacer clic en el botón [Opciones...].", "Se abre el cuadro 'ANOVA de un factor: Opciones'."),
                ("5", "Marcar Homogeneidad y Descriptivos", "Cuadro: ANOVA: Opciones", "Sección 'Estadísticas'", "Casillas de verificación", "Marcar: '[X] Descriptivos' y '[X] Prueba de homogeneidad de las varianzas'.", "Permite obtener la prueba de Levene para verificar supuestos de varianza."),
                ("6", "Confirmar Opciones", "Cuadro: ANOVA: Opciones", "Panel inferior", "Botón [Continuar]", "Hacer clic en el botón [Continuar].", "Regresa al cuadro principal de ANOVA."),
                ("7", "Configurar Post Hoc", "Cuadro: ANOVA de un factor", "Panel derecho", "Botón [Post Hoc...]", "Hacer clic en el botón [Post Hoc...].", "Se abre el cuadro 'ANOVA de un factor: Comparaciones múltiples post hoc'."),
                ("8", "Elegir Prueba Post Hoc", "Cuadro: Post Hoc", "Secciones de pruebas", "Casillas [X] Tukey / Games-Howell", "Si varianzas son iguales (Levene p>=0.05): Marcar '[X] Tukey'. Si varianzas NO son iguales (Levene p<0.05): Marcar '[X] Games-Howell'. Verificar 'Nivel de significación: 0.05'.", "Post Hoc configurado según resultado de Levene."),
                ("9", "Confirmar Post Hoc", "Cuadro: Post Hoc", "Panel inferior", "Botón [Continuar]", "Hacer clic en el botón [Continuar].", "Regresa al cuadro principal."),
                ("10", "Ejecutar ANOVA", "Cuadro: ANOVA de un factor", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "Se obtienen las tablas: 'Descriptivos', 'Prueba de homogeneidad de varianzas' (Levene), 'ANOVA' (F y Sig.) y 'Comparaciones múltiples' (Post Hoc).")
            ]
        },
        {
            "code": "06_Mann_Whitney",
            "title": "PASO 5: U DE MANN-WHITNEY (NO PARAMÉTRICA - 2 GRUPOS INDEPENDIENTES)",
            "desc": "Comparación de 2 grupos independientes cuando no se cumple el supuesto de normalidad.",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Pruebas no paramétricas > Cuadros de diálogo antiguos > 2 muestras independientes...", "Ir a 'Analizar' > 'Pruebas no paramétricas' > 'Cuadros de diálogo antiguos' > '2 muestras independientes...'.", "Se abre el cuadro 'Pruebas para dos muestras independientes'."),
                ("2", "Asignar Variable de Prueba", "Cuadro: 2 muestras indep.", "Panel izquierdo", "Botón Flecha [ ► ] superior", "Seleccionar la variable numérica y hacer clic en la flecha [ ► ] hacia 'Lista contrastar variables'.", "Variable numérica asignada."),
                ("3", "Asignar Variable de Grupo", "Cuadro: 2 muestras indep.", "Panel izquierdo", "Botón Flecha [ ► ] inferior", "Seleccionar la variable categórica de 2 niveles y mover con [ ► ] hacia 'Variable de agrupación'.", "Variable de agrupación asignada."),
                ("4", "Abrir Definición de Grupos", "Cuadro: 2 muestras indep.", "Debajo de Variable de agrupación", "Botón [Definir grupos...]", "Hacer clic en el botón [Definir grupos...].", "Se abre el cuadro emergente 'Definir grupos'."),
                ("5", "Digitar Códigos de Grupo", "Cuadro: Definir grupos", "Campos 'Grupo 1' y 'Grupo 2'", "Teclado", "En el campo 'Grupo 1' escribir '1'. En el campo 'Grupo 2' escribir '2'.", "Grupos identificados."),
                ("6", "Confirmar Grupos", "Cuadro: Definir grupos", "Panel inferior", "Botón [Continuar]", "Hacer clic en el botón [Continuar].", "Regresa al cuadro principal mostrando '(1 2)'."),
                ("7", "Marcar Tipo de Prueba", "Cuadro: 2 muestras indep.", "Sección 'Tipo de prueba'", "Casilla [X] U de Mann-Whitney", "Verificar que esté marcada la casilla '[X] U de Mann-Whitney'.", "Prueba no paramétrica activada."),
                ("8", "Ejecutar Análisis", "Cuadro: 2 muestras indep.", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "Se generan las tablas: 'Rangos' (Rango promedio y Suma de rangos) y 'Estadísticos de prueba' (U de Mann-Whitney, W de Wilcoxon, Z, Sig. asintót. bilateral).")
            ]
        },
        {
            "code": "07_T_Student",
            "title": "PASO 5: T DE STUDENT (PARAMÉTRICA - 2 GRUPOS INDEPENDIENTES)",
            "desc": "Comparación de 2 grupos independientes cuando ambos presentan distribución normal.",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Comparar medias > Prueba T para muestras independientes...", "Ir a 'Analizar' > 'Comparar medias' > 'Prueba T para muestras independientes...'.", "Se abre el cuadro 'Prueba T para muestras independientes'."),
                ("2", "Asignar Variable de Prueba", "Cuadro: Prueba T indep.", "Panel izquierdo", "Botón Flecha [ ► ] superior", "Seleccionar la variable numérica y mover con [ ► ] hacia 'Variables de prueba'.", "Variable asignada."),
                ("3", "Asignar Variable de Agrupación", "Cuadro: Prueba T indep.", "Panel izquierdo", "Botón Flecha [ ► ] inferior", "Seleccionar la variable de grupo y mover con [ ► ] hacia 'Variable de agrupación'.", "Variable de grupo asignada."),
                ("4", "Definir Grupos", "Cuadro: Prueba T indep.", "Debajo de agrupación", "Botón [Definir grupos...]", "Hacer clic en [Definir grupos...], ingresar '1' en Grupo 1 y '2' en Grupo 2 -> Clic en [Continuar].", "Grupos 1 y 2 definidos."),
                ("5", "Ejecutar Prueba T", "Cuadro: Prueba T indep.", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "Se generan 2 tablas: 'Estadísticas de grupo' y 'Prueba de muestras independientes' (con prueba de Levene y valor t con su p-valor Sig.).")
            ]
        },
        {
            "code": "08_Muestras_Pareadas",
            "title": "PRUEBA T PARA MUESTRAS RELACIONADAS (ANTES - DESPUÉS / PRE - POST)",
            "desc": "Comparación de dos mediciones tomadas sobre el mismo sujeto o unidad de análisis.",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Comparar medias > Prueba T para muestras emparejadas...", "Ir a 'Analizar' > 'Comparar medias' > 'Prueba T para muestras emparejadas...'.", "Se abre el cuadro 'Prueba T para muestras emparejadas'."),
                ("2", "Seleccionar Par de Variables", "Cuadro: Muestras emparejadas", "Panel izquierdo", "Clic en Variable 1 y Variable 2", "Hacer clic primero sobre la variable 'Antes' (o Pretest) y luego sobre 'Después' (o Postest).", "Ambas variables quedan seleccionadas simultáneamente."),
                ("3", "Mover el Par", "Cuadro: Muestras emparejadas", "Entre ambos paneles", "Botón Flecha [ ► ] central", "Hacer clic en la flecha [ ► ] central apuntando hacia 'Variables emparejadas'.", "Aparece en 'Par 1': Variable1 - Variable2."),
                ("4", "Ejecutar Análisis", "Cuadro: Muestras emparejadas", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "Se generan 3 tablas: 'Estadísticas de muestras emparejadas', 'Correlaciones de muestras emparejadas' y 'Prueba de muestras emparejadas' (t, gl y Sig. bilateral).")
            ]
        },
        {
            "code": "09_Chi_Cuadrado",
            "title": "PRUEBA CHI-CUADRADO DE INDEPENDENCIA Y HOMOGENEIDAD",
            "desc": "Evaluación de asociación entre 2 variables categóricas o cualitativas mediante Tablas Cruzadas.",
            "steps": [
                ("1", "Acceso al Menú", "Barra Superior > Analizar", "Menú Desplegable", "Estadísticos descriptivos > Tablas cruzadas...", "Ir a 'Analizar' > 'Estadísticos descriptivos' > 'Tablas cruzadas...'.", "Se abre el cuadro de diálogo 'Tablas cruzadas'."),
                ("2", "Asignar Filas y Columnas", "Cuadro: Tablas cruzadas", "Panel izquierdo", "Botones Flecha [ ► ]", "Seleccionar la primera variable cualitativa y mover a 'Filas'. Seleccionar la segunda variable y mover a 'Columnas'.", "Variables ubicadas en filas y columnas de la tabla de contingencia."),
                ("3", "Activar Prueba Chi-Cuadrado", "Cuadro: Tablas cruzadas", "Panel derecho", "Botón [Estadísticos...]", "Hacer clic en el botón [Estadísticos...]. En la ventana emergente, MARCAR la casilla: '[X] Chi-cuadrado' -> Clic en [Continuar].", "Prueba Chi-cuadrado activada."),
                ("4", "Configurar Celdas (Esperadas)", "Cuadro: Tablas cruzadas", "Panel derecho", "Botón [Casillas...]", "Hacer clic en el botón [Casillas...]. En 'Recuentos', marcar '[X] Observadas' y '[X] Esperadas' -> Clic en [Continuar].", "Permite contrastar valores observados vs esperados."),
                ("5", "Ejecutar Análisis", "Cuadro: Tablas cruzadas", "Panel derecho", "Botón [Aceptar]", "Hacer clic en el botón [Aceptar].", "Se generan las tablas: 'Tabla de contingencia' y 'Pruebas de chi-cuadrado' (mostrando el Chi-cuadrado de Pearson, gl y Sig. asintótica bilateral).")
            ]
        },
        {
            "code": "10_Guardar_SAV_SPV",
            "title": "GUARDADO OFICIAL DE ARCHIVOS (.SAV Y .SPV) PARA ENTREGA",
            "desc": "Procedimiento exacto para guardar la base de datos (.sav) y el visor de resultados (.spv) requeridos por el docente.",
            "steps": [
                ("1", "Guardar Datos (.SAV)", "Ventana del Editor de Datos SPSS", "Menú Superior: Archivo > Guardar como...", "Opción 'Guardar como...'", "Hacer clic en 'Archivo' > 'Guardar como...'.", "Se abre el cuadro de guardado de Windows."),
                ("2", "Nombre del Archivo SAV", "Cuadro: Guardar datos como", "Campo 'Nombre de archivo'", "Teclado", "Escribir el formato oficial: 'Ejercicio01_GF.sav' (o 'Ejercicio02_GF.sav', con tus iniciales).", "Nombre estandarizado asignado."),
                ("3", "Verificar Tipo SAV", "Cuadro: Guardar datos como", "Menú 'Guardar como tipo'", "Menú desplegable", "Asegurarse de que esté seleccionado: 'SPSS Statistics (*.sav)'.", "Formato nativo de datos SPSS asegurado."),
                ("4", "Confirmar Guardado SAV", "Cuadro: Guardar datos como", "Panel inferior derecho", "Botón [Guardar]", "Hacer clic en el botón [Guardar].", "Se crea el archivo .sav visible en el explorador de archivos."),
                ("5", "Pasar al Visor de Resultados", "Barra de tareas de Windows", "Ventana: IBM SPSS Statistics Viewer", "Clic en la ventana de Visor", "Ir a la ventana donde se generaron las tablas y salidas estadísticas (Visor).", "Ventana de resultados activa."),
                ("6", "Guardar Resultados (.SPV)", "Ventana del Visor de Resultados", "Menú Superior: Archivo > Guardar como...", "Opción 'Guardar como...'", "Hacer clic en 'Archivo' > 'Guardar como...'.", "Se abre el cuadro de guardado del visor."),
                ("7", "Nombre del Archivo SPV", "Cuadro: Guardar resultados como", "Campo 'Nombre de archivo'", "Teclado", "Escribir el formato oficial: 'Ejercicio01.spv' (o 'Ejercicio02.spv', según corresponda).", "Nombre asignado."),
                ("8", "Verificar Tipo SPV", "Cuadro: Guardar resultados como", "Menú 'Guardar como tipo'", "Menú desplegable", "Asegurarse de que esté seleccionado: 'Visor de SPSS Statistics (*.spv)'.", "Formato nativo de salidas SPSS asegurado."),
                ("9", "Confirmar Guardado SPV", "Cuadro: Guardar resultados como", "Panel inferior derecho", "Botón [Guardar]", "Hacer clic en el botón [Guardar].", "Se genera el archivo .spv con todas las tablas originales para presentación.")
            ]
        }
    ]

    for s_info in sheets_info:
        ws = wb.create_sheet(title=s_info["code"])
        ws.views.sheetView[0].showGridLines = True

        # Título
        ws.merge_cells("A1:G1")
        ws["A1"] = s_info["title"]
        ws["A1"].font = font_title
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells("A2:G2")
        ws["A2"] = s_info["desc"]
        ws["A2"].font = font_subtitle
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

        # Cabecera de columnas
        for c_idx, h_text in enumerate(headers, 1):
            cell = ws.cell(row=4, column=c_idx, value=h_text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_header
            cell.border = thin_border
        
        ws.row_dimensions[4].height = 28

        # Filas de datos
        for r_idx, step_data in enumerate(s_info["steps"], 5):
            fill_curr = fill_zebra if r_idx % 2 == 1 else fill_white
            ws.row_dimensions[r_idx].height = 42

            for c_idx, val in enumerate(step_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.fill = fill_curr
                cell.border = thin_border

                if c_idx == 1:
                    cell.font = font_bold
                    cell.alignment = align_center
                elif c_idx == 5:
                    # Columna de botón o pestaña específica
                    cell.font = font_bold
                    cell.alignment = align_left
                    cell.fill = fill_btn
                elif c_idx in [2, 3, 4]:
                    cell.font = font_bold if c_idx == 2 else font_body
                    cell.alignment = align_left
                else:
                    cell.font = font_body
                    cell.alignment = align_left

        # Ancho de columnas
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(filename)
    print(f"Archivo generado exitosamente: {filename}")

if __name__ == "__main__":
    build_spss_excel_guide()
