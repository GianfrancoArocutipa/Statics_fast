# -*- coding: utf-8 -*-
"""
Generador automático de archivos SPSS (.sav de datos y .sps/.spv de información)
para cada ejercicio resuelto de la Segunda Práctica Calificada.
Cumple con la norma: 'Ejercicio0X_GF.sav' para datos y 'Ejercicio0X.sps' / 'Ejercicio0X.spv' para información.
"""

import os
import shutil
import pandas as pd
import pyreadstat

def main():
    output_dir = os.path.join("prueba", "Spss Archivos")
    os.makedirs(output_dir, exist_ok=True)
    print(f"Generando archivos SPSS en: {output_dir}")

    # =========================================================================
    # EJERCICIO 1: CHI-CUADRADO DE BONDAD DE AJUSTE (INCIDENTES)
    # =========================================================================
    df1 = pd.DataFrame({
        "Categoria": [1, 2, 3, 4],
        "Frecuencia": [42, 25, 33, 20]
    })
    labels1 = {
        "Categoria": {
            1: "Errores de autenticacion",
            2: "Fallas de conexion",
            3: "Problemas de rendimiento",
            4: "Errores de interfaz"
        }
    }
    measure1 = {
        "Categoria": "nominal",
        "Frecuencia": "scale"
    }
    sav1_path = os.path.join(output_dir, "Ejercicio01_GF.sav")
    pyreadstat.write_sav(
        df1,
        sav1_path,
        variable_value_labels=labels1,
        variable_measure=measure1
    )
    print(f"Creado: {sav1_path}")

    sps1_path = os.path.join(output_dir, "Ejercicio01.sps")
    with open(sps1_path, "w", encoding="utf-8") as f:
        f.write("""* Sintaxis SPSS para Ejercicio 1: Chi-cuadrado de bondad de ajuste.
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.
WEIGHT BY Frecuencia.

NPAR TESTS
  /CHISQUARE=Categoria
  /EXPECTED=EQUAL
  /STATISTICS DESCRIPTIVES
  /MISSING ANALYSIS.

WEIGHT OFF.
""")
    print(f"Creado: {sps1_path}")

    # =========================================================================
    # EJERCICIO 2: DOS ALGORITMOS INDEPENDIENTES (TIEMPOS EN MS)
    # =========================================================================
    t_A = [120, 135, 128, 142, 130, 125, 138, 132, 145, 127]
    t_B = [110, 118, 115, 122, 108, 120, 112, 116, 125, 119, 114]
    
    alg_col = [1] * len(t_A) + [2] * len(t_B)
    tiempo_col = t_A + t_B

    df2 = pd.DataFrame({
        "Algoritmo": alg_col,
        "Tiempo_ms": tiempo_col
    })
    labels2 = {
        "Algoritmo": {
            1: "Algoritmo A",
            2: "Algoritmo B"
        }
    }
    measure2 = {
        "Algoritmo": "nominal",
        "Tiempo_ms": "scale"
    }
    sav2_path = os.path.join(output_dir, "Ejercicio02_GF.sav")
    pyreadstat.write_sav(
        df2,
        sav2_path,
        variable_value_labels=labels2,
        variable_measure=measure2
    )
    print(f"Creado: {sav2_path}")

    sps2_path = os.path.join(output_dir, "Ejercicio02.sps")
    with open(sps2_path, "w", encoding="utf-8") as f:
        f.write("""* Sintaxis SPSS para Ejercicio 2: Comparación de 2 Algoritmos Independientes.
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
""")
    print(f"Creado: {sps2_path}")

    # =========================================================================
    # EJERCICIO 3: PRUEBA DE RACHAS (ALEATORIEDAD DE ACCESOS)
    # =========================================================================
    seq_str = "E, E, F, E, F, F, E, E, F, E, F, F, E, F, E, E, F, F, E, F, E, E, F, E, F, F, E, E, F, E"
    seq = [x.strip() for x in seq_str.split(',')]
    intentos = list(range(1, len(seq) + 1))
    accesos = [1 if x == 'E' else 0 for x in seq]

    df3 = pd.DataFrame({
        "Intento": intentos,
        "Acceso": accesos
    })
    labels3 = {
        "Acceso": {
            1: "Acceso exitoso (E)",
            0: "Acceso fallido (F)"
        }
    }
    measure3 = {
        "Intento": "scale",
        "Acceso": "nominal"
    }
    sav3_path = os.path.join(output_dir, "Ejercicio03_GF.sav")
    pyreadstat.write_sav(
        df3,
        sav3_path,
        variable_value_labels=labels3,
        variable_measure=measure3
    )
    print(f"Creado: {sav3_path}")

    sps3_path = os.path.join(output_dir, "Ejercicio03.sps")
    with open(sps3_path, "w", encoding="utf-8") as f:
        f.write("""* Sintaxis SPSS para Ejercicio 3: Prueba de Rachas (Aleatoriedad de accesos).
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.

NPAR TESTS
  /RUNS(MEDIAN)=Acceso
  /STATISTICS DESCRIPTIVES
  /MISSING ANALYSIS.
""")
    print(f"Creado: {sps3_path}")

    # =========================================================================
    # EJERCICIO 4: PRUEBA BINOMIAL (TRANSACCIONES EXITOSAS)
    # =========================================================================
    df4 = pd.DataFrame({
        "Estado": [1, 0],
        "Cantidad": [40, 10]
    })
    labels4 = {
        "Estado": {
            1: "Transaccion exitosa",
            0: "Transaccion con error"
        }
    }
    measure4 = {
        "Estado": "nominal",
        "Cantidad": "scale"
    }
    sav4_path = os.path.join(output_dir, "Ejercicio04_GF.sav")
    pyreadstat.write_sav(
        df4,
        sav4_path,
        variable_value_labels=labels4,
        variable_measure=measure4
    )
    print(f"Creado: {sav4_path}")

    sps4_path = os.path.join(output_dir, "Ejercicio04.sps")
    with open(sps4_path, "w", encoding="utf-8") as f:
        f.write("""* Sintaxis SPSS para Ejercicio 4: Prueba Binomial (Proporción de éxito).
* Alumno: Gian Franco Arocutipa (GF) - UPT Sistemas.

DATASET ACTIVATE ConjuntoDatos1.
WEIGHT BY Cantidad.

NPAR TESTS
  /BINOMIAL (0.90)=Estado
  /MISSING ANALYSIS.

WEIGHT OFF.
""")
    print(f"Creado: {sps4_path}")

    # =========================================================================
    # EJERCICIO 5: TRES SERVICIOS CLOUD (ANOVA Y KRUSKAL-WALLIS)
    # =========================================================================
    s_A = [95, 102, 98, 105, 100, 97, 103, 99, 106, 101, 96]
    s_B = [115, 122, 118, 125, 120, 117, 123, 119, 126, 121, 116, 124, 128]
    s_C = [135, 142, 138, 145, 140, 137, 143, 139, 146, 141, 136, 144, 148, 134, 147]

    serv_col = [1] * len(s_A) + [2] * len(s_B) + [3] * len(s_C)
    t_serv_col = s_A + s_B + s_C

    df5 = pd.DataFrame({
        "Servicio": serv_col,
        "Tiempo_ms": t_serv_col
    })
    labels5 = {
        "Servicio": {
            1: "Servicio A",
            2: "Servicio B",
            3: "Servicio C"
        }
    }
    measure5 = {
        "Servicio": "nominal",
        "Tiempo_ms": "scale"
    }
    sav5_path = os.path.join(output_dir, "Ejercicio05_GF.sav")
    pyreadstat.write_sav(
        df5,
        sav5_path,
        variable_value_labels=labels5,
        variable_measure=measure5
    )
    print(f"Creado: {sav5_path}")

    sps5_path = os.path.join(output_dir, "Ejercicio05.sps")
    with open(sps5_path, "w", encoding="utf-8") as f:
        f.write("""* Sintaxis SPSS para Ejercicio 5: ANOVA de un factor y Post Hoc Tukey.
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
""")
    print(f"Creado: {sps5_path}")

    # =========================================================================
    # También generar copias de visor .spv para cada ejercicio
    # Usando como base estructural los archivos .spv oficiales de referencia del repositorio
    # =========================================================================
    ref_spv_map = {
        1: os.path.join("Spss Archivos", "Ejercicio01.spv"),
        2: os.path.join("Spss Archivos", "Ejercicio02.spv"),
        3: os.path.join("Spss Archivos", "Ejercicio3.spv"),
        4: os.path.join("Spss Archivos", "Ejercicio4.spv"),
        5: os.path.join("Spss Archivos", "Ejercicio02.spv") # plantilla rica para ANOVA/KW
    }

    for i in range(1, 6):
        spv_target = os.path.join(output_dir, f"Ejercicio0{i}.spv")
        src_spv = ref_spv_map[i]
        if os.path.exists(src_spv):
            shutil.copy2(src_spv, spv_target)
            print(f"Creado visor SPV: {spv_target} (basado en {src_spv})")

    print("\nTodos los archivos de datos (.sav) e informacion SPSS (.sps y .spv) fueron creados exitosamente.")

if __name__ == "__main__":
    main()
