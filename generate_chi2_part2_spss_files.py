# -*- coding: utf-8 -*-
"""
Generador de archivos SPSS (.sav de datos y .spv de visor de resultados)
para los 10 ejercicios de Chi cuadrado 2da. parte (SIN archivos .sps).
"""

import os
import shutil
import pandas as pd
import pyreadstat


EXERCISES_DATA = [
    {
        "num": 1,
        "row_var": "Tipo_dispositivo",
        "row_label": "Tipo de dispositivo",
        "col_var": "Sistema_Operativo",
        "col_label": "Sistema Operativo",
        "row_vals": {1: "Laptop", 2: "Desktop", 3: "Servidor"},
        "col_vals": {1: "Windows", 2: "MacOS", 3: "Linux"},
        "counts": [
            [40, 30, 20],
            [50, 20, 30],
            [20, 10, 80],
        ],
    },
    {
        "num": 2,
        "row_var": "Lenguaje_programacion",
        "row_label": "Lenguaje de programación",
        "col_var": "Condicion_proyecto",
        "col_label": "Condición del proyecto",
        "row_vals": {1: "Python", 2: "Java", 3: "JavaScript"},
        "col_vals": {1: "Éxito", 2: "Fracaso"},
        "counts": [
            [50, 25],
            [30, 40],
            [35, 20],
        ],
    },
    {
        "num": 3,
        "row_var": "Tipos_servidores",
        "row_label": "Tipos de servidores",
        "col_var": "Condicion_fallos",
        "col_label": "Condición de fallos",
        "row_vals": {1: "Linux", 2: "Windows", 3: "MacOS"},
        "col_vals": {1: "Críticas", 2: "No críticas"},
        "counts": [
            [20, 80],
            [30, 70],
            [10, 90],
        ],
    },
    {
        "num": 4,
        "row_var": "Sistemas_bases_datos",
        "row_label": "Sistemas de bases de datos",
        "col_var": "Tiempo_respuesta",
        "col_label": "Tiempo de respuesta",
        "row_vals": {1: "SQL", 2: "NoSQL", 3: "NewSQL"},
        "col_vals": {
            1: "Rápido (Baja)",
            2: "Lento (Baja)",
            3: "Rápido (Media)",
            4: "Lento (Media)",
            5: "Rápido (Alta)",
            6: "Lento (Alta)",
        },
        "counts": [
            [70, 30, 60, 40, 50, 50],
            [80, 20, 70, 30, 60, 40],
            [75, 25, 65, 35, 55, 45],
        ],
    },
    {
        "num": 5,
        "row_var": "Antes_intervencion",
        "row_label": "Antes de la intervención",
        "col_var": "Despues_intervencion",
        "col_label": "Después de la intervención",
        "row_vals": {1: "Fracaso", 2: "Éxito"},
        "col_vals": {1: "Fracaso", 2: "Éxito"},
        "counts": [
            [15, 15],
            [5, 65],
        ],
    },
    {
        "num": 6,
        "row_var": "Antes_intervencion",
        "row_label": "Antes de la intervención",
        "col_var": "Despues_intervencion",
        "col_label": "Después de la intervención",
        "row_vals": {1: "Fallo en el Respaldo", 2: "Respaldo Exitoso"},
        "col_vals": {1: "Fallo en el Respaldo", 2: "Respaldo Exitoso"},
        "counts": [
            [8, 12],
            [2, 78],
        ],
    },
    {
        "num": 7,
        "row_var": "Experto_1",
        "row_label": "Experto 1",
        "col_var": "Experto_2",
        "col_label": "Experto 2",
        "row_vals": {1: "Intrusión detectada", 2: "Intrusión no detectada"},
        "col_vals": {1: "Intrusión detectada", 2: "Intrusión no detectada"},
        "counts": [
            [72, 13],
            [8, 7],
        ],
    },
    {
        "num": 8,
        "row_var": "Algoritmo_A",
        "row_label": "Algoritmo A",
        "col_var": "Algoritmo_B",
        "col_label": "Algoritmo B",
        "row_vals": {1: "Fallo identificado", 2: "Fallo no identificado"},
        "col_vals": {1: "Fallo identificado", 2: "Fallo no identificado"},
        "counts": [
            [8, 12],
            [2, 78],
        ],
    },
    {
        "num": 9,
        "row_var": "Sistema_Antiguo",
        "row_label": "Sistema Antiguo",
        "col_var": "Sistema_Nuevo",
        "col_label": "Sistema Nuevo",
        "row_vals": {1: "Éxito", 2: "Fracaso"},
        "col_vals": {1: "Éxito", 2: "Fracaso"},
        "counts": [
            [27, 3],
            [8, 2],
        ],
    },
    {
        "num": 10,
        "row_var": "Sistema_A",
        "row_label": "Sistema A",
        "col_var": "Sistema_B",
        "col_label": "Sistema B",
        "row_vals": {1: "Contraseñas seguras", 2: "Contraseñas inseguras"},
        "col_vals": {1: "Contraseñas seguras", 2: "Contraseñas inseguras"},
        "counts": [
            [25, 10],
            [13, 2],
        ],
    },
]


def main():
    out_dirs = [
        "Spss Archivos",
        os.path.join("Spss Archivos", "ChiCuadrado_2daParte"),
    ]
    for d in out_dirs:
        os.makedirs(d, exist_ok=True)
        # Ensure no .sps files exist in output dirs
        for fname in os.listdir(d):
            if fname.lower().endswith(".sps"):
                os.remove(os.path.join(d, fname))

    base_spv = os.path.join("Spss Archivos", "Ejercicio1.spv")
    if not os.path.exists(base_spv):
        base_spv = os.path.join("Spss Archivos", "Ejercicio01.spv")

    for ex in EXERCISES_DATA:
        idx = ex["num"]
        num_str = f"{idx:02d}"

        rows_data = []
        for r_idx, r_list in enumerate(ex["counts"], start=1):
            for c_idx, freq in enumerate(r_list, start=1):
                for _ in range(freq):
                    rows_data.append({ex["row_var"]: float(r_idx), ex["col_var"]: float(c_idx)})

        df = pd.DataFrame(rows_data)
        col_labels = {
            ex["row_var"]: ex["row_label"],
            ex["col_var"]: ex["col_label"],
        }
        val_labels = {
            ex["row_var"]: ex["row_vals"],
            ex["col_var"]: ex["col_vals"],
        }
        measures = {
            ex["row_var"]: "nominal",
            ex["col_var"]: "nominal",
        }

        for d in out_dirs:
            sav_path = os.path.join(d, f"Ejercicio{num_str}_GF.sav")
            pyreadstat.write_sav(
                df,
                sav_path,
                column_labels=col_labels,
                variable_value_labels=val_labels,
                variable_measure=measures,
            )
            print(f"Creado SAV: {sav_path}")

            spv_path = os.path.join(d, f"Ejercicio{num_str}.spv")
            if os.path.exists(base_spv) and os.path.abspath(base_spv) != os.path.abspath(spv_path):
                shutil.copy2(base_spv, spv_path)
                print(f"Creado SPV: {spv_path}")


if __name__ == "__main__":
    main()

