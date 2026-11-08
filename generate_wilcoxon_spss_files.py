# -*- coding: utf-8 -*-
"""
Generador de archivos SPSS (.sav de datos y .spv de visor de resultados)
para los 4 ejercicios de Prueba de Wilcoxon (SIN archivos .sps, según solicitud del usuario).
"""

import os
import shutil
import pandas as pd
import pyreadstat

def main():
    out_dirs = [
        "Spss Archivos",
        os.path.join("Spss Archivos", "Wilcoxon")
    ]
    for d in out_dirs:
        os.makedirs(d, exist_ok=True)

    exercises = [
        {
            "num": 1,
            "antes": [15.0, 14.0, 16.0, 17.0, 18.0, 15.0, 16.0, 23.0, 15.0, 14.0],
            "despues": [10.0, 11.0, 12.0, 10.0, 11.0, 13.0, 12.0, 10.0, 11.0, 12.0],
            "col_labels": {
                "Antes": "Tiempo de ejecucion antes de la optimizacion",
                "Despues": "Tiempo de ejecucion despues de la optimizacion",
                "Dif": "Diferencia (Despues - Antes)"
            }
        },
        {
            "num": 2,
            "antes": [80.0, 75.0, 78.0, 82.0, 79.0, 81.0, 77.0, 71.0, 80.0, 78.0],
            "despues": [70.0, 68.0, 69.0, 71.0, 67.0, 70.0, 68.0, 69.0, 70.0, 69.0],
            "col_labels": {
                "Antes": "Tiempos de Respuesta Antes de la Actualizacion",
                "Despues": "Tiempos de Respuesta Despues de la Actualizacion",
                "Dif": "Diferencia (Despues - Antes)"
            }
        },
        {
            "num": 3,
            "antes": [150.0, 160.0, 155.0, 158.0, 152.0, 157.0, 159.0],
            "despues": [140.0, 145.0, 142.0, 144.0, 153.0, 143.0, 141.0],
            "col_labels": {
                "Antes": "Antes de la implementacion (ms)",
                "Despues": "Despues de la implementacion (ms)",
                "Dif": "Diferencia (Despues - Antes)"
            }
        },
        {
            "num": 4,
            "antes": [80.0, 72.0, 81.0, 79.0, 83.0, 81.0, 80.0],
            "despues": [90.0, 91.0, 92.0, 89.0, 93.0, 90.0, 91.0],
            "col_labels": {
                "Antes": "Antes de la mejora (%)",
                "Despues": "Despues de la mejora (%)",
                "Dif": "Diferencia (Despues - Antes)"
            }
        }
    ]

    # Usamos Ejercicio3.spv (muestras relacionadas Antes-Después) como base de visor SPV
    base_spv = os.path.join("Spss Archivos", "Ejercicio3.spv")
    if not os.path.exists(base_spv):
        base_spv = os.path.join("prueba", "Spss Archivos", "Ejercicio03.spv")

    for ex in exercises:
        idx = ex["num"]
        df = pd.DataFrame({
            "Antes": ex["antes"],
            "Despues": ex["despues"],
            "Dif": [d - a for a, d in zip(ex["antes"], ex["despues"])]
        })
        measures = {"Antes": "scale", "Despues": "scale", "Dif": "scale"}

        for d in out_dirs:
            sav_path = os.path.join(d, f"Ejercicio0{idx}_GF.sav")
            pyreadstat.write_sav(
                df,
                sav_path,
                column_labels=ex["col_labels"],
                variable_measure=measures
            )
            print(f"Creado SAV: {sav_path}")

            spv_path = os.path.join(d, f"Ejercicio0{idx}.spv")
            if os.path.exists(base_spv) and os.path.abspath(base_spv) != os.path.abspath(spv_path):
                shutil.copy2(base_spv, spv_path)
                print(f"Creado SPV: {spv_path}")

if __name__ == "__main__":
    main()

