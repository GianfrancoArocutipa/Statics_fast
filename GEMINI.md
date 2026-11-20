# GEMINI.md - Directrices de Asistencia para Estadística Inferencial

Sigue estrictamente las directrices y metodología documentadas en [AGENTS.md](AGENTS.md) y [.agents/rules/metodologia_resolucion_estadistica.md](.agents/rules/metodologia_resolucion_estadistica.md).

### Resumen Rápido:
1. **Comando "prepárate":** Si el usuario escribe "preparate" o "prepárate", ejecuta de inmediato:
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```
2. **Formato UPT Sistemas:** Respetar los 7 u 8 pasos del docente (Dr. Luis A. Fernández Vizcarra).
3. **Tablas SPSS:** Usar capturas renderizadas con estética oficial de IBM SPSS Statistics 26-29.
4. **Conclusiones Breves y Conclusión Administrativa (PASO 7):**
   - En pruebas numéricas (ANOVA / Kruskal-Wallis / Mann-Whitney / Wilcoxon), redactar una sola frase concluyente directa sin expandirse.
   - En pruebas con **PASO 7: Conclusión administrativa** (Chi cuadrado, McNemar, Kappa, Fisher), usar estrictamente los dos modelos del docente según el contexto:
     - **Si son iguales ($p > \alpha$):** *"Dado que los [N] [sistemas/grupos], con respecto a [variable] son iguales, se podría sugerir adquirir cualquiera de estos sistemas. El factor diferenciador se podría estar considerando: 1.Costos / 2.Garantias / 3.Interaccion"*
     - **Si hay efectividad/mejora/asociación ($p \le \alpha$):** *"Es favorable implementar [el sistema/opción], porque a mostrado efectividad en los datos mostrados y en sus respectivas pruebas, uno de los factores fue [la reducción de porcentajes de errores (X%) / mayor porcentaje de éxito o coincidencia (X%)]."*
5. **Entrega Dual:** Generar y enlazar siempre tanto el `.docx` editable como el `.pdf` compilado.
6. **Preservación Total de Formato:** Mantener siempre todo el formato original (colores, fuentes, bordes, encabezados, diseño) del documento `.docx` recibido; solo actualizar la información.
7. **Archivos SPSS por Ejercicio (.sav y .spv):** Por cada ejercicio resuelto, generar obligatoriamente su archivo de datos (`Ejercicio0X_GF.sav`) y su archivo de información de salida/visor SPSS (`Ejercicio0X.spv` / `.sps`).
8. **Circuitos en Tinkercad:** Si el usuario envía un documento para armar un circuito en Tinkercad, entregar obligatoriamente: (a) lista de componentes exactos en Tinkercad, (b) guía paso a paso de conexiones pin por pin / protoboard, y (c) código C++ (Arduino) 100% funcional y compatible con el simulador de Tinkercad (además de completar el documento si lo requiere).


