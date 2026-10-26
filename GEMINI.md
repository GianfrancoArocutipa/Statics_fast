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
4. **Conclusiones Breves:** En el Paso 7/8 no expandirse en explicaciones largas; redactar una sola frase concluyente siguiendo el modelo de los Ejercicios 1 y 2.
5. **Entrega Dual:** Generar y enlazar siempre tanto el `.docx` editable como el `.pdf` compilado.
