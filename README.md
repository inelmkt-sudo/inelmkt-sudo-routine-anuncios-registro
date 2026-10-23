# routine-anuncios-registro

Registra cada semana, **por anuncio**, el gasto y los leads de Meta Ads en la hoja `ANUNCIOS` del Excel Maestro de Marketing, con la temperatura (caliente/tibio/frío) de cada lead. `GASTOS` suma desde `ANUNCIOS` con fórmulas: el gasto de Meta se escribe una sola vez.

Responde: **"este anuncio está caro, pero ¿trae calientes o no?"** (columna `Veredicto` de ANUNCIOS).

## Qué hace
1. Lee los códigos de **Intake 3** en `PROGRAMAS`.
2. Meta Ads (solo lectura): gasto y resultados por anuncio, semana anterior y semana en curso.
3. Google Drive: leads por `ad_id` desde el Sheet de cada código.
4. Consolidado de leads de AMG de cada programa (drive de TI, `Intake 3/<código>/02_AMG/Consolidado Leads`): temperatura (columna TIPO) cruzada por **correo**.
5. Corte semanal en `LEADS` por código: leads, calientes y tibios (del consolidado de AMG) y atendidos ≤24h (de la matriz de atención `03_AGC`). Ventas y orgánicos no se escriben.
6. Upsert en `ANUNCIOS` (Semana + ID anuncio) y en `GASTOS` (Semana + ID campaña, con fórmulas SUMIFS).

## Archivo
- `EXCEL_MAESTRO_MKT_v2.xlsx` — `09. Marketing/INSTITUTE/REGISTRO` (OneDrive natalieaguirre)
- Drive ID / Item ID: ver `.env.example`

## Configuración del Routine (claude.ai/code/routines)
- **Environment**: `bash setup.sh` · **Network**: Full
- **Connectors**: Meta ADS + Composio (Excel/OneDrive + Google Sheets activos) + Google Drive
- **Trigger**: Weekly, lunes 09:00 (Lima)
- **Prompt**:
```
Lee CLAUDE.md y ejecuta la automatización descrita ahí. Operas con autonomía
plena: no preguntas, no pides confirmación, no esperas input. Si algo es
genuinamente imposible, detente con nota clara y exit error. Nunca escribas en
Meta Ads. Reporta resumen al final.
```
