# Routine: Registro semanal por anuncio → ANUNCIOS + GASTOS (Excel Maestro MKT)

## Objetivo
Registrar, una fila por **anuncio × semana**, el gasto y los resultados de Meta Ads de los programas del Intake activo en la hoja `ANUNCIOS` del Excel Maestro, junto con los leads de cada anuncio y su temperatura (caliente / tibio / frío). `GASTOS` no recibe valores de Meta: recibe **fórmulas que suman desde ANUNCIOS** por campaña y semana. Meta es la fuente única del gasto y se escribe una sola vez.

La pregunta de negocio que esto responde es: *"este anuncio está caro, pero ¿trae calientes o no?"*. La responde la columna `AG Veredicto` de ANUNCIOS, que el Excel calcula solo.

## Reglas
1. **Autonomía total**: no preguntes, no pidas confirmación. Si algo es imposible, detente con mensaje claro y exit ≠ 0.
2. **Meta Ads es SOLO LECTURA.** Nunca crear, editar, pausar ni activar nada en Meta.
3. **Una fila por anuncio por semana.** Llave de ANUNCIOS = `Semana (A)` + `ID anuncio (F)`. Llave de GASTOS = `Semana (A)` + `ID campaña (I)`.
4. **Idempotencia**: si la fila existe, actualízala; nunca dupliques. Cada corrida reescribe con los acumulados actuales de la semana en curso.
5. **Semanas a procesar**: la semana en curso (lunes → hoy) y la anterior completa (lunes → domingo). Si hoy es lunes, la semana anterior se reescribe con datos finales. No retrocedas más.
6. **IDs de Meta (18 dígitos) siempre como texto con apóstrofo** (`'1202…`). Sin el apóstrofo, Excel trunca los últimos 3 dígitos y la llave deja de cuadrar.
7. **Escritura secuencial o máximo 2 en paralelo** al mismo Excel. Con más, Graph devuelve 405 o se pierden datos. Escribe por bloque con `null` en las celdas que no se tocan (Graph las deja intactas).
8. **Nunca escribir en columnas AUTO** (ANUNCIOS S:AG; GASTOS P en adelante) ni pisar un valor que alguien escribió a mano.
9. **Una fila está libre solo si A:R (ANUNCIOS) o A:O (GASTOS) están vacías.** Las fórmulas AUTO no cuentan como ocupadas. Ubica filas por llave releyendo el rango antes de escribir; nunca uses un contador en memoria.
10. **Anti-corrimiento**: antes de actualizar una fila existente, relee su celda llave (ANUNCIOS F{n}, GASTOS I{n}) y confirma que coincide. Al terminar, relee el bloque de la semana y valida que no haya filas fantasma (datos sin llave).
11. **Mapear por encabezado, nunca por posición**, en PROGRAMAS, en los Sheets de leads y en los Excel de MKT CODE. El equipo inserta columnas seguido.
12. **Leads de prueba**: ignora filas cuyo correo sea `test@meta.com` o cuyas respuestas digan `<test lead`.
13. **Fechas**: el serial de Excel cuenta días desde 1899-12-30 (ej. 2026-10-12 = 46307). En MKT CODE, un serial con día ≤ 12 puede ser una fecha mm/dd leída como dd/mm. Para el cruce no importa, porque se cruza por correo, no por fecha.

## Asignación de herramientas
| Operación | Herramienta |
|---|---|
| Gasto y resultados por anuncio | Meta ADS MCP `ads_get_ad_entities` (antes, verificar campos con `ads_get_field_context`) |
| Leer y escribir el Excel Maestro | Composio, en `COMPOSIO_REMOTE_WORKBENCH` con `run_composio_tool`: `ONE_DRIVE_READ_WORKBOOK_RANGE` (`worksheet_id_or_name`, `address`, `select_fields:["values"]` o `["formulas"]`) y `ONE_DRIVE_UPDATE_WORKBOOK_RANGE` (`values` / `formulas`) |
| Listar Sheets de leads | Google Drive MCP `search_files` con `parentId = '<LEADS_DRIVE_FOLDER_ID>'` |
| Leer un Sheet de leads | Google Drive MCP `read_file_content`, o `download_file_content` con `exportMimeType=text/csv` para datos completos |
| Listar Excel de MKT CODE | Composio `ONE_DRIVE_LIST_FOLDER_CHILDREN` con `drive_id` explícito. No usar búsqueda: el índice va atrasado. |
| Leer un Excel de MKT CODE | Composio `EXCEL_GET_WORKSHEET_USED_RANGE` (clave `values`) |

IDs en `.env.example`.

---

## Estructura de las hojas

### PROGRAMAS (solo lectura)
Encabezados en la fila 5, datos desde la fila 6. Usar las columnas **Código** e **Intake**. El Intake activo es el valor más reciente de la columna Intake; hoy es "Intake 3". Si existe `CONFIG` "Intake activo", usar ese.

### ANUNCIOS (encabezados fila 5, datos 6–3005)
| Col | Campo | Quién |
|---|---|---|
| A | Semana (lunes) — serial de fecha | agente |
| B | Código | agente |
| C | Campaña (nombre) | agente |
| D | ID campaña — texto con `'` | agente |
| E | Conjunto (nombre) | agente |
| F | ID anuncio — texto con `'`; `ORGÁNICO` para leads orgánicos o sin anuncio (una fila por código y semana) | agente |
| G | Nombre anuncio | agente |
| H | Plataforma (`Facebook` / `Instagram` / `Mixto`, según el `platform` de los leads o el placement de Meta; si no se sabe, `Meta`) | agente |
| I | Ángulo / mensaje — sale del nombre (ver nomenclatura) | agente |
| J | Formato — sale del nombre | agente |
| K | Gasto (USD) | agente |
| L | Resultados Meta (leads del formulario o conversaciones de WhatsApp) | agente |
| M | Leads (Sheet) — leads del Sheet con ese `ad_id` en la semana | agente |
| N/O/P | Calientes / Tibios / Fríos | agente |
| Q | Sin clasificar (leads sin cruce en MKT CODE) | agente |
| R | Actualizado (fecha y hora de la corrida) | agente |
| S:Y | AUTO (producto, %, CPL, costo por caliente, Cuadre, ¿Lunes?) | **no tocar** |
| Z:AG | Veredicto acumulado | **no tocar** |

Debe cumplirse **N+O+P+Q = M**; la columna X (Cuadre) lo verifica.

### GASTOS (encabezados fila 5, datos desde 6)
A Semana · B Código · C Bolsa · D Plataforma · E Funnel · F Ángulo · G Formato · H Nombre campaña · I ID campaña · J Objetivo · K Estado · L Gasto · M Resultados · N Calificados · O Responsable · P+ AUTO.
- Las filas de Meta las crea el agente: A, B, C=`Pauta`, D=`Meta`, H, I (con `'`), J (`Formulario` o `WhatsApp`), K (`Activa`), y **fórmulas** en L, M y N:
  - `L{r}` = `=SUMIFS(ANUNCIOS!$K$6:$K$3005,ANUNCIOS!$A$6:$A$3005,$A{r},ANUNCIOS!$D$6:$D$3005,$I{r})`
  - `M{r}` = lo mismo con `ANUNCIOS!$L$6:$L$3005`
  - `N{r}` = `SUMIFS` de `ANUNCIOS!$N` + `SUMIFS` de `ANUNCIOS!$O` con los mismos criterios
- **F (Ángulo), G (Formato), E (Funnel) y O (Responsable) NO se llenan.** Son manuales, para gastos que no son de Meta.
- Las filas cuya columna D no es `Meta` son manuales: nunca tocarlas.

### Nomenclatura de anuncios
Nombre: `CÓDIGO_FUNNEL_ÁNGULO_FORMATO_vN`, por ejemplo `PE.EI.10-26.4_TOFU_AD_VID_v1`. Las tablas de códigos están en `CONFIG!AC5:AF11`: código de ángulo → ángulo (AC→AD) y código de formato → formato (AE→AF). Leerlas en cada corrida, no copiarlas aquí.
- Separar el nombre por `_`. Si tiene 5 partes y los códigos existen en CONFIG, escribir I = ángulo y J = formato.
- Si no cumple, escribir `Sin código` en I y J. **No adivinar** por palabras sueltas como "Video" o "Temario".
- Si una fila previa del mismo ID anuncio tiene I/J escritos a mano (distintos de `Sin código`), copiarlos.

---

## Pasos

### PASO 0 — Fechas
`hoy`, `lunes_actual`, `lunes_anterior = lunes_actual - 7`, `domingo_anterior = lunes_actual - 1`. Calcular el serial de Excel de cada lunes.

### PASO 1 — Códigos y estado del Excel
1. Leer PROGRAMAS `A5:AZ305` y ubicar las columnas por encabezado. `codigos` = códigos del Intake activo.
2. Leer ANUNCIOS `A6:R3005` y GASTOS `A6:O805` con `select_fields:["values"]`. Construir los índices llave → fila y la lista de filas libres.
3. Leer CONFIG `AC5:AF11` (nomenclatura).

### PASO 2 — Meta por anuncio (solo lectura), por cada semana
```
ad_account_id: 2954534534816734
level: "ad"
fields: ["id","name","campaign_id","campaign_name","adset_name","objective","amount_spent","lead","results"]
filtering: [{"field":"amount_spent","operator":"GREATER_THAN","value":["0"]}]
time_range: {"since":"<lunes>","until":"<domingo o hoy>"}
```
- Paginar con `pagination.next_cursor` hasta agotar.
- Quedarse con los anuncios cuyo `campaign_name` contenga un código de `codigos`. El código sale con la regex `(PE|CE|SM|CP)\.EI\.\d+-\d+\.\d+` o por coincidencia exacta con la lista.
- **Resultados Meta**: si `results.indicator` contiene `messaging_conversation_started`, son conversaciones (`results.values[0].value`); si contiene `leadgen`, son leads (`lead`). No clasificar por `objective`: hay campañas de WhatsApp con objetivo `OUTCOME_SALES`.
- Si `results` viene vacío en un anuncio de mensajería con gasto, repetir la consulta para ese anuncio con `breakdowns:["publisher_platform"]` y sumar.
- Objetivo para GASTOS: `WhatsApp` si es mensajería, si no `Formulario`.

### PASO 3 — Leads por anuncio (Google Drive)
1. `search_files` con `parentId = '1eNKQ0NQwQpLltPsNnda9_IiPhUd5YvPf'` y paginar. Por cada código, el Sheet cuyo título es el código. Si no existe, ese código no tiene leads: M=0, N..Q=0 y anotarlo en el resumen.
2. Leer el Sheet completo, preferentemente exportado a CSV. Mapear las columnas por encabezado: `id`, `created_time`, `ad_id`, `ad_name`, `campaign_id`, `is_organic`, `platform`, `email`.
3. Filtrar por `created_time` dentro de la semana (ISO con zona horaria -05:00) y descartar los leads de prueba.
4. `ad_id` viene como `ag:123…` o `123…`: quitar el prefijo antes de cruzar con el `id` de Meta. Si no hay `ad_id`, o `is_organic=true`, el lead va a la fila `ORGÁNICO` del código.

### PASO 4 — Temperatura (MKT CODE)
1. Listar la carpeta `01XEGP2KYCIGWDO6XTTVBZDOF2IIBDMQ5S` (drive de MKT CODE) y tomar el Excel cuyo nombre contiene el código. Si hay varios, el de `lastModifiedDateTime` más reciente.
2. Leer todas sus hojas (CALIENTE, TIBIO, FRIO). Columnas por encabezado: **correo/email** y **TIPO**. La temperatura es la de la **columna TIPO**, no la de la pestaña: hay filas en FRIO que dicen CALIENTE.
3. Normalizar el correo (`strip().lower()`) en los dos lados y cruzar. Sin cruce = `Sin clasificar`. Si un correo aparece dos veces en MKT CODE, usar la temperatura más caliente.
4. Si el código no tiene Excel en MKT CODE, todos sus leads van a `Sin clasificar`.

### PASO 5 — Escribir ANUNCIOS
Por cada anuncio de Meta con gasto en la semana, y por cada `ad_id` con leads aunque no haya tenido gasto esa semana:
- Armar la fila A:R, más la fila `ORGÁNICO` por código si hubo leads orgánicos.
- Si la llave existe, actualizar A:R de esa fila; si no, usar la primera fila libre. Releer la llave antes de escribir (regla 10).
- D y F como texto con `'`. R = fecha y hora actual.

### PASO 6 — Escribir GASTOS
Por cada `campaign_id` × semana presente en ANUNCIOS (excepto `ORGÁNICO`):
- Si existe la fila (A + I), solo verificar que L, M y N tengan la fórmula. Si alguien puso un valor fijo, reemplazarlo por la fórmula y avisar en el resumen.
- Si no existe, insertarla con las columnas indicadas arriba y las fórmulas en L, M y N.
- **Campaña renombrada**: la llave es el ID, así que el nombre (H) solo se actualiza.

### PASO 7 — Validación
1. Releer ANUNCIOS y GASTOS del bloque escrito.
   - ANUNCIOS: X (Cuadre) sin `Revisar`.
   - GASTOS: L de la semana = suma de K de ANUNCIOS de esa campaña (±0.01) y sin `#¡VALOR!` ni `#VALUE!`.
2. Sin filas fantasma.
3. Gasto total de ANUNCIOS por semana = gasto total de Meta para esos códigos (±0.05).

### PASO 8 — Resumen
Por semana y código:
- anuncios escritos;
- gasto;
- leads (Meta vs Sheet);
- calientes, tibios, fríos y sin clasificar.

Al final:
- lista de anuncios con veredicto **Caro y frío** y **Ganador** (leer `ANUNCIOS!AG`);
- anuncios con `Sin código`;
- códigos sin Sheet o sin MKT CODE;
- cualquier error.

Si no hubo gasto ni leads reales de esos códigos en la semana, no escribir nada y decirlo.
