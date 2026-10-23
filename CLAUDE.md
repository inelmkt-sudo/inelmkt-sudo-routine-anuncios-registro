# Routine: Registro semanal por anuncio → ANUNCIOS + GASTOS (Excel Maestro MKT)

## Objetivo
Registrar el corte semanal de **LEADS** (arribos, calientes, tibios y atención ≤24h, una fila por código y semana) y, una fila por **anuncio × semana**, el gasto y los resultados de Meta Ads de los programas del Intake activo en la hoja `ANUNCIOS` del Excel Maestro, junto con los leads de cada anuncio y su temperatura (caliente / tibio / frío). `GASTOS` no recibe valores de Meta: recibe **fórmulas que suman desde ANUNCIOS** por campaña y semana. Meta es la fuente única del gasto y se escribe una sola vez.

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
11. **Mapear por encabezado, nunca por posición**, en PROGRAMAS, en los Sheets de leads y en los consolidados de AMG. El equipo inserta columnas seguido.
12. **Leads de prueba**: ignora filas cuyo correo sea `test@meta.com` o cuyas respuestas digan `<test lead`.
13. **Fechas**: el serial de Excel cuenta días desde 1899-12-30 (ej. 2026-10-12 = 46307). En los consolidados, un serial con día ≤ 12 puede ser una fecha mm/dd leída como dd/mm. Para el cruce no importa, porque se cruza por correo, no por fecha.

## Asignación de herramientas
| Operación | Herramienta |
|---|---|
| Gasto y resultados por anuncio | Meta ADS MCP `ads_get_ad_entities` (antes, verificar campos con `ads_get_field_context`) |
| Leer y escribir el Excel Maestro | Composio, en `COMPOSIO_REMOTE_WORKBENCH` con `run_composio_tool`: `ONE_DRIVE_READ_WORKBOOK_RANGE` (`worksheet_id_or_name`, `address`, `select_fields:["values"]` o `["formulas"]`) y `ONE_DRIVE_UPDATE_WORKBOOK_RANGE` (`values` / `formulas`) |
| Listar Sheets de leads | Google Drive MCP `search_files` con `parentId = '<LEADS_DRIVE_FOLDER_ID>'` |
| Leer un Sheet de leads | Google Drive MCP `read_file_content`, o `download_file_content` con `exportMimeType=text/csv` para datos completos |
| Listar consolidados de leads (drive TI) | Composio `ONE_DRIVE_LIST_FOLDER_CHILDREN` con `drive_id` explícito. No usar búsqueda: el índice va atrasado. |
| Escribir `lead_status` en un Sheet | Composio Google Sheets `GOOGLESHEETS_BATCH_UPDATE` (solo esa columna) |
| Leer un consolidado | Composio `EXCEL_GET_WORKSHEET_USED_RANGE` (clave `values`) |

IDs en `.env.example`.

---

## Estructura de las hojas

### PROGRAMAS (solo lectura)
Encabezados en la fila 5, datos desde la fila 6. Usar las columnas **Código** e **Intake**. **Por ahora se procesa solo "Intake 3"** (decisión del usuario). Para sumar otro intake hay que ampliar la ruta del PASO 4.

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
| Q | Sin clasificar (leads sin cruce en el consolidado) | agente |
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

### PASO 4 — Temperatura (Consolidado de leads AMG)
La carpeta MKT CODE ya no se usa. La clasificación vive en el drive de TI (`TI_DRIVE_ID`), en un Excel por programa:
`/INEL/Institute/Open Live/2026/Intake 3/<CÓDIGO>_<Nombre>/02_AMG/Consolidado Leads/<archivo>.xlsx`
1. Listar `/INEL/Institute/Open Live/2026/Intake 3` con Composio `ONE_DRIVE_LIST_FOLDER_CHILDREN` (`drive_id` = TI, `folder_path`). Tomar la carpeta cuyo nombre empieza con el código, entrar a `02_AMG` y luego a `Consolidado Leads` (con `folder_item_id`), y tomar el `.xlsx` más reciente (nunca un `.xlsx.url`).
2. **Ignorar por completo los archivos `.xlsx.url`**: son accesos directos que se copiaron del original. No abrirlos, no seguirlos y no editarlos. Usar solo el `.xlsx` real de la carpeta. Si no hay ninguno, ese código va a `Sin clasificar` y se anota en el resumen.
3. Leer las hojas CALIENTE, TIBIO y FRIO con `EXCEL_GET_WORKSHEET_USED_RANGE` (encabezado en la fila 1). Columnas por nombre: **`Correo`** y **`TIPO`**. La temperatura sale de la **columna TIPO**, no de la pestaña.
4. Normalizar el correo (`strip().lower()`) y cruzarlo con los leads del Sheet. Sin cruce = `Sin clasificar`. Si un correo aparece repetido, usar la temperatura más caliente.

### PASO 4-bis — Escribir la temperatura en el Sheet (columna `lead_status`)
Una vez cruzados los correos, cada lead del Sheet recibe su temperatura en la columna **`lead_status`**, que se ubica por encabezado:
- Valores: `CALIENTE`, `TIBIO` o `FRIO`, según la columna TIPO del consolidado de AMG. Si el lead no cruzó, **no tocar** su `lead_status`: queda como lo dejó Meta (`CREATED`, `QUALIFIED`…).
- Herramienta: Composio Google Sheets (`GOOGLESHEETS_BATCH_UPDATE` o `GOOGLESHEETS_VALUES_UPDATE`), con el `spreadsheet_id` del Sheet y la hoja `Hoja 1`. Escribir **solo la columna `lead_status`**, celda por celda o en bloque, con `null` o el valor actual en las filas que no cambian. Nunca modificar ninguna otra columna ni borrar filas: el Sheet lo alimenta Meta.
- Aplicar a todos los leads del Sheet (no solo los de la semana), para mantener el estado al día. Respetar el límite de Google Sheets: 60 escrituras por minuto.
- Para contar en ANUNCIOS (N/O/P/Q), usar la temperatura resultante de este paso. Es la misma fuente, así que el Sheet y ANUNCIOS siempre dicen lo mismo.
- Ignorar los leads de prueba (`test@meta.com` o `<test lead`): no escribirles estado.

### PASO 4-ter — LEADS (corte semanal por programa)
La hoja `LEADS` del Excel Maestro (encabezados fila 5, datos desde la 6) se llena **desde el consolidado de AMG** (hojas CALIENTE / TIBIO / FRIO de `02_AMG/Consolidado Leads`, el mismo del PASO 4) y la atención comercial desde el AGC. **No usa el Sheet de Meta de Drive.** Una fila por código y semana; llave = `Fecha corte (A)` + `Código (B)`. `A` = el **lunes** de la semana (serial de Excel). Si la fila existe, se reescribe con los acumulados actuales; si no, se usa la primera fila con A:G vacías. **Nunca tocar J:O** (son AUTO).

| Col | Campo | De dónde sale |
|---|---|---|
| A | Fecha corte | lunes de la semana |
| B | Código | código del programa |
| C | Leads pagados | total de leads del consolidado (CALIENTE + TIBIO + FRIO, según la columna `TIPO`) cuya `Fecha` cae en la semana. Un correo repetido cuenta una vez |
| D | Leads orgánicos | **no se escribe**: el consolidado no distingue pagado de orgánico. Queda como lo ponga el equipo |
| E | Calientes (pagados) | leads del consolidado de la semana con `TIPO` = CALIENTE |
| F | Tibios (pagados) | leads del consolidado de la semana con `TIPO` = TIBIO |
| G | Atendidos ≤24h | AGC (ver abajo) |
| H | Ventas (inscritos) | **no se escribe** (pendiente de definir). Dejarla como está |
| I | Nota | solo para avisos (p. ej. `AGC no disponible`); **no pisar una nota escrita a mano** |

**Fecha del consolidado:** la columna `Fecha` trae formatos mezclados. Un serial de Excel con día ≤12 es una fecha mm/dd que Excel leyó como dd/mm: intercambiar día y mes. El texto con segundos y AM/PM (formato Pabbly) es dd/mm literal. Si una fecha no se puede interpretar, no contar ese lead y anotarlo en el resumen.

**Control:** `E + F` no puede superar `C`. Si el consolidado del código no existe o no tiene leads en la semana, no escribir la fila y decirlo en el resumen.

**Atendidos ≤24h (G)** — archivo de atención del programa en `<programa>/03_AGC/` (drive de TI):
1. Usar solo el `.xlsx` real. **Ignorar los `.xlsx.url`** (como en el PASO 4). Hoy no tienen archivo AGC: CE.EI.06-26.2 y PE.EI.15-26.2. En ese caso **dejar G vacío** (no escribir 0) y poner `AGC no disponible` en la Nota si está vacía.
2. Hojas `CALIENTE`, `TIBIO`, `FRIO` y `MANUAL`. **Encabezados en la fila 6, datos desde la 7** (las filas 1-5 son títulos y SLAs). Mapear por nombre: `Fecha Ingreso Lead`, `Hora`, `Estado del lead`, y `Fecha CALL 1..4`, `Hora CALL 1..4`, `Fecha WTS 1..4`, `Hora WTS 1..4`.
3. **Leer en bloques** (≤300 filas) y solo las columnas necesarias: el archivo pesa ~2 MB y leerlo entero da error 413.
4. Para cada lead cuya `Fecha Ingreso Lead` cae en la semana (en cualquiera de las 4 hojas): **primer contacto** = la menor fecha+hora entre CALL 1-4 y WTS 1-4 con fecha no vacía. Atendido si el primer contacto ocurrió ≤24 h después del ingreso. Si al contacto le falta la hora, comparar solo fechas (≤1 día). Si el contacto figura antes del ingreso por falta de hora, cuenta como atendido. Sin ninguna gestión = no atendido.
5. Los leads con `Estado del lead` = `Lead Invalido` cuentan como atendidos: son bajas deliberadas, no desatención. LEADS calcula `Sin atender en SLA = Total arribos - Atendidos`.
6. `Fecha Ingreso Lead` puede venir como texto `dd/mm/aaaa` (formato Pabbly) o como serial. Interpretar el texto como día/mes. Los seriales leerlos literales.
7. `G` es el total de atendidos de la semana, sumando las 4 hojas.

**Ventas (H):** no se automatiza por ahora (se define aparte). La rutina no escribe esa columna.

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
Por código, LEADS: leads, calientes, tibios, atendidos ≤24h y sin atender en SLA (leer LEADS `O`); marcar los códigos sin AGC y las diferencias de coherencia con ANUNCIOS.
Incluir por código: leads del Sheet a los que se les actualizó `lead_status` (cuántos CALIENTE, TIBIO y FRIO).

Por semana y código:
- anuncios escritos;
- gasto;
- leads (Meta vs Sheet);
- calientes, tibios, fríos y sin clasificar.

Al final:
- lista de anuncios con veredicto **Caro y frío** y **Ganador** (leer `ANUNCIOS!AG`);
- anuncios con `Sin código`;
- códigos sin Sheet o sin consolidado;
- cualquier error.

Si no hubo gasto ni leads reales de esos códigos en la semana, no escribir nada y decirlo.
