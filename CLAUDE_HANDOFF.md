# Handoff MAYU Agents

Actualizado: 2026-05-07

## Estado actual

Migracion del fiscalizador de reuniones MAYU a **GitHub Actions como runtime directo** publicada en `main`.

PR #1 mergeado a `main`.

- Commit main: `269f28d`
- Workflow directo: `.github/workflows/run-fiscalizador.yml`
- Runtime: `runtime/FiscalizadorReunionesMAYU.core.ps1`
- Sin Azure Automation para operacion normal.
- Sin Azure login.
- Sin environment con aprobacion.

La logica del agente ya esta validada. No hay que redisenarla.

Validaciones directas ya ejecutadas:

- GitHub Actions run: `25502057346`
- Modo: `test`
- Duracion: 15s
- Resultado: OK
- Output clave:
  - `Modo resuelto: test / LookbackDays=7`
  - `Fiscalizador MAYU iniciado. Modo=test Tipo= Fecha= HoraMAYU=2026-05-07T10:27:53`
  - `Prueba enviada a fjescudero@imayu.cl.`

- GitHub Actions run: `25503779552`
- Modo: `weekly_report`
- Duracion: 22s
- Resultado: OK
- Periodo: `2026-04-27` a `2026-05-03`
- Filas fiscalizadas: 8
- Resultado clave: lectura OpenAI generada, correo enviado a `fjescudero@imayu.cl`, copia guardada en SharePoint.

- GitHub Actions run: `25503848173`
- Modo: `manual_due_sweep`
- Duracion: 14s
- Resultado: OK
- Periodo: `2026-04-30` a `2026-05-07`
- Alertas enviadas: 0
- Lectura: comportamiento esperado; las alertas anteriores ya estaban registradas en `minutas_archivadas/alertas_enviadas/`.

Conclusion: GitHub Actions directo queda validado para `test`, `weekly_report` y `manual_due_sweep`.

Nota I+D / `id`:

- El archivo real de I+D del `2026-05-06` fue encontrado en SharePoint como `AGENDA_I+D_06_05_2026.docx`.
- Codex lo renombro en SharePoint a `AGENDA_id_06_05_2026.docx` para que el fiscalizador publicado lo reconozca de inmediato.
- Cambio local preparado en `runtime/FiscalizadorReunionesMAYU.core.ps1`: mantener `id` como tipo interno, pero aceptar evidencia con nombres `AGENDA_I+D_DD_MM_YYYY.docx` y `AGENDA_id_DD_MM_YYYY.docx`.
- Hacia adelante, Felix prefiere que el nombre humano del archivo sea `I+D`, no `id`.

Nueva caracteristica en desarrollo: Agente de Continuidad Operacional.

Felix aprobo implementar en 2 fases:

1. Seguimiento inteligente de compromisos y bloqueos semana contra semana.
2. Calificacion mensual 0-100 por responsable.

Fase 1 preparada localmente por Codex:

- `config/fiscalizador_config.json`: agrega `sharepoint.seguimiento_folder`.
- `runtime/FiscalizadorReunionesMAYU.core.ps1`:
  - lee texto desde DOCX/HTML/MD,
  - extrae compromisos, bloqueos, decisiones, riesgos y asistencia detectada con OpenAI,
  - compara con seguimiento anterior del mismo tipo de reunion si existe,
  - guarda JSON en `minutas_archivadas/seguimiento_compromisos/`,
  - incorpora compromisos/bloqueos al reporte semanal y mensual de Felix,
  - envia seguimiento semanal por correo a los responsables de cada reunion, con Felix en copia.
- `README.md`: documenta seguimiento inteligente.

Fase 2 preparada localmente:

- score mensual 0-100 por responsable.
- criterios:
  - 25% asistencia / evidencia,
  - 35% cumplimiento de tareas,
  - 20% resolucion de problemas,
  - 20% avances limpios.
- bandas:
  - `90-100`: excelente,
  - `75-89`: bien,
  - `60-74`: requiere seguimiento,
  - `<60`: critico.
- En `monthly_report`, el agente:
  - genera scorecards por responsable,
  - envia correo individual a cada responsable con Felix en copia,
  - guarda copia HTML en `minutas_archivadas/reportes/scorecards_YYYY-MM/`.

Nueva decision despues de validar `weekly_report` y `manual_due_sweep`:

Felix quiere que los barridos **post-reunion** sigan existiendo en el entorno nuevo, no solo el barrido general. Por lo tanto, no pausar Azure Automation hasta que GitHub Actions tenga publicados y validados los schedules `post_reunion`.

Cambio local preparado por Codex:

- Archivo: `.github/workflows/run-fiscalizador.yml`
- Agrega `run_mode: post_reunion` en `workflow_dispatch`.
- Agrega inputs manuales opcionales:
  - `tipo`
  - `date`
- Pasa `-Tipo` y `-Date` al runtime.
- Agrega schedules post-reunion en GitHub Actions.

Schedules post-reunion propuestos en GitHub Actions:

| Cron UTC | Equivalente Chile aprox. | Proposito |
| --- | --- | --- |
| `40 14 * * 1-5` | 10:40 lun-vie | Planta diario |
| `45 15 * * 1` | 11:45 lunes | Comercial |
| `0 17 * * 2` | 13:00 martes | Proyectos |
| `30 21 * * 3` | 17:30 miercoles | I+D |
| `30 15 * * 4` | 11:30 jueves | Financiero semanal |
| `0 18 * * 5` | 14:00 viernes | Ejecutivo |

Pendiente para cerrar migracion operacional:

1. Publicar el cambio de `.github/workflows/run-fiscalizador.yml` en GitHub.
2. Validar manualmente `post_reunion`, idealmente con `tipo=planta_diario` y fecha de hoy o ultima reunion esperada.
3. Cuando el workflow con post-reunion este en `main`, pausar los schedules legacy de Azure Automation que siguen enabled.
4. Mantener Automation Account `aa-mayu-agent-runtime` como respaldo frio por 1-2 semanas, salvo instruccion distinta de Felix.
5. Opcional posterior, solo con confirmacion de Felix: limpieza definitiva de Automation Account, runbook, schedules, jobSchedules, OIDC app/env legacy.

## Repo y carpeta

- Repo GitHub: `fjescudero123/mayu-agents`
- Carpeta local: `C:\Users\felix\OneDrive\Escritorio\MAYU\estrategia\agentes\mayu-agents-github`

Archivos clave:

- Workflow actual Azure Automation: `.github/workflows/deploy-mayu-agents.yml`
- Loader Azure Automation: `runbooks/FiscalizadorReunionesMAYU.ps1`
- Runtime operativo validado: `runtime/FiscalizadorReunionesMAYU.core.ps1`
- Config no sensible: `config/fiscalizador_config.json`
- Setup Azure OIDC: `SETUP_PIPELINE_ONCE.ps1`

## Estado validado

GitHub Actions + Azure Automation quedo operativo.

Setup completado:

- Azure OIDC app `mayu-agents-github-deploy` creada.
- Federated credentials para `main` y environment `mayu-production`.
- 6 GitHub variables + 2 GitHub secrets configurados.
- Environment `mayu-production` con Felix como reviewer.
- Runbook + runtime publicados en Azure Automation:
  - Automation account: `aa-mayu-agent-runtime`
  - Runbook: `FiscalizadorReunionesMAYU`

Pruebas validadas:

| Modo | Resultado |
| --- | --- |
| `test` | OK, correo enviado a `fjescudero@imayu.cl` |
| `weekly_report` | OK, 8 filas fiscalizadas, lectura OpenAI, correo enviado, copia en SharePoint |
| `manual_due_sweep` | OK, 10 alertas enviadas y registradas en SharePoint |

Run `manual_due_sweep` validado:

- GitHub Actions run: `25475272465`
- Azure Automation job: `5ced4704-7781-4489-8ac9-145aba5919d8`
- Periodo: `2026-04-29` a `2026-05-06`
- Resultado: `Barrido manual completo. Alertas enviadas: 10`

Alertas enviadas:

1. `planta_diario` 2026-04-29 -> `fjerez`, cc `fjescudero`, `clecaros`
2. `id` 2026-04-29 -> `m.epelman`, cc `fjescudero`
3. `planta_diario` 2026-04-30 -> `fjerez`, cc `fjescudero`, `clecaros`
4. `financiero_semanal` 2026-04-30 -> `vescudero`, cc `fjescudero`
5. `planta_diario` 2026-05-04 -> `fjerez`, cc `fjescudero`, `clecaros`
6. `comercial` 2026-05-04 -> `fescudero`, `efernandez`, cc `fjescudero`
7. `planta_diario` 2026-05-05 -> `fjerez`, cc `fjescudero`, `clecaros`
8. `eerr_mensual` 2026-05-05 -> `vescudero`, cc `fjescudero`
9. `planta_diario` 2026-05-06 -> `fjerez`, cc `fjescudero`, `clecaros`
10. `id` 2026-05-06 -> `m.epelman`, cc `fjescudero`

SharePoint:

- Los 10 JSON quedaron en `minutas_archivadas/alertas_enviadas/`.
- Feriado `2026-05-01` respetado.
- Weekend `2026-05-02/03` respetado.

Observacion menor:

- `proyectos` del martes `2026-05-05` no genero alerta. Puede haber evidencia subida o un caso de calendario a revisar despues. No bloquea la migracion.

## Fixes ya aplicados

Commit Claude `dc91379`:

- `SETUP_PIPELINE_ONCE.ps1`: `gh secret set` ajustado porque `gh 2.89.0` no soportaba `--body-file`.
- `runbooks/FiscalizadorReunionesMAYU.ps1`: fix de encoding al cargar core script desde SharePoint. `Invoke-WebRequest` podia devolver `byte[]`; ahora usa `UTF8.GetString()` defensivo.

Codex tambien corrigio:

- JSON de federated credentials en `SETUP_PIPELINE_ONCE.ps1`, pasando `--parameters @file` en vez de JSON inline.

## Decision implementada

Azure quiere cobrar. Como el agente no necesita infraestructura Azure para correr, migrar a GitHub Actions directo.

Modelo nuevo:

- GitHub Actions ejecuta `runtime/FiscalizadorReunionesMAYU.core.ps1` directamente.
- No usar Azure login.
- No usar Azure Automation.
- Mantener GitHub Secrets:
  - `MAYU_GRAPH_CLIENT_SECRET`
  - `MAYU_OPENAI_API_KEY`
- Mantener GitHub Variables:
  - `MAYU_GRAPH_TENANT_ID`
  - `MAYU_GRAPH_CLIENT_ID`
  - `MAYU_OPENAI_MODEL`
- El script lee SharePoint y envia mail por Graph igual que hoy.

## Workflow directo

Publicado en `main` con `test`, `weekly_report`, `manual_due_sweep` y `monthly_report`.

Cambio local pendiente de publicar agrega `post_reunion` manual y schedules post-reunion:

`.github/workflows/run-fiscalizador.yml`

Soporta:

- `workflow_dispatch`
  - `run_mode`: `post_reunion`, `test`, `weekly_report`, `manual_due_sweep`, `monthly_report`
  - `lookback_days`: default `7`
  - `tipo`: opcional para `post_reunion`
  - `date`: opcional para `post_reunion`
- `schedule`
  - `0 12 * * 1`: `weekly_report`, lunes AM Chile.
  - `40 14 * * 1-5`: `post_reunion`, planta diario.
  - `45 15 * * 1`: `post_reunion`, comercial.
  - `0 17 * * 2`: `post_reunion`, proyectos.
  - `30 21 * * 3`: `post_reunion`, I+D.
  - `30 15 * * 4`: `post_reunion`, financiero semanal.
  - `0 18 * * 5`: `post_reunion`, ejecutivo.
  - `0 15 * * 1-5`: `manual_due_sweep`, lunes a viernes.
  - `30 21 * * 1-5`: segundo `manual_due_sweep`, lunes a viernes.
  - `30 22 28-31 * *`: `monthly_report`; el runtime solo envia si es ultimo dia del mes.

El workflow:

1. Hacer checkout.
2. Leer `config/fiscalizador_config.json`.
3. Setear variables de entorno:
   - `MayuFiscalizadorConfigJson`
   - `MayuTenantId`
   - `MayuClientId`
   - `MayuClientSecret`
   - `MayuOpenAiApiKey`
   - `MayuOpenAiModel`
4. Ejecutar:

```powershell
pwsh ./runtime/FiscalizadorReunionesMAYU.core.ps1 -Mode <run_mode> -Tipo <tipo> -Date <date> -LookbackDays <lookback_days>
```

## Validacion GitHub Actions directo

Validado correctamente:

1. `test`
2. `weekly_report`
3. `manual_due_sweep` con `lookback_days=7`

Pendiente solo si se quiere probar manualmente: `monthly_report`. El schedule mensual ya esta publicado; el runtime solo envia si es ultimo dia del mes.

## Azure Automation legacy

Schedules encontrados en `aa-mayu-agent-runtime` / `rg-mayu-minutas` / subscription `ea7a6973-b97f-4c64-a11e-f4afce935d06`:

| Schedule | Runbook ligado | Proxima corrida Chile | Estado |
| --- | --- | --- | --- |
| `mayuPostPlanta1040` | `FiscalizadorReunionesMAYU` | vie 2026-05-08 10:40 | enabled, semanal |
| `mayuPostComercial1145` | `FiscalizadorReunionesMAYU` | lun 2026-05-11 11:45 | enabled, semanal |
| `mayuPostProyectos1300` | `FiscalizadorReunionesMAYU` | mar 2026-05-12 13:00 | enabled, semanal |
| `mayuPostID1730` | sin link | mie 2026-05-13 17:30 | enabled, semanal, huerfano |

No existe schedule legado para `weekly_report`, `manual_due_sweep` global ni `monthly_report`.

Recomendacion: pausar los 4 schedules con PATCH ARM `isEnabled:false`. Es reversible y evita ejecuciones pagas de Azure Automation. No eliminar la cuenta todavia.

## Despues de validar GitHub Actions directo

1. Pausar/eliminar schedules de Azure Automation si existen activos.
2. Opcional: eliminar Automation Account `aa-mayu-agent-runtime` si Felix confirma que ya no se usara.
3. Mantener repo `mayu-agents` como fuente operacional.

## Prompt corto para Claude

Usar si hace falta ejecutar la parte externa con GitHub/portal:

```text
Continua MAYU Agents en repo `fjescudero123/mayu-agents`.

PR #1 ya fue mergeado a main con commit `269f28d`.

GitHub Actions directo ya quedo validado:
- `test`: run `25502057346`, success.
- `weekly_report`: run `25503779552`, success, correo y copia SharePoint.
- `manual_due_sweep`: run `25503848173`, success, 0 alertas.

Felix quiere mantener tambien los barridos post-reunion en el entorno nuevo antes de pausar Azure Automation.

Implementa en `.github/workflows/run-fiscalizador.yml`:
1. Agregar `post_reunion` como opcion de `workflow_dispatch.run_mode`.
2. Agregar inputs opcionales `tipo` y `date`.
3. Pasar `-Tipo` y `-Date` al runtime.
4. Agregar schedules post-reunion:
   - `40 14 * * 1-5` -> planta diario
   - `45 15 * * 1` -> comercial
   - `0 17 * * 2` -> proyectos
   - `30 21 * * 3` -> I+D
   - `30 15 * * 4` -> financiero semanal
   - `0 18 * * 5` -> ejecutivo

Hay un cambio local preparado por Codex con esa logica. Si puedes acceder al workspace local, revisa ese diff y publicalo por PR. Si no, replica el cambio en GitHub.

Despues de mergear:
1. Ejecuta manualmente workflow `Run Fiscalizador MAYU`, modo `post_reunion`, `tipo=planta_diario`, `date=<fecha de hoy o ultima fecha habil>`.
2. Confirma que termina verde y no reenvia alertas duplicadas si ya existen JSON en SharePoint.
3. Solo despues de validar `post_reunion`, pausar de forma reversible los 4 schedules legacy de Azure Automation:
   - `mayuPostPlanta1040`
   - `mayuPostComercial1145`
   - `mayuPostProyectos1300`
   - `mayuPostID1730`

No eliminar nada todavia:
- no eliminar schedules,
- no eliminar jobSchedules,
- no eliminar runbook,
- no eliminar Automation Account.

No tocar secretos ni subir archivos sensibles.
```

## Regla de trabajo

- Codex: arquitectura, codigo del agente, bugs logicos.
- Claude: ejecucion con navegador/sesiones/aprobaciones si Codex queda bloqueado por entorno.
- Si falla por codigo, llevar logs a Codex.
- Si falla por portal/login/permisos, resolver con Claude.

## No subir

- `mayu_mail_config.json`
- claves OpenAI reales
- client secrets reales
