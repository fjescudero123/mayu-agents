# MAYU Agents Runtime

Repositorio operacional para agentes autonomos MAYU.

El fiscalizador de reuniones corre directamente en GitHub Actions. Ya no necesita Azure Automation, Azure login ni despliegue de runbook para su operacion normal.

## Que ejecuta

- Workflow directo: `.github/workflows/run-fiscalizador.yml`
- Runtime operativo: `runtime/FiscalizadorReunionesMAYU.core.ps1`
- Config no sensible: `config/fiscalizador_config.json`
- SharePoint y correos via Microsoft Graph
- Resumenes con OpenAI usando GitHub Secrets
- Seguimiento inteligente de compromisos, bloqueos, decisiones y continuidad entre reuniones

El workflow antiguo `.github/workflows/deploy-mayu-agents.yml` queda como referencia legacy de Azure Automation.

## Modelo operativo

1. Codex edita `runtime/FiscalizadorReunionesMAYU.core.ps1` o `config/fiscalizador_config.json`.
2. GitHub Actions ejecuta el runtime directo con PowerShell.
3. El workflow carga `config/fiscalizador_config.json` en la variable `MayuFiscalizadorConfigJson`.
4. Las credenciales se leen desde GitHub Secrets/Variables.
5. El agente lee SharePoint, envia correos por Graph y registra outputs en SharePoint igual que antes.
6. Cuando hay minuta, extrae compromisos y bloqueos, guarda JSON estructurado y los incorpora al reporte.

## Setup unico

1. Crear o usar el repositorio privado `fjescudero123/mayu-agents`.
2. Subir el contenido de esta carpeta como raiz del repo.
3. Configurar las variables y secrets indicados abajo en GitHub.
4. Confirmar que el workflow **Run Fiscalizador MAYU** aparece en Actions.

No se requiere OIDC, Azure login ni environment con aprobacion para la ejecucion directa del fiscalizador.

## Variables de GitHub requeridas

- `MAYU_GRAPH_TENANT_ID`
- `MAYU_GRAPH_CLIENT_ID`
- `MAYU_OPENAI_MODEL`

## Secrets de GitHub requeridos

- `MAYU_GRAPH_CLIENT_SECRET`
- `MAYU_OPENAI_API_KEY`

## Ejecucion manual en GitHub Actions

En GitHub Actions, usar workflow **Run Fiscalizador MAYU**:

- `post_reunion`: revisa si falta evidencia despues de una reunion especifica o de las reuniones vencidas del dia.
- `test`: envia correo de prueba.
- `weekly_report`: genera reporte semanal.
- `manual_due_sweep`: barre alertas pendientes.
- `monthly_report`: genera reporte mensual.

Parametro adicional:

- `lookback_days`: dias hacia atras para `manual_due_sweep`; default `7`.
- `tipo`: tipo de reunion para `post_reunion`; opcional.
- `date`: fecha `YYYY-MM-DD` para `post_reunion`; opcional.

Para I+D, el tipo interno sigue siendo `id`, pero el fiscalizador acepta archivos con nombre `AGENDA_I+D_DD_MM_YYYY.docx` y `AGENDA_id_DD_MM_YYYY.docx`.

## Seguimiento inteligente

El agente guarda seguimiento estructurado en:

- `minutas_archivadas/seguimiento_compromisos/`

Por cada reunion con evidencia, genera un JSON con:

- resumen,
- asistencia detectada,
- compromisos,
- compromisos cumplidos desde la reunion anterior,
- pendientes arrastrados,
- bloqueos,
- decisiones,
- riesgos,
- seguimiento recomendado.

El reporte semanal y mensual de Felix incluye esta lectura. En el reporte semanal, los responsables tambien reciben por correo su bloque de compromisos y bloqueos.

## Calificacion mensual 0-100

El cierre mensual genera una ficha por responsable y la envia por correo, con Felix en copia.

La nota usa estos criterios:

- 25% asistencia / evidencia: reuniones esperadas con documento o minuta.
- 35% cumplimiento de tareas: compromisos cumplidos, parciales, pendientes, vencidos o bloqueados.
- 20% resolucion de problemas: bloqueos y pendientes arrastrados.
- 20% avances limpios: bloqueos, riesgos, alertas y arrastres.

Bandas:

- `90-100`: excelente.
- `75-89`: bien.
- `60-74`: requiere seguimiento.
- `<60`: critico.

Las fichas quedan guardadas en SharePoint:

- `minutas_archivadas/reportes/scorecards_YYYY-MM/`

Secuencia sugerida para validar:

1. Ejecutar `test` y confirmar que llega correo a `fjescudero@imayu.cl`.
2. Ejecutar `weekly_report` y confirmar correo + copia en SharePoint.
3. Ejecutar `manual_due_sweep` con `lookback_days=7`; puede enviar `0` alertas si ya estaban registradas.
4. Ejecutar `monthly_report` solo si se quiere probar manualmente fuera del schedule.

## Schedules activos

GitHub usa cron en UTC:

- `0 12 * * 1`: reporte semanal, lunes AM Chile.
- `40 14 * * 1-5`: post reunion planta diario.
- `45 15 * * 1`: post reunion comercial.
- `0 17 * * 2`: post reunion proyectos.
- `30 21 * * 3`: post reunion I+D.
- `30 15 * * 4`: post reunion financiero semanal.
- `0 18 * * 5`: post reunion ejecutivo.
- `0 15 * * 1-5`: barrido de alertas vencidas, lunes a viernes.
- `30 21 * * 1-5`: segundo barrido de alertas vencidas, lunes a viernes.
- `30 22 28-31 * *`: reporte mensual; el runtime solo envia si es ultimo dia del mes.

## Seguridad

No subir `mayu_mail_config.json`, claves OpenAI reales, client secrets reales ni archivos sensibles. Las credenciales viven como GitHub Secrets.
