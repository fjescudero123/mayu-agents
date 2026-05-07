# MAYU Agents Runtime

Repositorio operacional para agentes autonomos MAYU.

El fiscalizador de reuniones corre directamente en GitHub Actions. Ya no necesita Azure Automation, Azure login ni despliegue de runbook para su operacion normal.

## Que ejecuta

- Workflow directo: `.github/workflows/run-fiscalizador.yml`
- Runtime operativo: `runtime/FiscalizadorReunionesMAYU.core.ps1`
- Config no sensible: `config/fiscalizador_config.json`
- SharePoint y correos via Microsoft Graph
- Resumenes con OpenAI usando GitHub Secrets

El workflow antiguo `.github/workflows/deploy-mayu-agents.yml` queda como referencia legacy de Azure Automation.

## Modelo operativo

1. Codex edita `runtime/FiscalizadorReunionesMAYU.core.ps1` o `config/fiscalizador_config.json`.
2. GitHub Actions ejecuta el runtime directo con PowerShell.
3. El workflow carga `config/fiscalizador_config.json` en la variable `MayuFiscalizadorConfigJson`.
4. Las credenciales se leen desde GitHub Secrets/Variables.
5. El agente lee SharePoint, envia correos por Graph y registra outputs en SharePoint igual que antes.

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

- `test`: envia correo de prueba.
- `weekly_report`: genera reporte semanal.
- `manual_due_sweep`: barre alertas pendientes.
- `monthly_report`: genera reporte mensual.

Parametro adicional:

- `lookback_days`: dias hacia atras para `manual_due_sweep`; default `7`.

Secuencia sugerida para validar:

1. Ejecutar `test` y confirmar que llega correo a `fjescudero@imayu.cl`.
2. Ejecutar `weekly_report` y confirmar correo + copia en SharePoint.
3. Ejecutar `manual_due_sweep` con `lookback_days=7`; puede enviar `0` alertas si ya estaban registradas.
4. Ejecutar `monthly_report` solo si se quiere probar manualmente fuera del schedule.

## Schedules activos

GitHub usa cron en UTC:

- `0 12 * * 1`: reporte semanal, lunes AM Chile.
- `0 15 * * 1-5`: barrido de alertas vencidas, lunes a viernes.
- `30 21 * * 1-5`: segundo barrido de alertas vencidas, lunes a viernes.
- `30 22 28-31 * *`: reporte mensual; el runtime solo envia si es ultimo dia del mes.

## Seguridad

No subir `mayu_mail_config.json`, claves OpenAI reales, client secrets reales ni archivos sensibles. Las credenciales viven como GitHub Secrets.
