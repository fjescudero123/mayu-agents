# Handoff para ejecutar el pipeline MAYU Agents con Claude

## Objetivo

Dejar operativo el pipeline de GitHub Actions para desplegar y probar el fiscalizador de reuniones MAYU sin depender de PowerShell manual para cada cambio.

Codex ya preparo:

- Repo local: `C:\Users\felix\OneDrive\Escritorio\MAYU\estrategia\agentes\mayu-agents-github`
- Repo GitHub: `fjescudero123/mayu-agents`
- Workflow: `.github/workflows/deploy-mayu-agents.yml`
- Loader Azure Automation: `runbooks/FiscalizadorReunionesMAYU.ps1`
- Runtime operativo: `runtime/FiscalizadorReunionesMAYU.core.ps1`
- Config no sensible: `config/fiscalizador_config.json`
- Setup unico: `SETUP_PIPELINE_ONCE.ps1`

## Estado actual

El repo fue subido a GitHub correctamente.

Falta terminar el setup unico:

1. Azure OIDC para GitHub Actions.
2. GitHub repository variables.
3. GitHub repository secrets.
4. Environment `mayu-production` con aprobacion.
5. Primera ejecucion del workflow `Deploy MAYU Agents`.

## Usuarios/cuentas

- GitHub: `fjescudero123`
- Azure tenant: `d63ffa68-a19e-496f-aa33-2e28b2034369`
- Azure subscription: `ea7a6973-b97f-4c64-a11e-f4afce935d06`
- Resource group: `rg-mayu-minutas`
- Automation account: `aa-mayu-agent-runtime`
- Sender mail: `notificaciones@imayu.cl`

## Comando principal para Claude

En PowerShell normal de Felix:

```powershell
cd "C:\Users\felix\OneDrive\Escritorio\MAYU\estrategia\agentes\mayu-agents-github"
powershell -ExecutionPolicy Bypass -File .\SETUP_PIPELINE_ONCE.ps1
```

Si pide login:

- Azure: cuenta admin MAYU usada para `MAYU-Produccion`.
- GitHub: `fjescudero123`.

Si pide OpenAI API key, usar la key creada para los reportes inteligentes.

## Error corregido por Codex

El setup fallo antes en:

```text
Failed to parse string as JSON
az ad app federated-credential create --parameters ...
```

Codex corrigio `SETUP_PIPELINE_ONCE.ps1` para pasar los federated credentials como archivo JSON `@file`, no como string inline.

Antes de ejecutar, conviene subir esta correccion:

```powershell
git add SETUP_PIPELINE_ONCE.ps1 CLAUDE_HANDOFF.md
git commit -m "Fix Azure OIDC setup and add Claude handoff"
git push
```

## Despues del setup

En GitHub:

1. Ir a `fjescudero123/mayu-agents`.
2. Abrir **Actions**.
3. Workflow: **Deploy MAYU Agents**.
4. Ejecutar `Run workflow`.
5. Primer test recomendado:
   - `run_mode`: `test`
   - `lookback_days`: `7`
6. Si queda esperando aprobacion de environment, aprobar `mayu-production`.

Validacion esperada:

- Llega correo de prueba a `fjescudero@imayu.cl`.
- El workflow termina verde.

## Pruebas siguientes

Ejecutar desde GitHub Actions:

1. `weekly_report` para probar reporte semanal.
2. `manual_due_sweep` con `lookback_days=7` para probar alertas pendientes.

## Regla de trabajo futura

- Codex mantiene arquitectura, codigo, runtime y bugs logicos del agente.
- Claude ejecuta pasos que requieren navegador, sesiones interactivas, aprobaciones o troubleshooting de portal.
- Si el workflow falla por error de codigo, copiar logs a Codex para parchear en `runtime/FiscalizadorReunionesMAYU.core.ps1` o scripts.
- Si falla por permisos/login/portal, Claude lo resuelve en navegador/Azure/GitHub.

## Archivos sensibles

No subir:

- `mayu_mail_config.json`
- claves OpenAI
- client secrets

Los secretos deben vivir en GitHub Secrets:

- `MAYU_GRAPH_CLIENT_SECRET`
- `MAYU_OPENAI_API_KEY`

Variables GitHub esperadas:

- `AZURE_CLIENT_ID`
- `AZURE_TENANT_ID`
- `AZURE_SUBSCRIPTION_ID`
- `MAYU_GRAPH_TENANT_ID`
- `MAYU_GRAPH_CLIENT_ID`
- `MAYU_OPENAI_MODEL`
