# MAYU Agents Runtime

Repositorio de despliegue para agentes autonomos MAYU.

Este repo esta pensado para que Codex pueda modificar la logica de los agentes, abrir cambios y dejar que GitHub Actions publique en Azure con aprobacion explicita.

## Que despliega

- Azure Automation Account: `aa-mayu-agent-runtime`
- Runbook estable: `FiscalizadorReunionesMAYU`
- Runtime operativo en SharePoint: `minutas_archivadas/runtime/FiscalizadorReunionesMAYU.core.ps1`
- Variables seguras de Automation para Graph y OpenAI

## Modelo operativo

1. Codex edita `runtime/FiscalizadorReunionesMAYU.core.ps1` o `config/fiscalizador_config.json`.
2. GitHub Actions valida sintaxis.
3. GitHub Actions sube el runtime a SharePoint.
4. GitHub Actions publica el loader en Azure Automation.
5. Con `workflow_dispatch`, GitHub Actions puede ejecutar una prueba real.

El runbook de Azure queda estable; las correcciones futuras van al runtime.

## Setup unico

1. Crear un repositorio privado en GitHub, sugerido: `mayu-agents`.
2. Subir el contenido de esta carpeta como raiz del repo.
3. En GitHub, crear el environment `mayu-production` y activar aprobacion requerida por Felix.
4. Ejecutar una vez el bootstrap OIDC:

```powershell
pwsh ./scripts/Bootstrap-AzureOidc.ps1 -GitHubOwner fjescudero123 -GitHubRepo mayu-agents
```

5. Crear las variables y secretos que imprime el bootstrap.

## Variables de GitHub

- `AZURE_CLIENT_ID`
- `AZURE_TENANT_ID`
- `AZURE_SUBSCRIPTION_ID`
- `MAYU_GRAPH_TENANT_ID`
- `MAYU_GRAPH_CLIENT_ID`
- `MAYU_OPENAI_MODEL`

## Secrets de GitHub

- `MAYU_GRAPH_CLIENT_SECRET`
- `MAYU_OPENAI_API_KEY`

## Ejecucion manual

En GitHub Actions, usar workflow **Deploy MAYU Agents**:

- `deploy_only`: publica sin prueba.
- `test`: envia correo de prueba.
- `weekly_report`: genera reporte semanal.
- `manual_due_sweep`: barre alertas pendientes.
- `monthly_report`: genera reporte mensual.

## Seguridad

No subir `mayu_mail_config.json` ni claves locales. Las credenciales viven como GitHub Secrets o Azure Automation encrypted variables.
