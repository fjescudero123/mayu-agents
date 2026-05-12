# MAYU Agents Handoff

Actualizado: 2026-05-12

## Estado actual

- Runtime operativo: GitHub Actions.
- Flujos activos:
  - `.github/workflows/run-fiscalizador.yml`
  - `.github/workflows/run-mayu-agents.yml`
  - `.github/workflows/run-minutas.yml`
- SharePoint y correo: via Microsoft Graph.
- IA: OpenAI via GitHub Secrets.

## Fuera de Azure

Este repo ya no usa:

- Azure Automation
- Logic Apps
- Azure login
- OIDC de despliegue a Azure
- runbooks
- workflows de deploy a Azure

## Minutas

- Pipeline activo: `run-minutas.yml`
- Runtime: `runtime/MinutasMAYU.core.ps1`
- Config: `config/minutas_config.json`
- Entrada: `minutas_entrada`
- Salida: `minutas_archivadas`
- Envio: `notificaciones@imayu.cl`

Validacion reciente:

- `AGENDA_PROYECTOS_12_05_2026.docx` procesado correctamente.
- Se genero `minuta_proyectos_2026-05-12.html`.
- El formato fue ajustado para volver al estilo visual historico.

## Fiscalizador

- Pipeline activo: `run-fiscalizador.yml`
- Runtime: `runtime/FiscalizadorReunionesMAYU.core.ps1`
- Config: `config/fiscalizador_config.json`

## Agentes operativos

- Pipeline activo: `run-mayu-agents.yml`
- Runtime: `runtime/MayuAgents.core.ps1`
- Config: `config/mayu_agents_config.json`

## Variables y secrets requeridos

Variables:

- `MAYU_GRAPH_TENANT_ID`
- `MAYU_GRAPH_CLIENT_ID`
- `MAYU_OPENAI_MODEL`
- `MAYU_FIREBASE_API_KEY`

Secrets:

- `MAYU_GRAPH_CLIENT_SECRET`
- `MAYU_OPENAI_API_KEY`

## Nota importante

Todavia existe dependencia de Microsoft Entra / Graph para autenticar SharePoint y correo. Eso es parte de Microsoft 365 y no del runtime Azure que se retiro.
