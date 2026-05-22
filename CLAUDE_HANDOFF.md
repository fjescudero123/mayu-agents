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

## BICE cartola por correo

- Modo: `bice_cartola_mail`.
- Runtime: GitHub Actions en `run-mayu-agents.yml`.
- Schedule: lunes a viernes 15:10 UTC, despues de la llegada usual del correo BICE.
- Lee `notificaciones@imayu.cl` via Microsoft Graph.
- Descarga adjuntos BICE `.dat.xls`, guarda crudos en SharePoint bajo `agentes_mayu/bice_cartolas/raw/YYYY-MM-DD/`.
- Parser productivo para Excel antiguo (`application/vnd.ms-excel`) via `xlrd`.
- Deduplica contra `fin_mov_bancarios` y dentro del mismo lote antes de escribir.
- Crea solo movimientos nuevos con fecha, monto y direccion clara; filas ambiguas quedan en reporte.
- Registra `fin_importaciones` con `tipo=CARTOLA_BICE` y `origen=BICE_EMAIL`.

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
