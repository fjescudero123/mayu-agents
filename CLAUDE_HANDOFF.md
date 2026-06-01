# MAYU Agents Handoff

Actualizado: 2026-06-01

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

### Administrador Bodega + Materiales MAYU

- Modo principal: `bodega_materiales_admin`.
- Respondedor: `bodega_materiales_respuestas`.
- Estado: operativo Nivel 2 asistido.
- Firestore: casos persistentes en `bma_admin_cases`.
- SharePoint: `agentes_mayu/bodega_materiales/administrador`.
- Correo: desde `notificaciones@imayu.cl` a Carlos y Valentina, con Felix en copia.
- Codigos de caso: `BMA-XXXXXXXX`.
- Respuestas esperadas:
  - `BMA-12345678: A`
  - `BMA-12345678: corregir: [criterio]`
  - `BMA-12345678: bloquear`
  - `BMA-12345678: rechazar`
- Seguridad: no modifica OCs, stock, costos, recepciones, entregas a fabrica ni movimientos por correo. Solo registra decision/aprendizaje hasta que exista regla aprobada y evidencia suficiente.
- Reenvio con lenguaje simplificado validado el 2026-06-01 en run `26732517160`: 807 casos, 12 preguntas, 6 reglas candidatas, 0 autoejecuciones; correo enviado a `clecaros@imayu.cl` y `vescudero@imayu.cl`.

### Patron de lenguaje para agentes con responsables operativos

Este patron debe replicarse en otros agentes cuando la pregunta va a personas operativas, no tecnicas.

- No exponer campos internos como `itemCode`, `matchCode`, `skuCode`, `matProjectId`, nombres de colecciones o jerga de base de datos.
- Traducir siempre a objetos de negocio: OC, proyecto, proveedor, descripcion del producto, cantidad, factura, recepcion, entrega a fabrica.
- Si se necesita conservar un codigo, mostrarlo como dato secundario: `codigo interno: EST-0035`.
- Estructura recomendada del caso:
  - Codigo de caso y titulo simple.
  - Contexto: `OC: ... | Proyecto: ... | Proveedor: ...`.
  - Producto por descripcion, no por codigo.
  - `El problema:` en una frase operacional.
  - `Que necesito:` la decision concreta que debe tomar el responsable.
  - Nota de seguridad: el agente solo guarda la respuesta; no cambia datos sensibles.
- Alternativas deben ser decisiones naturales:
  - `A) Corregir producto en la OC`
  - `B) Autorizar excepcion`
  - `C) Dejar pendiente para revision`
  - `D) No hay problema`
- Usar "reglas candidatas" en vez de "skills candidatas" en correos a usuarios.
- Para Finanzas se puede tolerar mas lenguaje contable; para Carlos/Bodega se debe priorizar descripcion fisica y accion concreta.

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
