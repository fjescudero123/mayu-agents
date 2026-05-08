# MAYU Agents - Arquitectura operacional

Actualizado: 2026-05-08

## Diagnostico del runtime actual

El runtime vigente del fiscalizador de reuniones esta sano y debe convivir sin cambios:

- GitHub Actions ejecuta PowerShell directo.
- Config no sensible vive en `config/fiscalizador_config.json`.
- Secretos y variables se leen desde GitHub Secrets/Variables.
- Microsoft Graph resuelve token, SharePoint y correo.
- OpenAI se usa como capa opcional de lectura inteligente.
- Azure Automation queda como legacy pausado, no como runtime normal.

La nueva arquitectura de agentes mantiene ese patron, pero separa runtime, config y workflow:

- Runtime nuevo: `runtime/MayuAgents.core.ps1`.
- Config nuevo: `config/mayu_agents_config.json`.
- Workflow nuevo: `.github/workflows/run-mayu-agents.yml`.
- Output SharePoint nuevo: `agentes_mayu/...`.

## Principio operacional

Control Documental (`chk_projects`) es la fuente madre de la operacion.

La "biblia del proyecto" se valida desde documentos aprobados/versionados:

- BOM.
- Planos de arquitectura.
- Planos de fabricacion.
- Carta Gantt de fabricacion.
- Plan de despachos.
- Carta Gantt de compras.
- Protocolo transporte.
- Archivo calidad.
- Analisis financiero.

El flujo vigente es:

1. BOM aprobado/versionado en Control Documental.
2. Materiales sincroniza el BOM a `mat_projects.bomVersionUsada`.
3. Fabricacion queda alineada en `fab_proyecto_recetas.bomVersionSincronizada`.
4. Fabricacion genera `fab_units` y `fab_packs`.
5. Bodega arma/entrega packs y descuenta stock.
6. Finanzas cruza OCs, facturas, caja y costos por proyecto.

No se usa como concepto operativo "congelar recetas".

## Mapa de colecciones Firestore

### Fuente madre

- `chk_projects`: Control Documental. Proyecto, CRM link, areas, docs, versiones, archivos, aprobaciones y estado general.
- `chk_productos_tipo`: productos tipo, fuera del pulso V1 salvo evolucion futura.
- `chk_users`: usuarios Control Documental, no requerido para el pulso V1.

### Comercial

- `projects`: CRM. Estado comercial, linea, ingreso/costo proyectado, fechas y responsable.
- `plan`: meta anual comercial.
- `billing`: metas mensuales y facturacion CRM legacy.

### Materiales y compras

- `mat_projects`: proyecto operativo de Materiales; puente entre CRM, Control y BOM.
- `mat_cotizaciones`: cotizaciones por item, no requerida para pulso V1.
- `mat_ordenes`: OCs por item/BOM/SKU.
- `mat_recepciones`: recepciones contra OCs.
- `mat_entregas`: legacy entregas a fabrica, hoy desplazado por packs.

### Bodega

- `inv_catalogo`: SKU maestro, stock, costos, `aliasDe`, activo/inactivo.
- `inv_movimientos`: kardex valorizado.
- `inv_solicitudes`: solicitudes legacy/directas de materiales.

### Fabricacion

- `fab_units`: PODs/unidades individuales, estado real y RF.
- `fab_step_logs`: historial de avance por modulo.
- `fab_proyecto_recetas`: receta operacional sincronizada desde BOM.
- `fab_packs`: packs por fecha/modulo, compromiso bodega, armado, entrega, recepcion y cierre.
- `fab_recetas_cambios`: auditoria de cambios de receta.

### Finanzas

- `fin_facturas_ap`: CxP, facturas proveedores, clasificacion y links a OC.
- `fin_facturas_ar`: CxC, facturas clientes.
- `fin_pagos`: pagos AP.
- `fin_cobranzas`: cobros AR.
- `fin_mov_bancarios`: cartola.
- `fin_saldos_bancarios`: saldo base por periodo.
- `fin_proveedores`: maestro proveedores, stubs e higiene.
- `fin_clientes`: maestro clientes.
- `fin_opex`: gastos recurrentes.
- `fin_nomina_mensual`: nomina.
- `fin_tarjetas_credito`: pasivos TC.
- `fin_proyectos_cerrados`: cierre formal por proyecto.
- `fin_importaciones`: auditoria de imports.

## Diseno de config JSON

El config comun declara:

- `timezone`.
- `sharepoint`: host, carpeta base, links de apps.
- `mail`: sender y destinatarios por agente.
- `firebase`: projectId y nombre de variable donde vive el API key.
- `agents`: frecuencia conceptual, severidad y limites de cada agente.
- `collections`: nombres canonicos de Firestore.
- `doc_requirements`: equivalencia entre Biblia y documentos Control.
- `thresholds`: vencimientos, stock critico, semanas, limite de filas.

El API key de Firebase se toma de `MAYU_FIREBASE_API_KEY` o `MayuFirebaseApiKey`.
No es un secreto de servidor, pero se configura como variable para no duplicar valores en codigo.

## Outputs SharePoint y correo

Raiz propuesta: `agentes_mayu`.

- `agentes_mayu/pulso_gerencial/YYYY-MM-DD.html`.
- `agentes_mayu/pulso_gerencial/YYYY-MM-DD.json`.
- `agentes_mayu/biblia_proyecto/YYYY-MM-DD.json`.
- `agentes_mayu/traspaso_control_operacion/YYYY-MM-DD.json`.
- `agentes_mayu/abastecimiento/YYYY-MM-DD.json`.
- `agentes_mayu/higiene_erp/YYYY-MM-DD.json`.
- `agentes_mayu/caja_costos_directorio/YYYY-MM-DD.json`.

Correo diario del Pulso:

- Para: Felix.
- CC: opcional Valentina en secciones financieras futuras.
- Subject: `Pulso gerencial MAYU - YYYY-MM-DD - R/N A/M`.
- Formato corto: decisiones Felix hoy, rojos, amarillos, responsables/acciones y referencias.

## Plan por PRs pequenos

1. PR 1 - Esqueleto runtime y Pulso V1:
   - Runtime/config/workflow separados.
   - Lectura Firestore por REST.
   - Pulso con secciones parciales deterministicas.
   - Biblia y Traspaso como fuentes internas del pulso.

2. PR 2 - Biblia del Proyecto completa:
   - Semaforo por etapa: comprar, fabricar, despachar, RF.
   - JSON propio por proyecto.
   - Links profundos a Control/Materiales cuando existan.

3. PR 3 - Traspaso Control -> Operacion:
   - Reglas de BOM aprobado vs sincronizado.
   - `mat_project` sin `chkId`/`crmId`.
   - Receta desfasada.
   - Proyectos aprobados sin unidades/packs.

4. PR 4 - Packs y Abastecimiento:
   - Destrabador de Packs.
   - Stock + OCs + recepciones.
   - Items BOM sin SKU y SKUs costo cero.

5. PR 5 - Finanzas / Caja / Directorio:
   - CxP, CxC, caja 13 semanas y presion de OCs.
   - Costos por proyecto vs presupuesto.

6. PR 6 - Higiene ERP y Calidad/RF:
   - Higiene semanal.
   - RF, calidad, despacho y protocolo transporte.

## Estado implementado en este bloque

El primer bloque implementa el Pulso Gerencial Diario como agregador central con fuentes parciales:

- Biblia del Proyecto.
- Traspaso Control -> Operacion.
- Packs.
- Abastecimiento.
- Finanzas.
- Comercial.
- Calidad/RF.

Las secciones que aun no tengan datos suficientes aparecen como parciales, no inventadas.
