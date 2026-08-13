# Módulo Despachador — INFOREST

## Descripción

Módulo de despacho y central de pedidos del sistema INFOREST. Corre como ejecutable independiente `Despachador.exe`.
Permite visualizar, gestionar y despachar los pedidos de delivery, integra con la base `CENTRALDELIVERY` para operación multi-local y soporta órdenes externas de Rappi.

## Ejecutable Legacy

`Despachador.exe` — `Despachador.vbp`

## Formularios Legacy

| Formulario | Descripción | Migrado |
|---|---|---|
| `mdiDespachador.frm` | MDI principal del módulo | .NET: `DespachadorForm` |
| `frmDespachador.frm` | Grilla de pedidos en despacho (vDespachador) | .NET: `DespachadorForm` |
| `frmCentralPedidos.frm` | Correlativo Pedidos Central | .NET: `CentralPedidosForm` |
| `frmBusquedaCentralPedido.frm` | Búsqueda en base CENTRALDELIVERY | IN_PROGRESS |
| `frmNuevoDelivery.frm` | Alta/edición cliente delivery frecuente | .NET: `NuevoDeliveryForm` |
| `frmPedidoDelivery.frm` | Detalle de pedido delivery | IN_PROGRESS |
| `frmClienteDelivery.frm` | Lista clientes delivery | IN_PROGRESS |
| `frmClienteDeliveryDetalle.frm` | Detalle de cliente delivery | IN_PROGRESS |
| `frmClienteDeliveryFrecuente.frm` | Clientes frecuentes | IN_PROGRESS |
| `frmBusquedaDelivery.frm` | Búsqueda de delivery | IN_PROGRESS |
| `frmOrdenesConsola.frm` | Panel de órdenes externas (Rappi) | IN_PROGRESS |
| `frmEnvioDatosRappi.frm` | Envío de datos a Rappi | IN_PROGRESS |
| `frmAsignacionMotorizado.frm` | Asignación de motorizado | .NET: `AsignacionMotorizadoForm` |
| `frmReasignacionMotorizado.frm` | Reasignación de motorizado | .NET: `ReasignacionMotorizadoForm` |
| `frmTarifaMotorizado.frm` | Gestión tarifa activa | .NET: `TarifaMotorizadoForm` |
| `frmPedido.frm` | Pedido en contexto Despachador | IN_PROGRESS |
| `frmDetallePedido.frm` | Detalle de pedido | IN_PROGRESS |
| `frmBuscaPedido.frm` | Búsqueda de pedido | IN_PROGRESS |

## Módulos BAS

| Módulo | Descripción |
|---|---|
| `modDespachador.bas` | Sub Main, configuración de conexiones, flags TCAJA |
| `modDeclaracion.bas` | Variables globales (543 vars) |
| `modProcedimiento.bas` | Utilidades núcleo |
| `modConexionIp.bas` | Conexión IP para comunicación entre cajas |
| `modBlueVision.bas` | Integración BlueVision TVS |
| `modAuditoriaIntegral.bas` | Auditoría |

## Tablas SQL Principales

| Tabla | Descripción |
|---|---|
| `TDELIVERY` | Clientes de delivery |
| `TDELIVERYCLIENTE` | Relación cliente delivery ↔ cliente cuenta corriente |
| `TDELIVERYINVITADO` | Invitados del cliente delivery |
| `MPEDIDO` | Cabecera de pedido (incluye campos delivery) |
| `DPEDIDO` | Detalle de pedido |
| `TMOTORIZADODATOS` | Datos del motorizado con tarifas LV/SD/ES |
| `TLOCAL` | Locales del restaurante (multi-local) |
| `TCAJA` | Configuración de caja (flags lCD, lOrdenesRappi) |
| `TPARAMETRO` | Parámetros globales (tTarifaActualMotorizado) |
| `TTIPOPEDIDODETALLE` | Flags de comportamiento por tipo de pedido |

## Vistas SQL

| Vista | Descripción |
|---|---|
| `vDespachador` | Panel de despacho con pedidos tipo 02, estado 02 |
| `vDelivery` | Pedidos delivery con datos de cliente |
| `vMotorizado` | Motorizados y asignaciones |
| `vDistrito` | Distritos activos |
| `vZona` | Zonas activas |
| `vTablasCentralizada` | Datos del Central de Pedidos |
| `vTipoPedido` | Tipos de pedido activos |
| `vPedidoCabecera` | Cabecera de pedido |
| `vPedidoDetalle` | Detalle de pedido |
| `vPedidoGrilla` | Grilla consolidada de pedidos |
| `vPreCuentaDelivery` | Pre-cuenta delivery |
| `vEstadoFrecuente` | Clientes frecuentes con estado |

## Stored Procedures

| SP | Descripción |
|---|---|
| `sp_Inforest_PedidosCentralPedido` | Consulta de pedidos en CENTRALDELIVERY |
| `sp_CD_Modificar_EstadoDelivery_Cabecera` | Cambia estado en CENTRALDELIVERY |
| `SP_COVER_INSERTA_PEDIDO_ADICION` | Cobertura de adición al pedido |
| `spUpdate_DPEDIDO_Ina` | Inactiva ítem de detalle de pedido |
| `sp_UpdFotoDelivery` | Actualiza foto del delivery |
| `spRep_AnaliticoMotorizadoIntegrado` | Reporte analítico de motorizado integrado |

## Reglas de Negocio

| ID | Nombre | Estado |
|---|---|---|
| BR-DEL-001 | Tipo de pedido determina si requiere motorizado y cliente frecuente | MIGRATED |
| BR-DEL-002 | Tarifa motorizado varía por tipo de día: LV/SD/ES | MIGRATED |
| BR-DEL-003 | Flag `lCD` en TCAJA habilita modo CENTRALDELIVERY | MIGRATED |
| BR-DEL-004 | Flag `lOrdenesRappi` en TCAJA habilita integración Rappi | MIGRATED |
| BR-DEL-005 | `tTarifaActualMotorizado` en TPARAMETRO indica tarifa vigente | MIGRATED |
| BR-DEL-006 | Motorizado.exe inicia directamente en frmLlegadaSalida | MIGRATED |
| BR-DEL-007 | TLOCAL controla multi-local: IP, BD INFOREST por local, flag réplica | MIGRATED |
| BR-DEL-008 | Central Pedidos usa conexión separada a CENTRALDELIVERY | MIGRATED |
| BR-DEL-009 | Estado delivery se modifica vía sp_CD_Modificar_EstadoDelivery_Cabecera | MIGRATED |
| BR-DEL-010 | Tipo de pedido tiene flags canal: lCanalCentralPedidos, lCanalDelivery | MIGRATED |

## Decisiones Arquitectónicas Pendientes

- **ADR-009**: Estrategia sincronización multi-local (polling vs. Service Bus vs. SQL replication)
- **ADR-010**: Integración Rappi (adaptador embebido vs. webhook/API endpoint)

## Estado de Migración

| Componente | Estado |
|---|---|
| Entidades Domain | IN_PROGRESS |
| Interfaces repositorio | COMPLETED |
| Application handlers | COMPLETED |
| Infrastructure repositorios | IN_PROGRESS |
| WinForms Desktop | IN_PROGRESS |
| Reportes FastReport | NOT_STARTED |
| Tests | IN_PROGRESS |

## Gaps Conocidos

- `sp_UpdFotoDelivery` — backend migrado (handler + repositorio con SP legacy); falta formulario dedicado equivalente a `frmClienteDeliveryDetalle.frm`
- Integración Rappi API HTTP — conector real pendiente (ver docs/migration/known-gaps.md GAP-DEL-004)
- ADR-009 no resuelto: sincronización multi-local actual mantiene conexión directa a CENTRALDELIVERY
