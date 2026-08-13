# Módulo Motorizado — INFOREST

## Descripción

Módulo de gestión de motorizados de delivery. Corre como ejecutable independiente `Motorizado.exe`.
Permite controlar la llegada y salida de motorizados, ver pedidos asignados, gestionar reasignaciones y registrar tarifas.

## Ejecutable Legacy

`Motorizado.exe` — `Motorizados.vbp`

## Formularios Legacy

| Formulario | Descripción | Migrado |
|---|---|---|
| `mdiMotorizado.frm` | MDI principal del módulo | IN_PROGRESS |
| `frmLlegadaSalida.frm` | **Startup**: Control de llegada/salida motorizado | .NET: `LlegadaSalidaForm` |
| `frmMotorizado.frm` | Grilla de motorizados activos | IN_PROGRESS |
| `frmMotorizadoDetalle.frm` | Detalle de motorizado | IN_PROGRESS |
| `frmAsignacionMotorizado.frm` | Asignación de motorizado a pedido | .NET: `AsignacionMotorizadoForm` |
| `frmReasignacionMotorizado.frm` | Reasignación de motorizado | .NET: `ReasignacionMotorizadoForm` |
| `frmTarifaMotorizado.frm` | Gestión tarifa activa (requiere supervisión) | .NET: `TarifaMotorizadoForm` |
| `fmrPedidosEntregado.frm` | Registro de pedidos entregados | IN_PROGRESS |
| `frmRepAnaliticoMotorizado.frm` | Reporte analítico | NOT_STARTED |
| `frmRepAnaliticoMotorizadoIntegrado.frm` | Reporte analítico multi-local | NOT_STARTED |
| `frmRepControlMotorizado.frm` | Reporte de control | NOT_STARTED |
| `frmRepPlanillaMovilidadMotorizado.frm` | Planilla de movilidad | NOT_STARTED |
| `frmRepTiempoDelivery.frm` | Reporte tiempos de delivery | NOT_STARTED |
| `frmRepTiempoDeliveryIntegrado.frm` | Tiempos multi-local | NOT_STARTED |

## Módulo BAS

| Módulo | Descripción |
|---|---|
| `modMotorizado.bas` | Sub Main — conecta BD, carga TPARAMETRO, muestra frmLlegadaSalida |
| `modDeclaracion.bas` | Variables globales |
| `modProcedimiento.bas` | Utilidades núcleo |
| `modConexionIp.bas` | Conexión IP |

## Reglas de Negocio

| ID | Nombre | Estado |
|---|---|---|
| BR-DEL-002 | Tarifa varía por tipo de día: LV/SD/ES (`TMOTORIZADODATOS`) | MIGRATED |
| BR-DEL-005 | `tTarifaActualMotorizado` en TPARAMETRO indica tarifa vigente | MIGRATED |
| BR-DEL-006 | Exe inicia directamente en frmLlegadaSalida | MIGRATED |

## Comportamiento de Startup (BR-DEL-006)

```
Sub Main() →
  1. Lee INFOREST.INI (servidor, BD, caja, salon, empresa, CentralDelivery)
  2. Abre conexión INFOREST
  3. Lee flag lCD de TCAJA
  4. Lee TPARAMETRO (RUC, nLongitud)
  5. Verifica versión del exe
  6. frmLlegadaSalida.Show   ← pantalla directa, sin login previo
```

## Estado de Migración

| Componente | Estado |
|---|---|
| Entidades Domain (Motorizado, AsignacionMotorizado) | COMPLETED |
| Interfaces repositorio | COMPLETED |
| Application handlers | COMPLETED |
| Infrastructure MotorizadoRepository | COMPLETED |
| AsignacionMotorizadoForm | COMPLETED |
| ReasignacionMotorizadoForm | COMPLETED |
| LlegadaSalidaForm | COMPLETED |
| TarifaMotorizadoForm | COMPLETED |
| Reportes FastReport | NOT_STARTED |
| Tests | IN_PROGRESS |
