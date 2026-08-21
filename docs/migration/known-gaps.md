# Brechas Conocidas — INFOREST

> Inventario de brechas, pendientes y elementos que requieren atención durante la migración.
>
> Status: IN_PROGRESS — baseline .NET 8 de Fase 3 validado; los gaps remanentes se mantienen como backlog controlado para fases posteriores.

---

## Brechas Arquitectónicas

| ID | Descripción | Impacto | Prioridad | Estado |
|---|---|---|---|---|
| GAP-001 | Arquitectura .NET 8 aún no cubre módulos funcionales completos de operación | Alto | Alta | IN_PROGRESS — ADR-001 a ADR-012 aceptados y baseline transversal validado; faltan slices funcionales P3-04+ |
| GAP-002 | No existe todavía un módulo funcional completo listo para producción en .NET 8 | CRÍTICO | Alta | IN_PROGRESS — existe baseline ejecutable con seguridad/KDS/reportes/periféricos, pero ningún ejecutable reemplaza al Legacy end-to-end |
| GAP-003 | 206 reportes Crystal Reports — plan de migración definido, implementación en progreso | Alto | Alta | IN_PROGRESS — Etapa 10 |
| GAP-004 | Hardware POS (PinPad, fiscal, cajón) sin abstracción .NET | Alto | Alta | IN_PROGRESS — P3-11: PinPad (P/Invoke), CashDro y BlueVision implementados. Biometría SecuGen y Fiscal Epson BLOCKED (OCX 32-bit sin SDK .NET). |
| GAP-005 | Integraciones fiscales reales multi-país siguen pendientes pese a la arquitectura base ya definida | Alto | Alta | IN_PROGRESS — P3-11: IPaisPolicy + FacturacionElectronicaFactory + 5 gateways FE implementados como stubs. Integración real con SUNAT/AFIP/SII/SRI/SIAT pendiente. |
| GAP-006 | Base de datos Target no definida | Alto | Alta | OPEN |

---

## Brechas de Seguridad

| ID | Descripción | Archivo Legacy | Impacto | Prioridad |
|---|---|---|---|---|
| SEC-001 | Credenciales SQL hardcodeadas en código fuente | `modPuntoVenta.bas` Sub Main | CRÍTICO | Alta |
| SEC-002 | Cifrado débil XOR+César para datos sensibles | `ClsSeguridad.cls` | Alto | Alta |
| SEC-003 | SQL inline potencialmente vulnerable a injection | Múltiples `*.frm` | Alto | Media |
| SEC-004 | Variables globales sin control de acceso (543 vars) | `modDeclaracion.bas` | Medio | Media |
| SEC-005 | Licencia por dongle físico sin alternativa moderna | `License.cls`, `ModuloHardKey.bas` | Bajo | Baja |
| SEC-006 | `TUSUARIO.tPassword` (`varchar(8)`) no permite persistir BCrypt sin tabla sidecar — **RESUELTO**: tabla sidecar `TUSUARIO_HASH` implementada (ADR-013, script `migrations/001_TUSUARIO_HASH.sql`) | `TUSUARIO` | Alto | Alta — COMPLETED |

---

## Funcionalidades Legacy Sin Equivalente .NET

> Lista de componentes Legacy que siguen sin equivalente .NET completo o productivo.

| Componente Legacy | Tipo | Estado Target |
|---|---|---|
| Todos los 401 formularios VB6 | UI Forms | NOT_STARTED |
| Todos los 32 módulos BAS | Business Logic | NOT_STARTED |
| Todas las 10 clases | Domain Objects | NOT_STARTED |
| Todos los 206 reportes Crystal | Reports | NOT_STARTED |
| Integraciones hardware críticas (SecuGen, Epson fiscal, periféricos productivos) | Hardware | PARTIAL |

---

## Componentes de Propósito Incierto

> Componentes cuyo propósito no es determinable con certeza desde el código existente.

| Componente | Evidencia | Estado |
|---|---|---|
| `modMasticar.bas` | Nombre sugeriere procesamiento de datos, propósito exacto UNKNOWN | UNKNOWN |
| `modPvCorp.bas` | Posiblemente punto de venta corporativo, no confirmado | UNKNOWN |
| `Mapping.cls` | Mapeo de datos, uso exacto no determinable | UNKNOWN |
| `Infhotel.ini` | Integración con sistema hotelero, detalles UNKNOWN | UNKNOWN |

---

## Dependencias de Hardware Sin Plan de Migración

| Hardware | Módulo Legacy | Complejidad Migración | Plan .NET |
|---|---|---|---|
| Impresora térmica (ESC/POS) | TTIPODOCUMENTOIMPRESORA | Media | UNKNOWN |
| Cajón de dinero (RJ11) | API serial | Baja | UNKNOWN |
| CashDro (cajón automático) | `modProcedimientoNuevo.bas` | Media | `ICashDroService` + `CashDroService` (Process.Start). Validado en tests; wiring operativo real pendiente. |
| PinPad DLL3500 | `DLL3500.bas`, `CAJA_PINPAD.dll` | Alta — DLL Win32 | `IPinPadService` + `PinPadService` (P/Invoke). Falta validación con DLL/hardware real. |
| Impresora fiscal Epson | `modImpresoraFiscal.bas`, `IFEpson.ocx` | Alta — OCX Win32 | UNKNOWN |
| Biometría SecuGen | `FpLibX_Const.bas`, `sgfplibx.ocx` | Alta — OCX Win32 | `IBiometriaService` + `NullBiometriaService` (stub). UI en `FrmAcceso`: `btnHuella` (login 1:N), `btnMarcacion` (asistencia). Bloqueado hasta SDK .NET SecuGen disponible. GAP-004 BLOCKED. |
| KDS pantalla cocina | `modKDS.bas` | Media — XML sobre directorio | PARTIAL — existe baseline .NET para XML/directorio, falta pantalla operativa y confirmación con hardware real |
| BlueVision/TVS | `modBlueVision.bas` | Alta — COM Win32 | PARTIAL — `IBlueVisionService` + `BlueVisionHttpClient`; falta validar servidor/protocolo real en ambiente operativo |

---

## Integraciones Fiscales por País Sin Plan

| País | Integración | Módulo Legacy | Plan .NET |
|---|---|---|---|
| Perú | SUNAT FE, RUC, IGV | SP + `scriptPeruAlIniciar.sql` | `PeruFEGateway` stub + `PeruPaisPolicy`; integración OSE/SUNAT real pendiente |
| Chile | SII, IVA | `scriptChileAlIniciar.sql` | `ChileFEGateway` stub + `ChilePaisPolicy` |
| Bolivia | SIN código de control | `CodigoControl.bas` | `BoliviaFEGateway` stub + `BoliviaPaisPolicy`; código de control/fiscalidad real pendiente |
| Ecuador | SRI | `scriptEcuadorAlIniciar.sql` | `EcuadorFEGateway` stub + `EcuadorPaisPolicy` |
| Argentina | AFIP impresora fiscal | `modImpresoraFiscal.bas` | `ArgentinaFEGateway` stub + `ArgentinaPaisPolicy`; impresora fiscal sigue BLOCKED |
| España | IVA España | `scriptEspanaAlIniciar.sql` | `EspanaPaisPolicy`; gateway FE dedicado no implementado |

---

## TODOs Detectados en el Análisis

| Elemento | Descripción | Archivo/Ubicación |
|---|---|---|
| `modMasticar.bas` | Propósito no determinado — requiere revisión | `Modulos/modMasticar.bas` |
| `modPvCorp.bas` | Propósito no determinado — requiere revisión | `Modulos/modPvCorp.bas` |
| `Mapping.cls` | Uso no determinado — requiere revisión | `Clases/Mapping.cls` |
| `INFHOTEL.INI` | Integración hotelera — detalles sin documentar | Configuración runtime |
| Triggers SQL | Triggers eliminados en runtime (`ElminaTriguer`) — existencia real no confirmada | `modProcedimiento.bas` |
| HardKey físico | La validación de presencia/conexión del dongle (`hk.VerificaConexion`, `hk.IniciaConexion`) aún no tiene equivalente .NET | `frmAcceso.frm`, `License.cls` |

---

## Deuda Técnica del Legacy (A Resolver en Target)

| Deuda | Descripción | Archivo Legacy |
|---|---|---|
| Código spagueti | UI + lógica + datos mezclados en formularios | Todos los `.frm` |
| 543 variables globales | Estado global no controlado | `modDeclaracion.bas` |
| Hardcoding generalizado | Strings literales, credenciales, rutas | Múltiples módulos |
| Sin pruebas automatizadas | Cero tests en el Legacy | — |
| Crystal Reports dependiente | Motor de reportes propietario y legado | `Reportes/` |
| Timeout de 600s en SPs | `clsComando` — SP timeout muy alto | `Clases/clsComando.cls` |
| Timers para polling | Timer-based polling (250ms–60000ms) | Múltiples formularios |

---

*Este documento debe actualizarse a medida que se identifican nuevas brechas durante la migración.*

---

## Etapa POS-FUNC-008 — Caja y Cierre (frmLiquidacionDetalle.frm)

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-CAJA-001 | frmDescargo navigation | Al responder "No" en la confirmación de descargo (BR-CAJA-002, lActivaConsultaDescargo), el Legacy abre `frmDescargo` para ejecutar el descargo inmediatamente. En .NET esta navegación está pendiente hasta que `frmDescargo` sea migrado. | `frmLiquidacionDetalle.frm` línea 6062 | OPEN — depende de migración de `frmDescargo` |
| GAP-CAJA-002 | Integración Niubiz/Izipay al cierre (BR-CAJA-005) | Al cerrar el turno con `lActivaIntegracionNiubiz=1` o `lActivaIntegracionIzipay=1`, el Legacy inserta en `MCIERREPOST` y llama `SP_EJECUTA_MOTOR_INALAMBRICO`. Pendiente de decisión de hardware/integración de terminal POS externo. | `frmLiquidacionDetalle.frm` líneas 6076-6169 | BLOCKED — requiere decisión de integración terminal POS |

---

## Etapa 9 — Delivery, Despacho, Motorizados

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-DEL-001 | sp_UpdFotoDelivery | Actualización de foto del cliente delivery migrada a `ActualizarFotoClienteDeliveryHandler` + `ClienteDeliveryRepository.ActualizarFotoAsync` usando SP legacy y habilitada en `NuevoDeliveryForm` (alta/edición) | `frmClienteDeliveryDetalle.frm`, `frmNuevoDelivery.frm` | CLOSED |
| GAP-DEL-002 | ADR-009 pendiente | Estrategia de sincronización multi-local: polling actual a CENTRALDELIVERY no escala. Decidir entre polling, Service Bus o SQL replication | `modDespachador.bas` | OPEN |
| GAP-DEL-003 | ADR-010 pendiente | Integración Rappi: actualmente stub. Definir si se mantiene adaptador embebido o se expone endpoint | `frmOrdenesConsola.frm`, `frmEnvioDatosRappi.frm` | OPEN |
| GAP-DEL-004 | RappiOrderAdapter HTTP | `RappiOrderAdapter` retorna lista vacía. El conector HTTP real a la API de Rappi no está implementado | `Inforest.Infrastructure/Delivery/RappiOrderAdapter.cs` | NOT_STARTED |
| GAP-DEL-005 | Reportes FastReport Etapa 9 | 8 reportes del módulo Despachador/Motorizado: RepAnaliticoMotorizado, RepControlMotorizado, RepTiempoDelivery, RepPlanillaMovilidad, RepClienteDelivery, RepPedido, variantes integradas — handlers y repositorio implementados en Etapa 10; plantillas .frx pendientes | `Reportes/` | IN_PROGRESS — Etapa 10 |
| GAP-DEL-006 | Formularios WinForms faltantes | Pendiente operacional principal en Delivery: `OrdenesConsolaForm`. `FrmClienteDelivery`, `FrmPedidoDelivery` y `FrmBusquedaDelivery` ya fueron migrados. | `frmOrdenesConsola.frm` | IN_PROGRESS |

---

## POS-FUNC-017 — frmRepComanda (spRep_Comanda)

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-REP-020 | NCantidad no mapeado en modo Resumido | En modo Resumido (`@flagTipo=0`), el SP retorna la columna `Cantidad` (alias de `Sum(DPEDIDO.nCantidad)`). Dapper no la mapea a `ComandaRow.NCantidad` porque los nombres difieren. El valor de cantidad queda en 0 en el DTO Resumido. Se mantiene el mismo DTO `ComandaRow` para ambos modos. Solución futura: SP puede renombrar alias a `nCantidad` o usar un DTO separado `ComandaResumidoRow`. | `5. SP.sql` — `spRep_Comanda` modo `@flagTipo=0` | OPEN — no crítico, afecta solo la columna cantidad en Resumido |

---

## POS-FUNC-017 — frmRepLiquidacion (Liquidación de Cajero)

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-REP-021 | Diálogo búsqueda de turno (frmBusca) | El botón "..." en FrmRepLiquidacionReporte abre `frmBusca` en VB6 mostrando MTURNO con columnas Turno/Caja/FechaInicial/FechaFinal. En .NET 8 el botón solo otorga foco al campo de texto. Pendiente: implementar un diálogo de búsqueda de turnos equivalente a `frmBusca` reutilizable. | `frmRepLiquidacion.frm` — cmdBusca_Click | OPEN — el usuario puede ingresar el código de turno manualmente |

---

## POS-FUNC-017 — frmRepRegistroVenta (Registro de Ventas)

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-REP-022 | SP `spRep_RegVentaSunat_formaPago` no encontrado | El tipo 7 (Correlativo con Forma de Pago) de `frmRepRegistroVenta.frm` llama a `spRep_RegVentaSunat_formaPago`. Este SP no existe en `5. SP.sql` ni en los scripts opcionales. El handler devuelve resultado vacío para este tipo y la forma muestra un mensaje de advertencia. | `frmRepRegistroVenta.frm` — Sub Genera4() | OPEN — tipo 7 deshabilitado hasta que se encuentre o cree el SP |

---

## POS-FUNC-021 — Tarjetas RFID / Proximidad

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-RFID-001 | Flujo de recarga parcial vs legado | `FrmRecargarTarjeta.frm` + `FrmRecargarTarjetaDetalle.frm` ya migran listado de recargas y actualización transaccional de saldo RFID (`TMOVIMIENTOTARJETASRFID` + `TTARJETASRFID`). Sigue pendiente cerrar equivalencia completa del bloque legacy de anticipo/emisión documental/impresión fiscal dentro del mismo flujo. | `FrmRecargarTarjeta.frm`, `FrmRecargarTarjetaDetalle.frm` | IN_PROGRESS |
| GAP-RFID-002 | `FrmMovimientoTarjetas.frm` sin lógica relevante | El formulario `FrmMovimientoTarjetas.frm` aparece en `InfoRest.vbp` como cascarón visual sin código operativo útil. Se mantiene fuera del corte hasta confirmar si corresponde a un flujo real o a una pantalla obsoleta. | `FrmMovimientoTarjetas.frm` | ANALYSIS |

---

## POS-FUNC-006 — Generación de Documentos (frmDocumento.frm)

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-DOC-001 | Anulación de pagos de documento | `cmdOpcion_Click` Case 11 en frmDocumento.frm implementa la anulación de pagos con un flujo transaccional complejo: requiere supervisor "05", anula los registros en DPAGODOCUMENTO, revierte tEstadoDocumento a '01', y re-abre el pedido. En .NET 8 este flujo no está migrado. El botón "Anulación de Pagos" está temporalmente excluido de FrmDocumento. | `frmDocumento.frm` — cmdOpcion_Click Case 11 | OPEN |
| GAP-DOC-002 | Permiso de supervisor en reimpresión | En la versión legacy, `cmdOpcion_Click` Case 7 (Reimprimir) requiere validar `Supervisor("12")` antes de proceder. En .NET 8 la validación de permisos por código de supervisor no está implementada en FrmDocumento. | `frmDocumento.frm` — cmdOpcion_Click Case 7 | OPEN |
| GAP-DOC-003 | Integración Niubiz/Izipay en operaciones | frmDocumento.frm verifica integración con Niubiz (TCAJA.lActivaIntegracionNiubiz) antes de permitir ciertas operaciones. En .NET 8 esta comprobación no está implementada en FrmDocumento. | `frmDocumento.frm` — cmdOpcion_Click (validación lActivaIntegracionNiubiz) | OPEN |

---

## mdiPuntoVenta.frm — Formulario MDI Principal POS

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-MDI-001 | frmAnulado / frmCancelado | `mnuAnulado_Click` y `mnuCancelado_Click` abren formularios de consulta de anulados/cancelados. No existen en la implementación .NET 8 actual. | `mdiPuntoVenta.frm` — mnuAnulado_Click, mnuCancelado_Click | OPEN |
| GAP-MDI-002 | frmConsultaSaldo (SIAB) | `cmdConsultaSaldo_Click` abre formulario de consulta de saldo. No existe en .NET 8. El botón está oculto por defecto (lSiab=false). | `mdiPuntoVenta.frm` — cmdConsultaSaldo_Click | OPEN |
| GAP-MDI-003 | frmServidorEnlace (multiLocal) | `mnuCambiaLocal_Click` abre frmServidorEnlace para cambio de local. No existe en .NET 8. El menú solo aparece si lMultilocal=true. | `mdiPuntoVenta.frm` — mnuCambiaLocal_Click | OPEN |
| GAP-MDI-004 | frmGuiaTransporte | `mnuGuiaTransporte_Click` abre frmGuiaTransporte. No existe en .NET 8. | `mdiPuntoVenta.frm` — mnuGuiaTransporte_Click | OPEN |
| GAP-MDI-005 | frmDocumentoElectronicoCorrelativo | `mnuDocumentoElectronico_Click` abre frmDocumentoElectronicoCorrelativo. No existe en .NET 8. | `mdiPuntoVenta.frm` — mnuDocumentoElectronico_Click | OPEN |
| GAP-MDI-006 | frmMozoUsuario (MCPV) | En modo MCPV (`lMCPV=true`), el formulario oculta la ventana y abre frmMozoUsuario antes de mostrar el MDI. No implementado en .NET 8. El modo lMCPV se detecta desde ConfiguracionCaja. | `mdiPuntoVenta.frm` — MDIForm_Load lMCPV, cmdOpcion16_Click | OPEN |
| GAP-MDI-007 | PinPad VisaNet completo | El flujo completo de PinPad no-financiera (fiStartOperation + bucle fiGetStatus + MensajePinPad) requiere DLL nativa VisaNet no portada. FrmPagoPinPad existe pero el flujo legacy de no-financiera es una operación de consulta de estado diferente. | `mdiPuntoVenta.frm` — mnuPinPad_Click | OPEN |
| GAP-MDI-008 | comandaAuto (patio orders) | `Timerwebapp_Timer` con lIntPatio=true ejecuta `comandaAuto()` que imprime automáticamente pedidos de patio. La lógica requiere integración con el módulo de impresión (vAreaImpresora, ImprimePedidoAuto). No implementado en .NET 8. | `mdiPuntoVenta.frm` — Timerwebapp_Timer, comandaAuto() | OPEN |
| GAP-MDI-009 | InicializaMCPV | Modo multi-cajero: busca turno activo del usuario y actualiza estado. Requiere lMCPV=true. Pendiente de implementar flujo completo. | `mdiPuntoVenta.frm` — InicializaMCPV() | OPEN |
| GAP-MDI-010 | FrmActualizacion (versión) | TimerActualizador compara versión del .exe con TPARAMETROVERSION y lanza FrmActualizacion si hay diferencia. No implementado en .NET 8. | `mdiPuntoVenta.frm` — TimerActualizador_Timer | OPEN |
