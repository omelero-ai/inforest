# Matriz de Trazabilidad — INFOREST Legacy → .NET 8

> Status: IN_PROGRESS — baseline transversal de Fase 3 validado; la matriz refleja equivalencias parciales y gaps controlados.
>
> Última actualización: 2026-08-17

---

## Instrucciones de Uso

Esta matriz relaciona cada componente Legacy con su equivalente en .NET 8.

**Actualizar esta matriz obligatoriamente** cuando:
- Se inicia la migración de un componente (`ANALYSIS` → `IN_PROGRESS`)
- Se completa la implementación (`IN_PROGRESS` → `MIGRATED`)
- Se valida el comportamiento (`MIGRATED` → `COMPLETED`)

**Estados válidos:** `NOT_STARTED` | `ANALYSIS` | `IN_PROGRESS` | `MIGRATED` | `VALIDATING` | `COMPLETED` | `BLOCKED` | `NOT_APPLICABLE` | `UNKNOWN`

---

## Módulos Ejecutables

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `InfoRest.vbp` / `InfoRest.exe` | VBP Project | `Inforest.Desktop` (POS module) | WinForms Module | IN_PROGRESS | modern-net8/src/Inforest.Desktop/POS/, Pedidos/, Ventas/, Caja/, Turno/ |
| `CajaRapida.vbp` / `CajaRapida.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Adicion.vbp` / `Adicion.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Administracion.vbp` / `Administracion.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Consulta.vbp` / `Consulta.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Despachador.vbp` / `Despachador.exe` | VBP Project | `DespachadorModule` | WinForms Module | IN_PROGRESS | modern-net8/src/Inforest.Desktop/Delivery/ |
| `Motorizados.vbp` / `Motorizado.exe` | VBP Project | `MotorizadoModule` | WinForms Module | IN_PROGRESS | modern-net8/src/Inforest.Desktop/Motorizado/ |

---

## Módulos BAS (Lógica de Negocio)

| Legacy | Tipo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `modDeclaracion.bas` | Global State | — | — | NOT_STARTED | 543 variables globales — requiere análisis especial |
| `modProcedimiento.bas` | Core Utils | — | — | NOT_STARTED | Actualizador, FPay, botones, QR |
| `modProcedimientoNuevo.bas` | Extensions | `ICashDroService` + `CashDroService` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/Hardware/ICashDroService.cs`, `modern-net8/src/Inforest.Infrastructure/Hardware/CashDroService.cs` |
| `modPuntoVenta.bas` | App Entry | `Program.cs` + `Form1` | WinForms + DI | COMPLETED | `modern-net8/src/Inforest.Desktop/Program.cs`, `modern-net8/src/Inforest.Desktop/Form1.cs` |
| `modCajaRapida.bas` | App Entry | — | — | NOT_STARTED | Sub Main caja |
| `modAdicion.bas` | App Entry | — | — | NOT_STARTED | Sub Main adición |
| `modAdministracion.bas` | App Entry | — | — | NOT_STARTED | Sub Main admin |
| `modConsulta.bas` | App Entry | — | — | NOT_STARTED | Sub Main consultas |
| `modConsultaIntregrada.bas` | Multi-local | — | — | NOT_STARTED | Consulta multi-local |
| `modDespachador.bas` | App Entry | `Program.cs` (DespachadorModule) + `DespachadorForm` | WinForms + DI | IN_PROGRESS | modern-net8/src/Inforest.Desktop/Delivery/ |
| `modMotorizado.bas` | App Entry | `Program.cs` (MotorizadoModule) + `LlegadaSalidaForm` | WinForms + DI | IN_PROGRESS | modern-net8/src/Inforest.Desktop/Motorizado/ |
| `modKDS.bas` | Integration | `IProduccionCocinaService` + `KdsXmlDispatcher` + `KdsLegacyGateway` + `MensajeCocinaRepository` | Interface + Classes | MIGRATED | `modern-net8/src/Inforest.Application/Interfaces/IProduccionCocinaService.cs`, `modern-net8/src/Inforest.Infrastructure/Kitchen/KdsXmlDispatcher.cs`, `modern-net8/src/Inforest.Infrastructure/Kitchen/KdsLegacyGateway.cs`, `modern-net8/src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs` |
| `modBlueVision.bas` | Integration | `IBlueVisionService` + `BlueVisionHttpClient` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/Hardware/IBlueVisionService.cs`, `modern-net8/src/Inforest.Infrastructure/Hardware/BlueVisionHttpClient.cs` |
| `modImpresoraFiscal.bas` | Integration | `IImpresoraFiscalService` + `NullImpresoraFiscalService` | Interface + Stub | BLOCKED | OCX 32-bit sin SDK .NET; ver `modern-net8/src/Inforest.Infrastructure/Hardware/NullImpresoraFiscalService.cs` |
| `modAuditoria.bas` | Audit | — | — | NOT_STARTED | Auditoría |
| `modAuditoriaEquipo.bas` | Audit | — | — | NOT_STARTED | Auditoría equipo |
| `modAuditoriaIntegral.bas` | Audit | `IAuditoriaService` + `AuditoriaService` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/IAuditoriaService.cs`, `modern-net8/src/Inforest.Infrastructure/Security/AuditoriaService.cs` |
| `modBarcode.bas` | Utility | — | — | NOT_STARTED | Códigos de barras |
| `modSeguridadInfhotel.bas` | License | `ILicenseService` + `LicenseService` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/ILicenseService.cs`, `modern-net8/src/Inforest.Infrastructure/Security/LicenseService.cs` |
| `modConexionIp.bas` | Network | — | — | NOT_STARTED | Conectividad |
| `modCrearInis.bas` | Config | — | — | NOT_STARTED | Creación INIs |
| `modTime.bas` | Utility | — | — | NOT_STARTED | Control de tiempo |
| `modMasticar.bas` | Unknown | — | — | UNKNOWN | Propósito no determinado |
| `modCheffControl.bas` | Kitchen | — | — | NOT_STARTED | Chef control |
| `modGuias.bas` | Dispatch | — | — | NOT_STARTED | Guías transporte |
| `modPvCorp.bas` | Unknown | — | — | UNKNOWN | Propósito no determinado |
| `CodigoControl.bas` | Fiscal | — | — | NOT_STARTED | Bolivia código control |
| `DLL3500.bas` | Hardware | `IPinPadService` + `PinPadService` | Interface + P/Invoke | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/Hardware/IPinPadService.cs`, `modern-net8/src/Inforest.Infrastructure/Hardware/PinPadService.cs` |
| `FpLibX_Const.bas` | Hardware | `IBiometriaService` + `NullBiometriaService` | Interface + Stub | BLOCKED | OCX 32-bit sin SDK .NET; ver `modern-net8/src/Inforest.Infrastructure/Hardware/NullBiometriaService.cs` |
| `ModPictureBoxCustom.bas` | UI Utility | — | — | NOT_STARTED | PictureBox custom |
| `ModuloHardKey.bas` | License | `ILicenseService` + `LicenseService` | Interface + Class | IN_PROGRESS | Validación contractual/SQL migrada; validación física del dongle pendiente |
| `VBZipBas.bas` | Utility | — | — | NOT_STARTED | ZIP |

---

## Clases (Domain Objects)

| Legacy | Tipo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `clsComando.cls` | Data Access | `IDbConnectionFactory` + `SqlConnectionFactory` | Interface + Class | COMPLETED | `src/Inforest.Application/Interfaces/IDbConnectionFactory.cs`, `src/Inforest.Infrastructure/Data/SqlConnectionFactory.cs` |
| `clsComando.cls` ExecSP/GetSP | Data Access | `ISpExecutor` + `SpExecutor` | Interface + Class | COMPLETED | `src/Inforest.Application/Interfaces/ISpExecutor.cs`, `src/Inforest.Infrastructure/Data/SpExecutor.cs` |
| `clsComando.cls` transactions | Data Access | `UnitOfWork` | Class | COMPLETED | `src/Inforest.Infrastructure/Data/UnitOfWork.cs` |
| `clsComando.cls` MsgBox err | Error handling | `SqlExceptionTranslator` | Class | COMPLETED | `src/Inforest.Infrastructure/Data/SqlExceptionTranslator.cs` |
| `ClsDocumento.cls` | Domain | — | — | NOT_STARTED | Documentos almacén |
| `clsAlmacen.cls` | Domain | `IInventoryGateway` + `InventoryGateway` | Interface + Class | MIGRATED | `modern-net8/src/Inforest.Application/Interfaces/IInventoryGateway.cs`, `modern-net8/src/Inforest.Infrastructure/Almacen/InventoryGateway.cs` |
| `ClsSeguridad.cls` | Security | `IAuthService` + `AuthService` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/IAuthService.cs`, `modern-net8/src/Inforest.Infrastructure/Security/AuthService.cs` |
| `clsDiaContable.cls` | Domain | — | — | NOT_STARTED | Día contable |
| `claCorreoElectronico.cls` | Notification | `INotificacionEmailService` + `SmtpEmailService` | Interface + Class | MIGRATED | `modern-net8/src/Inforest.Application/Interfaces/INotificacionEmailService.cs`, `modern-net8/src/Inforest.Infrastructure/Notifications/SmtpEmailService.cs` |
| `clsxml.cls` | Utility | — | — | NOT_STARTED | XML |
| `clsTrama.cls` | Integration | — | — | NOT_STARTED | FE Paperlees |
| `License.cls` | License | `ILicenseService` + `LicenseService` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/ILicenseService.cs`, `modern-net8/src/Inforest.Infrastructure/Security/LicenseService.cs` |
| `Mapping.cls` | Mapping | — | — | UNKNOWN | Propósito no determinado |

---

## Tablas de Base de Datos (Selección Principal)

| Legacy (SQL Server) | Tipo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `MPEDIDO` | Table | `Pedido`, `IPedidoRepository`, `PedidoRepository` | Entity + Repository | MIGRATED | modern-net8/src/Inforest.Domain/Entities/Ventas/Pedido.cs, modern-net8/src/Inforest.Infrastructure/Pedidos/ |
| `DPEDIDO` | Table | `DetallePedido`, `PedidoRepository` | Entity + Repository | MIGRATED | modern-net8/src/Inforest.Infrastructure/Pedidos/ |
| `CPEDIDO` | Table | — | — | NOT_STARTED | Combos en pedido |
| `APEDIDO` | Table | — | — | NOT_STARTED | Auditoría anulados |
| `MDOCUMENTO` | Table | `Documento`, `IDocumentoRepository`, `DocumentoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Ventas/ |
| `DDOCUMENTO` | Table | `DetalleDocumento` | Domain Entity | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Ventas/ |
| `MTURNO` | Table | `Turno`, `ITurnoRepository`, `TurnoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Infrastructure/Turno/ |
| `TCAJA` | Table | `ConfiguracionCaja`, `IParametroRepository.ObtenerConfiguracionCajaAsync`, `ValidarInicioPosHandler` | Domain Record + Repository + Handler | COMPLETED | `modern-net8/src/Inforest.Domain/Entities/Configuracion/ConfiguracionCaja.cs`, `modern-net8/src/Inforest.Application/Configuracion/ConfiguracionHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Configuracion/ParametroRepository.cs`, `modern-net8/tests/Inforest.Application.Tests/Configuracion/ValidarInicioPosHandlerTests.cs`, `modern-net8/tests/Inforest.Infrastructure.Tests/Configuracion/ConfiguracionCajaTests.cs` |
| `TCAJATERMINAL` | Table | `IPinPadTerminalRepository`, `PinPadTerminalRepository`, `ObtenerTerminalesPinPadHandler` | Repository + Handler | COMPLETED | `modern-net8/src/Inforest.Infrastructure/Hardware/PinPadTerminalRepository.cs` + `modern-net8/src/Inforest.Application/Caja/PinPadHandlers.cs` |
| `TPRODUCTO` | Table | `ProductoMaestro`, `IProductoMaestroRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TGRUPO` | Table | `GrupoProducto`, `IGrupoProductoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TSUBGRUPO` | Table | `SubGrupoProducto`, `ISubGrupoProductoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TCLIENTE` | Table | `Cliente`, `IClienteRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TTERMINAL` | Table | `TerminalPinPad`, `IPinPadTerminalRepository`, `PinPadTerminalRepository` | Domain Record + Repository | COMPLETED | `modern-net8/src/Inforest.Domain/Entities/Hardware/TerminalPinPad.cs` + `modern-net8/src/Inforest.Infrastructure/Hardware/PinPadTerminalRepository.cs` |
| `TUSUARIO` | Table | — | — | NOT_STARTED | Usuarios |
| `TPARAMETRO` | Table | `ConfiguracionSistema`, `IParametroRepository`, `ParametroService`, `ValidarInicioPosHandler` | Domain Record + Repository + Service + Handler | COMPLETED | `modern-net8/src/Inforest.Domain/Entities/Configuracion/ConfiguracionSistema.cs`, `modern-net8/src/Inforest.Application/Configuracion/ConfiguracionHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Configuracion/ParametroRepository.cs`, `modern-net8/src/Inforest.Infrastructure/Configuracion/ParametroService.cs`, `modern-net8/tests/Inforest.Application.Tests/Configuracion/ValidarInicioPosHandlerTests.cs`, `modern-net8/tests/Inforest.Infrastructure.Tests/Configuracion/ParametroRepositoryTests.cs` |
| `DPAGODOCUMENTO` | Table | `PagoDocumento`, `IPagoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Caja/ |
| `MCIERRE` | Table | — | — | NOT_STARTED | Cierre contable |
| `TDIACONTABLE` | Table | `DiaContable`, `IDiaContableService`, `DiaContableService` | Entity + Service | IN_PROGRESS | modern-net8/src/Inforest.Infrastructure/Turno/ |

> Ver inventario completo: [docs/database/legacy-database.md](../database/legacy-database.md)

---

## Stored Procedures Clave

| Legacy SP | Módulo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `spIns_MPEDIDO` | Pedidos | — | — | NOT_STARTED | Inserción pedido |
| `spUpd_MPEDIDO` | Pedidos | — | — | NOT_STARTED | Actualización pedido |
| `spRep_RegVenta` | Reportes | — | — | NOT_STARTED | Reporte ventas |
| `spRep_Anulacion` | Reportes | — | — | NOT_STARTED | Reporte anulaciones |
| `USP_KDS_ObtenerProductoPedido` | KDS | `KdsLegacyGateway` | Class | IN_PROGRESS | `modern-net8/src/Inforest.Infrastructure/Kitchen/KdsLegacyGateway.cs` |
| `USP_MODIFICARINSUMOS` | Almacén | — | — | NOT_STARTED | Gestión insumos |
| `usp_Aud_ObtenerSalones` | Auditoría | — | — | NOT_STARTED | Auditoría salones |

> Ver inventario completo: [docs/database/stored-procedures.md](../database/stored-procedures.md)

---

## Reportes Crystal Reports

> 206 reportes en total. Ver inventario en [docs/modules/restaurante/README.md](../modules/restaurante/README.md)

| Legacy Reporte | Tipo | .NET 8 | Tipo | Estado |
|---|---|---|---|---|
| `CrtRepInventarioIntegrado.dsr` | Crystal | — | — | NOT_STARTED |
| `CrtRegComprasIntegrado.dsr` | Crystal | — | — | NOT_STARTED |
| `dsrPedidos*.dsr` | Crystal | — | — | NOT_STARTED |
| `dsrPaloteoVenta*.dsr` | Crystal | — | — | NOT_STARTED |
| *(206 en total)* | Crystal | — | — | NOT_STARTED |

---

## Integraciones Externas

| Legacy | Tecnología | .NET 8 | Tipo | Estado |
|---|---|---|---|---|
| KDS (Kitchen Display) | XML sobre directorio | `KdsXmlDispatcher` | Class | IN_PROGRESS |
| BlueVision/TVS | COM ActiveX | `BlueVisionHttpClient` | Class | IN_PROGRESS |
| Facturación Electrónica | Múltiples SDKs | `FacturacionElectronicaFactory` + gateways por país | Factory + Stubs | IN_PROGRESS |
| PinPad DLL3500 | Win32 DLL | `PinPadService` | P/Invoke | IN_PROGRESS |
| CashDro | HTTP API | `CashDroService` | Process adapter | IN_PROGRESS |
| Rappi | SQL SP | `IRappiOrderService` + `RappiOrderAdapter` | Interface + Stub | IN_PROGRESS |
| FPay/MercadoPago QR | SP + motor | — | — | NOT_STARTED |
| Biometría SecuGen | OCX Win32 | `IBiometriaService` + `NullBiometriaService` | Interface + Stub | BLOCKED |
| Impresora Fiscal Epson | OCX Win32 | `IImpresoraFiscalService` + `NullImpresoraFiscalService` | Interface + Stub | BLOCKED |
| Chilkat Email | COM | — | — | NOT_STARTED |

---

*Esta matriz debe actualizarse con cada avance de migración.*

---

## Componentes de Infraestructura Transversal (P3-01 — Baseline)

> Etapa P3-01 — Re-baseline de arquitectura ejecutable y gobierno técnico. Completada 2026-08-11.

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| *(transversal)* | — | `Inforest.Domain.Common.Result` | Class | COMPLETED | `src/Inforest.Domain/Common/Result.cs` |
| *(transversal)* | — | `Inforest.Domain.Common.Result<T>` | Class | COMPLETED | `src/Inforest.Domain/Common/Result.cs` |
| `clsComando.cls` On Error | Error handling | `Inforest.Infrastructure.Exceptions.InfrastructureException` | Class | COMPLETED | `src/Inforest.Infrastructure/Exceptions/InfrastructureException.cs` |

---

## Componentes de Seguridad (P3-03 — Inicio)

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmAcceso.frm` | Form | `src/Inforest.Desktop/Form1.cs` | WinForms Form | IN_PROGRESS | `modern-net8/src/Inforest.Desktop/Form1.cs`, `modern-net8/src/Inforest.Desktop/Form1.Designer.cs` |
| `frmAcceso.frm` / `USUARIO.INI` | Session bootstrap | `ISessionService` + `SessionService` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/ISessionService.cs`, `modern-net8/src/Inforest.Infrastructure/Security/SessionService.cs` |
| `frmDiaContable.frm` / `clsDiaContable.cls` (POS-FUNC-009) | Form + Class | `FrmDiaContable` + `AperturarDiaContableHandler` + `CerrarDiaContableHandler` + `ObtenerDiaContableHandler` + `IDiaContableService` (extendido: `CerrarDiaContableAsync`, `ObtenerFechaMaximaDiaContableAsync`) | WinForm + Handlers + Interface | MIGRATED | `src/Inforest.Desktop/Turno/FrmDiaContable.cs`, `src/Inforest.Application/Turno/DiaContableHandlers.cs`, `src/Inforest.Infrastructure/Turno/DiaContableService.cs` — Tests: `DiaContableHandlerTests` (10) |
| `frmPedido.frm` (POS-FUNC-003) | Form | `FrmPedido` + `Pedido` + `DetallePedido` + `CreatePedidoHandler` + `UpdatePedidoHandler` + `ObtenerPedidoHandlers` + `PedidoRepository` | WinForm + Domain + Handlers + Repository | MIGRATED | `modern-net8/src/Inforest.Desktop/Pedidos/FrmPedido.cs`, `modern-net8/src/Inforest.Domain/Entities/Ventas/Pedido.cs`, `modern-net8/src/Inforest.Application/Pedidos/`, `modern-net8/src/Inforest.Infrastructure/Pedidos/PedidoRepository.cs` — Tests: `CreatePedidoHandlerTests` + `UpdatePedidoHandlerTests` + `ObtenerPedidoHandlerTests` + `PedidoTests` + `PedidoExtendidoTests` |
| `TACCESO` / `TGRUPOACCESO` | RBAC | `IRbacService` + `RbacService` | Interface + Class | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/IRbacService.cs`, `modern-net8/src/Inforest.Infrastructure/Security/RbacService.cs` |
| `clsComando.cls` ADODB errors | DB errors | `Inforest.Infrastructure.Exceptions.DatabaseException` | Class | COMPLETED | `src/Inforest.Infrastructure/Exceptions/InfrastructureException.cs` |
| `Sub Main()` modPuntoVenta.bas | Startup | `Inforest.Desktop.Program` (DI-corrected bootstrap) | Class | COMPLETED | `src/Inforest.Desktop/Program.cs` |
| `modPuntoVenta.bas` (`App.PrevInstance`) + `mdiPuntoVenta.frm` | Startup rule + Form | `IAppInstanceGuard` + `AppInstanceGuard` + `Program.cs` + `FrmPuntoVenta` | Interface + Class + WinForm | MIGRATED | `modern-net8/src/Inforest.Application/Interfaces/IAppInstanceGuard.cs`, `modern-net8/src/Inforest.Infrastructure/Runtime/AppInstanceGuard.cs`, `modern-net8/src/Inforest.Desktop/Program.cs`, `modern-net8/src/Inforest.Desktop/POS/FrmPuntoVenta.cs` |
| *(suite Fase 3 domain)* | — | `Inforest.Domain.Tests` (51 tests) | xUnit | COMPLETED | `tests/Inforest.Domain.Tests/` |
| *(suite Fase 3 application)* | — | `Inforest.Application.Tests` (26 tests) | xUnit | COMPLETED | `tests/Inforest.Application.Tests/` |
| *(suite Fase 3 infrastructure)* | — | `Inforest.Infrastructure.Tests` (72 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/` |

### Etapa 9 — Delivery, Despacho, Motorizados

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmNuevoDelivery.frm` | Form | `NuevoDeliveryForm` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Delivery/NuevoDeliveryForm.cs` |
| `frmDespachador.frm` | Form | `DespachadorForm` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Delivery/DespachadorForm.cs` |
| `frmCentralPedidos.frm` | Form | `CentralPedidosForm` (extendido: ConfirmarEntrega + RevertirEntrega + ModificarFecha) | WinForm | MIGRATED | `src/Inforest.Desktop/Delivery/CentralPedidosForm.cs` |
| `frmPedidoDelivery.frm` | Form | `FrmPedidoDelivery` | WinForm | MIGRATED | `src/Inforest.Desktop/Delivery/FrmPedidoDelivery.cs` |
| `frmLlegadaSalida.frm` | Form | `LlegadaSalidaForm` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Motorizado/LlegadaSalidaForm.cs` |
| `frmAsignacionMotorizado.frm` | Form | `AsignacionMotorizadoForm` | WinForm | MIGRATED | `src/Inforest.Desktop/Motorizado/AsignacionMotorizadoForm.cs` |
| `frmReasignacionMotorizado.frm` | Form | `ReasignacionMotorizadoForm` | WinForm | MIGRATED | `src/Inforest.Desktop/Motorizado/ReasignacionMotorizadoForm.cs` |
| `frmTarifaMotorizado.frm` | Form | `TarifaMotorizadoForm` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Motorizado/TarifaMotorizadoForm.cs` |
| `TDELIVERY` | Table | `ClienteDelivery` | Entity | IN_PROGRESS | `src/Inforest.Domain/Entities/Delivery/ClienteDelivery.cs` |
| `MPEDIDO` (delivery fields) | Table | `PedidoDelivery` | Entity | IN_PROGRESS | `src/Inforest.Domain/Entities/Delivery/PedidoDelivery.cs` |
| `TLOCAL` | Table | `LocalRestaurante` | Entity | IN_PROGRESS | `src/Inforest.Domain/Entities/Delivery/LocalRestaurante.cs` |
| `TTIPOPEDIDODETALLE` | Table | `TipoPedidoDelivery` | Entity | IN_PROGRESS | `src/Inforest.Domain/Entities/Delivery/TipoPedidoDelivery.cs` |
| `TMOTORIZADODATOS` | Table | `Motorizado` | Entity | IN_PROGRESS | `src/Inforest.Domain/Entities/Motorizado/Motorizado.cs` |
| `(asignacion pedido-motorizado)` | Logic | `AsignacionMotorizado` | Entity | IN_PROGRESS | `src/Inforest.Domain/Entities/Motorizado/AsignacionMotorizado.cs` |
| `IClienteDeliveryRepository` | Interface | `IClienteDeliveryRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/IClienteDeliveryRepository.cs` |
| `IPedidoDeliveryRepository` | Interface | `IPedidoDeliveryRepository` (extendido: ConfirmarEntregaAsync + RevertirEntregaAsync + ModificarFechaProgramadaAsync + ObtenerEstadoPagoAsync + EstaEntregadoAsync) | Interface | MIGRATED | `src/Inforest.Domain/Repositories/IPedidoDeliveryRepository.cs` |
| `IMotorizadoRepository` | Interface | `IMotorizadoRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/IMotorizadoRepository.cs` |
| `ILocalRepository` | Interface | `ILocalRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/ILocalRepository.cs` |
| `ICentralPedidosRepository` | Interface | `ICentralPedidosRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/ICentralPedidosRepository.cs` |
| `sp_Inforest_PedidosCentralPedido` | SP | `CentralPedidosRepository.ObtenerPedidosCentralAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Delivery/CentralPedidosRepository.cs` |
| `sp_CD_Modificar_EstadoDelivery_Cabecera` | SP | `CentralPedidosRepository.ModificarEstadoDeliveryAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Delivery/CentralPedidosRepository.cs` |
| `sp_UpdFotoDelivery` | SP | `ActualizarFotoClienteDeliveryHandler` + `ClienteDeliveryRepository.ActualizarFotoAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Delivery/DeliveryHandlers.cs`, `src/Inforest.Infrastructure/Delivery/ClienteDeliveryRepository.cs` |
| `MPEDIDO.lEntregado / tusuarioentregado / fregentregado` | Field | `ConfirmarEntregaCentralHandler` + `RevertirEntregaCentralHandler` + `PedidoDeliveryRepository.ConfirmarEntregaAsync/RevertirEntregaAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Delivery/DeliveryHandlers.cs`, `src/Inforest.Infrastructure/Delivery/PedidoDeliveryRepository.cs` |
| `MPEDIDO.fProgramacion / fregistro` (Central) | Field | `ModificarFechaProgramadaDeliveryHandler` + `PedidoDeliveryRepository.ModificarFechaProgramadaAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Delivery/DeliveryHandlers.cs`, `src/Inforest.Infrastructure/Delivery/PedidoDeliveryRepository.cs` |
| `vDespachador` (seguimiento) | View | `ObtenerPedidosSeguimientoDeliveryHandler` + `FrmPedidoDelivery` | Handler + WinForm | MIGRATED | `src/Inforest.Application/Delivery/DeliveryHandlers.cs`, `src/Inforest.Desktop/Delivery/FrmPedidoDelivery.cs` |
| `spRep_AnaliticoMotorizadoIntegrado` | SP | *(reporte FastReport pendiente)* | Report | NOT_STARTED | — |
| `frmOrdenesConsola.frm` | Form | `ObtenerOrdenesExternasHandler` + `RappiOrderAdapter` | Handler+Service | IN_PROGRESS | `src/Inforest.Application/Delivery/RappiHandlers.cs` |
| `(test etapa 9 domain)` | — | `DeliveryDomainTests` (12 tests) + `MotorizadoDomainTests` (11 tests) | xUnit | IN_PROGRESS | `tests/Inforest.Domain.Tests/Delivery/` |
| `(test etapa 9 application)` | — | `DeliveryHandlersTests` (4 tests) + `MotorizadoHandlersTests` (5 tests) + `CentralPedidosHandlersTests` (10 tests) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/` |

| **Etapa 10 — Reportes, Consultas y Salidas Operativas** | | | | | |
| `spRep_Comanda` | SP | `ReporteRepository.ObtenerComandaAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrComandaD.Dsr` / `dsrComandaR.Dsr` | Crystal | `RepComanda.frx` (resumido) / `RepComandaDetallado.frx` (detallado) | FastReport | COMPLETED | `src/Inforest.Desktop/Reports/Templates/RepComanda.frx`, `src/Inforest.Desktop/Reports/Templates/RepComandaDetallado.frx` |
| `frmRepComanda.frm` | Form | `FrmComandaReporte.cs` + `ObtenerReporteComandaHandler` + `ComandaRow` + `ObtenerComandaAsync` + BR-REP-001 + 2 tests | WinForm + Handler + DTO + Repository | COMPLETED | `src/Inforest.Desktop/Forms/Reportes/FrmComandaReporte.cs`, `src/Inforest.Application/Reportes/` |
| `spRep_Propina` | SP | `ReporteRepository.ObtenerPropinaAsync` | Repository | COMPLETED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrPropinaD.Dsr` / `dsrPropinaR.Dsr` | Crystal | `RepPropina.frx` (Detallado, landscape, 10 col) / `RepPropinaResumido.frx` (Resumido, portrait, 5 col) | FastReport | COMPLETED | `src/Inforest.Desktop/Reports/Templates/RepPropina.frx`, `src/Inforest.Desktop/Reports/Templates/RepPropinaResumido.frx` |
| `frmRepPropina.frm` | Form | `FrmPropinaReporte.cs` + `ObtenerReportePropinaHandler` + `PropinaRow` + `ObtenerPropinaAsync` + BR-REP-002 + 2 tests | WinForm + Handler + DTO + Repository | COMPLETED | `src/Inforest.Desktop/Forms/Reportes/FrmPropinaReporte.cs`, `src/Inforest.Application/Reportes/` |
| `spRep_PrincipalCliente` | SP | `ReporteRepository.ObtenerPrincipalCliente*Async` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrPrincipalD.Dsr` / `dsrPrincipalR.Dsr` | Crystal | `RepPrincipalClienteDetalle.frx` / `ResumenRepPrincipalCliente.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `spRep_CtaCteN` | SP | `ReporteRepository.ObtenerCtaCteOperativaAsync` | Repository | COMPLETED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepCtaCte.frm` | Form | `FrmRepCtaCteReporte.cs` | WinForm | COMPLETED | `src/Inforest.Desktop/Forms/Reportes/FrmRepCtaCteReporte.cs` |
| `frmRepCtaCte.frx` | Form resource | `RepCtaCteConsolidado.frx` / `RepCtaCteResumido.frx` / `RepCtaCteDetallado.frx` | FastReport | COMPLETED | `src/Inforest.Desktop/Reports/Templates/` |
| `spRep_Anulacion` | SP | `ReporteRepository.ObtenerAnulacionAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepAnulado.frm` | Form | `FrmRepAnuladoReporte.cs` | WinForm | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepAnuladoReporte.cs` |
| `frmRepAnulado.frx` | Form resource | `RepAnulacion.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/Reports/Templates/` |
| `spRep_LiquidacionSuma` | SP | `ReporteRepository.ObtenerLiquidacionTicketAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepLiquidacionTicket.frm` | Form | `FrmRepLiquidacionTicketReporte.cs` | WinForm | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepLiquidacionTicketReporte.cs` |
| `frmRepLiquidacionTicket.frx` | Form resource | `RepLiquidacionTicket.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepLiquidacionTicket.frx` |
| `frmRepPaloteoTicket.frm` | Form | `FrmRepPaloteoTicketReporte.cs` | WinForm | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepPaloteoTicketReporte.cs` |
| `(query dinámica frmRepPaloteoTicket)` | Query | `ReporteRepository.ObtenerPaloteoTicketAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepPaloteoTicket.frx` | Form resource | `RepPaloteoTicket.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepPaloteoTicket.frx` |
| `frmRepDeliveryTicket.frm` | Form | `FrmRepDeliveryTicketReporte.cs` | WinForm | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepDeliveryTicketReporte.cs` |
| `(query dinámica frmRepDeliveryTicket)` | Query | `ReporteRepository.ObtenerDeliveryTicketAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepDeliveryTicket.frx` | Form resource | `RepDeliveryTicket.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepDeliveryTicket.frx` |
| `frmRepReservas.frm` | Form | `FrmRepReservasReporte.cs` | WinForm | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepReservasReporte.cs` |
| `(query dinámica frmRepReservas: TRESERVA+vEstadoReserva)` | Query | `ReporteRepository.ObtenerReservasReporteAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepReservas.frx` | Form resource | `RepReservas.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepReservas.frx` |
| `spRep_Entregas` | SP | `ReporteRepository.ObtenerEntregasAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepEntrega.frm` | Form | `FrmRepEntregaReporte.cs` | WinForm | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepEntregaReporte.cs` |
| `dsrReporteEntrega.Dsr` | Crystal | `RepEntregaFormato1.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepEntregaFormato1.frx` |
| `dsrReporteEntregaFormato2.Dsr` | Crystal | `RepEntregaFormato2.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepEntregaFormato2.frx` |
| `dsrReporteEntregaResumidoProd.Dsr` | Crystal | `RepEntregaResumidoProd.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepEntregaResumidoProd.frx` |
| `spRep_VentaFecha` | SP | `ReporteRepository.ObtenerVentaFechaAsync` + `ReporteRepository.ObtenerSubGruposAsync` | Repository | COMPLETED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepVentaFecha.frm` | Form | `FrmRepVentaFechaReporte.cs` + `ObtenerReporteVentaFechaHandler` + `VentaFechaRow` + `VentaFechaParametros` + `SubGrupoItem` + BR-REP-020 + 2 tests | WinForm + Handler + DTO + Repository | COMPLETED | `src/Inforest.Desktop/Forms/Reportes/FrmRepVentaFechaReporte.cs`, `src/Inforest.Application/Reportes/` |
| `frmRepVentaFecha.frx` | Form resource | `RepVentaFecha.frx` | FastReport | COMPLETED | `src/Inforest.Desktop/Reports/Templates/RepVentaFecha.frx` |
| `spRep_LiquidacionOutPut` | SP | `ReporteRepository.ObtenerLiquidacionOutputAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `spRep_Liquidacion` | SP | `ReporteRepository.ObtenerLiquidacionDocumentosAsync` + `ObtenerLiquidacionSumasGrupoAsync` + `ObtenerLiquidacionTarjetasAsync` + `ObtenerLiquidacionOtrosTiposAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `spRep_LiquidacionSuma` | SP | `ReporteRepository.ObtenerLiquidacionTiposPedidoAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `frmRepLiquidacion.frm` | Form | `FrmRepLiquidacionReporte.cs` + `ObtenerReporteLiquidacionHandler` + `LiquidacionParametros` + `LiquidacionOutput` + `LiquidacionRow` + `LiquidacionSumaGrupoRow` + `LiquidacionTarjetaRow` + `LiquidacionTipoPedidoRow` + `LiquidacionOtroTipoRow` + `LiquidacionResultado` + BR-REP-021 + 2 tests | WinForm + Handler + DTOs + Repository | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepLiquidacionReporte.cs`, `src/Inforest.Application/Reportes/` |
| `dsrLiquidacion.Dsr` / `dsrLiquidacionVenta.Dsr` | Crystal | `RepLiquidacion.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepLiquidacion.frx` |
| `spRep_RegVenta` | SP | `ReporteRepository.ObtenerRegistroVentaAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `spRep_RegVentaSunat` | SP | `ReporteRepository.ObtenerRegistroVentaSunatAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `spRep_RegVentaSunatAD` | SP | `ReporteRepository.ObtenerRegistroVentaSunatAdAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `spRep_ComprobanteDetallado` | SP | `ReporteRepository.ObtenerRegistroVentaDetalladoAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `spRep_RegVentaSunat_formaPago` | SP | — | — | NOT_APPLICABLE (GAP-REP-022) | SP no encontrado en `5. SP.sql` |
| `frmRepRegistroVenta.frm` | Form | `FrmRepRegistroVentaReporte.cs` + `ObtenerReporteRegistroVentaHandler` + `RegistroVentaParametros` + `RegistroVentaRow` + `RegistroVentaSunatRow` + `RegistroVentaSunatAdRow` + `RegistroVentaDetalladoRow` + `RegistroVentaResultado` + BR-REP-022 + 2 tests | WinForm + Handler + DTOs + Repository | MIGRATED | `src/Inforest.Desktop/Forms/Reportes/FrmRepRegistroVentaReporte.cs`, `src/Inforest.Application/Reportes/` |
| `RepRegistroVentaSunat.frx` | FastReport | `RepRegistroVentaSunat.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepRegistroVentaSunat.frx` |
| `RepRegistroVentaDetallado.frx` | FastReport | `RepRegistroVentaDetallado.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepRegistroVentaDetallado.frx` |
| `RepRegistroVentaConsolidado.frx` | FastReport | `RepRegistroVentaConsolidado.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepRegistroVentaConsolidado.frx` |
| `RepRegistroVentaSunatAd.frx` | FastReport | `RepRegistroVentaSunatAd.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepRegistroVentaSunatAd.frx` |
| `RepRegistroVentaComprobante.frx` | FastReport | `RepRegistroVentaComprobante.frx` | FastReport | MIGRATED | `src/Inforest.Desktop/Reports/Templates/RepRegistroVentaComprobante.frx` |
| `spRep_CtaCteIntegrado` | SP | `ReporteRepository.ObtenerCtaCteIntegradoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrRepCtaCteIntegradoConsolidado.Dsr` | Crystal | `RepCtaCteIntegradoConsolidado.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `frmRepCtaCteIntegrado.frm` | Form | `FrmCtaCteIntegradoReporte.cs` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Forms/Reportes/FrmCtaCteIntegradoReporte.cs` |
| `spRep_PaloteoComparativo` | SP | `ReporteRepository.ObtenerPaloteoComparativoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrPaloteoComparativo.Dsr` | Crystal | `RepPaloteoComparativo.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `frmRepPaloteoComparativo.frm` | Form | `FrmPaloteoComparativoReporte.cs` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Forms/Reportes/FrmPaloteoComparativoReporte.cs` |
| `spRep_PaloteoSubProd` | SP | `ReporteRepository.ObtenerPaloteoSubProdAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrPaloteoSPD.Dsr` | Crystal | `RepPaloteoSubProd.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `spRep_PaloteoVentaIntegrado` | SP | `ReporteRepository.ObtenerPaloteoVentaIntegradoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrPaloteoVentaIntegradol.Dsr` | Crystal | `RepPaloteoVentaIntegradoDetallado.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `spRep_RankingIntegrado` | SP | `ReporteRepository.ObtenerRankingIntegradoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrRankingIntegradol.Dsr` | Crystal | `RepRankingIntegradoDetallado.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `spRep_VentaMensualIntegrado` | SP | `ReporteRepository.ObtenerVentaMensualIntegradoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrVentaMensualIntegrada.Dsr` | Crystal | `RepVentaMensualIntegrado.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `frmRepVentaMensualIntegrado.frm` | Form | `FrmVentaMensualIntegradoReporte.cs` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Forms/Reportes/FrmVentaMensualIntegradoReporte.cs` |
| `spRep_AnaliticoMotorizadoIntegrado` | SP | `ReporteRepository.ObtenerAnaliticoMotorizadoIntegradoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrAnaliticoMotorizadoIntegradoConsolidado.Dsr` | Crystal | `RepAnaliticoMotorizadoIntegrado.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `frmRepAnaliticoMotorizadoIntegrado.frm` | Form | `FrmAnaliticoMotorizadoIntegradoReporte.cs` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Forms/Reportes/FrmAnaliticoMotorizadoIntegradoReporte.cs` |
| `USP_KDS_ResporteTiempoPedido` | SP | `ReporteRepository.ObtenerTiempoKdsPedidoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `USP_KDS_ResporteTiempoProducto` | SP | `ReporteRepository.ObtenerTiempoKdsProductoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrTiempoKdsP.Dsr` / `dsrTiempoKdsR.Dsr` | Crystal | `RepTiempoKdsPedido.frx` / `RepTiempoKdsProducto.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `(KDS tiempo form)` | Form | `FrmTiempoKdsReporte.cs` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Forms/Reportes/FrmTiempoKdsReporte.cs` |
| `IReporteRepository` | Interface | `IReporteRepository` | Interface | COMPLETED | `src/Inforest.Application/Reportes/IReporteRepository.cs` |
| `ReporteRepository` | — | `ReporteRepository` | Repository | COMPLETED | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `(test etapa 10 application)` | — | `ReportesHandlersTests` (18 tests) | xUnit | COMPLETED | `tests/Inforest.Application.Tests/Reportes/` |

## P3-11 — Periféricos Especializados e Integraciones por País

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `DLL3500.bas` | Module | `IPinPadService` | Interface | IN_PROGRESS | `src/Inforest.Application/Interfaces/Hardware/IPinPadService.cs` |
| `caja_pinpad.dll` | DLL Win32 | `PinPadService` | P/Invoke | IN_PROGRESS | `src/Inforest.Infrastructure/Hardware/PinPadService.cs` |
| — | — | `NullPinPadService` | Stub | COMPLETED | `src/Inforest.Infrastructure/Hardware/NullPinPadService.cs` |
| `modProcedimientoNuevo.bas` (CashDro) | Module | `ICashDroService` | Interface | IN_PROGRESS | `src/Inforest.Application/Interfaces/Hardware/ICashDroService.cs` |
| `MotorCashDrow.exe` | Process | `CashDroService` | Process.Start | IN_PROGRESS | `src/Inforest.Infrastructure/Hardware/CashDroService.cs` |
| — | — | `NullCashDroService` | Stub | COMPLETED | `src/Inforest.Infrastructure/Hardware/NullCashDroService.cs` |
| `modBlueVision.bas` | Module | `IBlueVisionService` | Interface | IN_PROGRESS | `src/Inforest.Application/Interfaces/Hardware/IBlueVisionService.cs` |
| `BlueVision_Core_TVS.dll` (COM) | COM DLL | `BlueVisionHttpClient` | HttpClient | IN_PROGRESS | `src/Inforest.Infrastructure/Hardware/BlueVisionHttpClient.cs` |
| — | — | `NullBlueVisionService` | Stub | COMPLETED | `src/Inforest.Infrastructure/Hardware/NullBlueVisionService.cs` |
| `FpLibX_Const.bas`, `sgfplibx.ocx` | OCX 32-bit | `IBiometriaService` | Interface | BLOCKED | `src/Inforest.Application/Interfaces/Hardware/IBiometriaService.cs` |
| — | — | `NullBiometriaService` | Stub (GAP) | COMPLETED | `src/Inforest.Infrastructure/Hardware/NullBiometriaService.cs` |
| `modImpresoraFiscal.bas`, `IFEpson.ocx` | OCX 32-bit | `IImpresoraFiscalService` | Interface | BLOCKED | `src/Inforest.Application/Interfaces/Hardware/IImpresoraFiscalService.cs` |
| — | — | `NullImpresoraFiscalService` | Stub (GAP) | COMPLETED | `src/Inforest.Infrastructure/Hardware/NullImpresoraFiscalService.cs` |
| `clsTrama.cls`, `clsxml.cls`, `lFacturacionE` | Class/Vars | `IFacturacionElectronicaGateway` | Interface | IN_PROGRESS | `src/Inforest.Application/Interfaces/Hardware/IFacturacionElectronicaGateway.cs` |
| — | — | `FacturacionElectronicaFactory` | Factory | COMPLETED | `src/Inforest.Infrastructure/Country/FacturacionElectronicaFactory.cs` |
| — (SUNAT/OSE Perú) | — | `PeruFEGateway` | Gateway stub | IN_PROGRESS | `src/Inforest.Infrastructure/Country/PeruFEGateway.cs` |
| — (AFIP Argentina) | — | `ArgentinaFEGateway` | Gateway stub | IN_PROGRESS | `src/Inforest.Infrastructure/Country/ArgentinaFEGateway.cs` |
| — (SII Chile) | — | `ChileFEGateway` | Gateway stub | IN_PROGRESS | `src/Inforest.Infrastructure/Country/FEGateways.cs` |
| — (SRI Ecuador) | — | `EcuadorFEGateway` | Gateway stub | IN_PROGRESS | `src/Inforest.Infrastructure/Country/FEGateways.cs` |
| — (SIAT Bolivia) | — | `BoliviaFEGateway` | Gateway stub | IN_PROGRESS | `src/Inforest.Infrastructure/Country/FEGateways.cs` |
| — | — | `NullFEGateway` | Stub | COMPLETED | `src/Inforest.Infrastructure/Country/NullFEGateway.cs` |
| `pais` global + scripts opcionales | Variable/SQL | `IPaisPolicy` | Interface | IN_PROGRESS | `src/Inforest.Application/Interfaces/Country/IPaisPolicy.cs` |
| — | — | `PeruPaisPolicy` | Policy | COMPLETED | `src/Inforest.Infrastructure/Country/PaisPolicies.cs` |
| — | — | `ArgentinaPaisPolicy` | Policy | COMPLETED | `src/Inforest.Infrastructure/Country/PaisPolicies.cs` |
| — | — | `ChilePaisPolicy` | Policy | COMPLETED | `src/Inforest.Infrastructure/Country/PaisPolicies.cs` |
| — | — | `EcuadorPaisPolicy` | Policy | COMPLETED | `src/Inforest.Infrastructure/Country/PaisPolicies.cs` |
| — | — | `BoliviaPaisPolicy` | Policy | COMPLETED | `src/Inforest.Infrastructure/Country/PaisPolicies.cs` |
| — | — | `EspanaPaisPolicy` | Policy | COMPLETED | `src/Inforest.Infrastructure/Country/PaisPolicies.cs` |
| — | — | `PaisPolicyFactory` | Factory | COMPLETED | `src/Inforest.Infrastructure/Country/PaisPolicies.cs` |
| `modDeclaracion.bas` (país/FE) | Domain Entity | `ConfiguracionPais` | Domain Entity | COMPLETED | `src/Inforest.Domain/Entities/Hardware/ConfiguracionPais.cs` |
| `clsTrama.cls` / `clsxml.cls` | Domain | `DocumentoFE` | Domain Entity | COMPLETED | `src/Inforest.Domain/Entities/Hardware/DocumentoFE.cs` |
| `modBlueVision.bas` (TvsTicket) | Domain | `BlueVisionTicket` / `BlueVisionLinea` | Domain Entity | COMPLETED | `src/Inforest.Domain/Entities/Hardware/BlueVisionTicket.cs` |
| `modImpresoraFiscal.bas` (zonas) | Domain | `FiscalDocumento` / `FiscalItem` | Domain Entity | COMPLETED | `src/Inforest.Domain/Entities/Hardware/FiscalDocumento.cs` |
| `frmPago.frm`, `frmPagoPinPad.frm` | Form | `FrmPago` / `FrmPagoPinPad` + `ProcesarPagoPinPadHandler` + `ObtenerTerminalesPinPadHandler` | WinForm + Handlers | COMPLETED | `src/Inforest.Desktop/Caja/FrmPago.cs` + `src/Inforest.Desktop/Caja/FrmPagoPinPad.cs` + `src/Inforest.Application/Caja/PinPadHandlers.cs` |
| (tests P3-11) | — | `NullHardwareServicesTests` (15 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| (tests P3-11) | — | `FacturacionElectronicaFactoryTests` (7 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| (tests P3-11) | — | `PaisPolicyTests` (8 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| (tests P3-11) | — | `DomainHardwareEntityTests` (8 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| `frmDocumentoCorrelativo.frm` + `modPuntoVenta.bas` | Form/Module | `CorrelativoDocumento` | Domain Entity | COMPLETED | `src/Inforest.Domain/Entities/Ventas/CorrelativoDocumento.cs` |
| `TTIPODOCUMENTOIMPRESORA` | Table | `ICorrelativoRepository` / `CorrelativoRepository` | Repository | COMPLETED | `src/Inforest.Infrastructure/Ventas/CorrelativoRepository.cs` |
| `frmDocumentoCorrelativo.frm` | Form | `FrmDocumentoCorrelativo` | WinForm | COMPLETED | `src/Inforest.Desktop/Ventas/FrmDocumentoCorrelativo.cs` |
| `frmPrecuentaImpresora.frm` | Form | `ImprimirPrecuentaHandler` / `FrmPrecuentaImpresora` | Handler + WinForm | COMPLETED | `src/Inforest.Application/Impresion/ImprimirPrecuentaHandler.cs` + `src/Inforest.Desktop/Impresion/FrmPrecuentaImpresora.cs` + `tests/Inforest.Application.Tests/Impresion/ImprimirPrecuentaHandlerTests.cs` |
| `TIMPRESORA` | Table | `IImpresoraRepository` / `ImpresoraRepository` | Repository | COMPLETED | `src/Inforest.Infrastructure/Impresion/ImpresoraRepository.cs` |
| `frmCambiarContrasenia.frm` | Form | `CambiarPasswordHandler` + `FrmCambiarContrasenia` | Handler + WinForm | COMPLETED | `src/Inforest.Application/Seguridad/CambiarPasswordHandler.cs` + `src/Inforest.Desktop/POS/FrmCambiarContrasenia.cs` + `tests/Inforest.Application.Tests/Seguridad/CambiarPasswordHandlerTests.cs` |
| `frmPassword.frm` | Form | `FrmPassword` | WinForm | COMPLETED | `src/Inforest.Desktop/POS/FrmPassword.cs` |
| `AuthService` (cambio password) | Service | `IAuthService.CambiarPasswordAsync` | Service method | COMPLETED | `src/Inforest.Infrastructure/Security/AuthService.cs` |
| `frmLiquidacionDetalle.frm` | Form | `FrmLiquidacionCierre` + `CerrarTurnoHandler` | WinForm + Handler | MIGRATED | `src/Inforest.Desktop/Caja/FrmLiquidacionCierre.cs` + `src/Inforest.Application/Turno/TurnoHandlers.cs` |
| `CierreTurnoBreakdown` (domain) | — | `CierreTurnoBreakdown` | Value Object | MIGRATED | `src/Inforest.Domain/Entities/Caja/CierreTurnoBreakdown.cs` |
| `UPDATE MTURNO (cierre completo)` | SQL | `TurnoRepository.CerrarAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Turno/TurnoRepository.cs` |
| `TCAJA.lObligaCierre` | Flag | `CerrarTurnoHandler` (BR-CAJA-001) | Business Rule | MIGRATED | `src/Inforest.Application/Turno/TurnoHandlers.cs` |
| `TPARAMETRO.lActivaConsultaDescargo` | Flag | `CerrarTurnoHandler` (BR-CAJA-002) | Business Rule | MIGRATED | `src/Inforest.Application/Turno/TurnoHandlers.cs` |

## Componentes POS-FUNC-004 — Registro de Venta

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmVenta.frm` | Form | `FrmVenta` | WinForm | MIGRATED | `src/Inforest.Desktop/Ventas/FrmVenta.cs` |
| `frmSolicitudBoleta.frm` | Form | `FrmSolicitudBoleta` + `FrmSeleccionClienteBoleta` | WinForm modal | MIGRATED | `src/Inforest.Desktop/Ventas/FrmSolicitudBoleta.cs` + `src/Inforest.Desktop/Ventas/FrmSeleccionClienteBoleta.cs` |
| `frmSolicitudBoleta.frm` / opción “Sin datos” | Business rule | `ObtenerClienteGeneralBoletaHandler` | Handler | MIGRATED | `src/Inforest.Application/Ventas/SolicitudBoletaHandlers.cs` |
| `modPuntoVenta.bas` / `CLIENTEGENERAL` + `TCLIENTE` | Config + Table | `IClienteRepository.ObtenerPorIdentidadAsync` + `ObtenerProximoCodigoAsync` + `ClienteRepository` | Interface + Repository | MIGRATED | `src/Inforest.Application/Maestros/IMaestrosRepository.cs` + `src/Inforest.Infrastructure/Maestros/ClienteRepository.cs` |
| `TPARAMETRO.lBODato` | Flag | `ConfiguracionSistema.lBODato` + `ParametroRepository` | Domain Record + Repository | MIGRATED | `src/Inforest.Domain/Entities/Configuracion/ConfiguracionSistema.cs` + `src/Inforest.Infrastructure/Configuracion/ParametroRepository.cs` |
| `usp_Inforest_ValidaClienteSel` | SP | `IClienteRepository.ValidarCompatibilidadDocumentoAsync` + `ClienteRepository` + `EmitirDocumentoHandler` | Interface + Repository + Handler | MIGRATED | `src/Inforest.Application/Maestros/IMaestrosRepository.cs` + `src/Inforest.Infrastructure/Maestros/ClienteRepository.cs` + `src/Inforest.Application/Ventas/VentaHandlers.cs` |
| (tests BR-BOLETA-001..003) | — | `SolicitudBoletaHandlersTests` + `VentaHandlersTests` | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Ventas/SolicitudBoletaHandlersTests.cs` + `tests/Inforest.Application.Tests/Ventas/VentaHandlersTests.cs` |

## Componentes POS-FUNC-006 — Facturación y Notas de Crédito

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmNotaCredito.frm` | Form | `FrmNotaCredito` | WinForm | MIGRATED | `src/Inforest.Desktop/Ventas/FrmNotaCredito.cs` |
| `frmNotaCreditoDetalle.frm` | Form | `FrmNotaCreditoDetalle` + `EmitirNotaCreditoHandler` + `AnularNotaCreditoHandler` | WinForm + Handlers | MIGRATED | `src/Inforest.Desktop/Ventas/FrmNotaCreditoDetalle.cs` + `src/Inforest.Application/Ventas/NotaCreditoHandlers.cs` |
| `MNOTACREDITO` | Table | `NotaCredito` + `INotaCreditoRepository` + `NotaCreditoRepository` | Entity + Interface + Repository | MIGRATED | `src/Inforest.Domain/Entities/Ventas/NotaCredito.cs` + `src/Inforest.Infrastructure/Ventas/NotaCreditoRepository.cs` |
| `frmDocumento.frm` (Generación) | Form | `FrmDocumento` (visualización + anulación) | WinForm | MIGRATED | `src/Inforest.Desktop/Ventas/FrmDocumento.cs` |
| `frmFactura.frm` | Form | Cubierto por `CorrelativoDocumento` + `FrmDocumentoCorrelativo` | Domain Entity + WinForm | NOT_APPLICABLE | POS-FUNC-007 — ya implementado |
| (tests BR-NC-001..006) | — | `NotaCreditoTests` (9 tests) + `NotaCreditoHandlersTests` (10 tests) | xUnit | MIGRATED | `tests/Inforest.Domain.Tests/Ventas/NotaCreditoTests.cs` + `tests/Inforest.Application.Tests/Ventas/NotaCreditoHandlersTests.cs` |

## Componentes POS-FUNC-010 — Cliente y cuentas corrientes

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmNuevoCliente.frm` | Form | `FrmNuevoCliente` | WinForm | MIGRATED | `src/Inforest.Desktop/Clientes/FrmNuevoCliente.cs` |
| `frmCtaCte.frm` | Form | `FrmCtaCte` | WinForm | MIGRATED | `src/Inforest.Desktop/Clientes/FrmCtaCte.cs` |
| `frmCuentaCobrar.frm` | Form | `FrmCuentaCobrar` | WinForm | MIGRATED | `src/Inforest.Desktop/Clientes/FrmCuentaCobrar.cs` |
| `TCLIENTE` | Table | `Cliente` + `IClienteRepository` + `ClienteRepository` | Entity + Interface + Repository | MIGRATED | `src/Inforest.Domain/Entities/Maestros/Cliente.cs` + `src/Inforest.Infrastructure/Maestros/ClienteRepository.cs` |
| `TDELIVERY` (lClienteCtaCte=1) / `vCompania` | Table/View | `CuentaCorriente` + `ICuentaCorrienteRepository` + `CuentaCorrienteRepository` | Entity + Interface + Repository | MIGRATED | `src/Inforest.Domain/Entities/Maestros/CuentaCorriente.cs` + `src/Inforest.Infrastructure/Maestros/CuentaCorrienteRepository.cs` |
| `vDocumentoGrilla` (tEstadoDocumento='03') | View | `DocumentoPendienteCobro` + `ObtenerDocumentosPendientesCobroHandler` | Record + Handler | MIGRATED | `src/Inforest.Domain/Entities/Maestros/DocumentoPendienteCobro.cs` + `src/Inforest.Application/Maestros/CuentaCorrienteHandlers.cs` |
| (BR-CLIENTE-001..004, BR-CTACTE-001..003 tests) | — | `CuentaCorrienteHandlerTests` (7 tests) + `CuentaCorrienteTests` (8 tests) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Maestros/CuentaCorrienteHandlerTests.cs` + `tests/Inforest.Domain.Tests/Maestros/CuentaCorrienteTests.cs` |

## Componentes POS-FUNC-011 — Reservas

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmReserva.frm` | Form | `FrmReserva` | WinForm | MIGRATED | `src/Inforest.Desktop/Reservas/FrmReserva.cs` |
| `frmReservaDetalle.frm` | Form | `FrmReservaDetalle` + `CrearReservaHandler` + `ModificarReservaHandler` + `AnularReservaHandler` | WinForm + Handlers | MIGRATED | `src/Inforest.Desktop/Reservas/FrmReservaDetalle.cs` + `src/Inforest.Application/Reservas/ReservaHandlers.cs` |
| `TRESERVA` | Table | `Reserva` + `IReservaRepository` + `ReservaRepository` | Entity + Interface + Repository | MIGRATED | `src/Inforest.Domain/Entities/Reservas/Reserva.cs` + `src/Inforest.Infrastructure/Reservas/ReservaRepository.cs` |
| `spIns_MPEDIDO_RESERVA` | SP | `ConvertirReservaAPedidoHandler` + `IReservaRepository.ConvertirAPedidoAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Reservas/ReservaHandlers.cs` + `src/Inforest.Infrastructure/Reservas/ReservaRepository.cs` |
| (BR-RESERVA-001..004 tests) | — | `ReservaTests` (9 tests domain) + `CrearReservaHandlerTests` (2) + `AnularReservaHandlerTests` (3) + `ModificarReservaHandlerTests` (1) + `ObtenerReservasPorFechaHandlerTests` (2) = 17 tests | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Reservas/ReservaTests.cs` |

## Componentes POS-FUNC-012 — Delivery dependiente POS

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmPedidoDelivery.frm` | Form | `FrmPedidoDelivery` | WinForm | MIGRATED | `src/Inforest.Desktop/Delivery/FrmPedidoDelivery.cs` |
| `frmCentralPedidos.frm` (extendido) | Form | `CentralPedidosForm` (ConfirmarEntrega + RevertirEntrega + ModificarFecha) | WinForm | MIGRATED | `src/Inforest.Desktop/Delivery/CentralPedidosForm.cs` |
| (DeliveryHandlers) | — | `ConfirmarEntregaCentralHandler` + `RevertirEntregaCentralHandler` + `ModificarFechaProgramadaDeliveryHandler` + `ObtenerPedidosSeguimientoDeliveryHandler` | Handlers | MIGRATED | `src/Inforest.Application/Delivery/DeliveryHandlers.cs` |
| (IPedidoDeliveryRepository ext.) | — | 5 nuevos métodos `IPedidoDeliveryRepository` + `PedidoDeliveryRepository` | Interface + Repository | MIGRATED | `src/Inforest.Infrastructure/Delivery/PedidoDeliveryRepository.cs` |
| (tests BR-DEL-012..014) | — | `CentralPedidosHandlersTests` (10 tests) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Delivery/CentralPedidosHandlersTests.cs` |

## Componentes POS-FUNC-013 — Insumos/descargo

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmInsumo.frm` | Form | `FrmInsumo` | WinForm | MIGRATED | `src/Inforest.Desktop/Maestros/FrmInsumo.cs` |
| `frmInsumoDetalle.frm` | Form | `FrmInsumoDetalle` + `AgregarInsumoHandler` + `ModificarInsumoHandler` + `EliminarInsumoHandler` | WinForm + Handlers | MIGRATED | `src/Inforest.Desktop/Maestros/FrmInsumoDetalle.cs` + `src/Inforest.Application/Maestros/InsumoHandlers.cs` |
| `TINSUMO` | Table | `Insumo` + `IInsumoRepository` + `InsumoRepository` | Entity + Interface + Repository | MIGRATED | `src/Inforest.Domain/Entities/Maestros/Insumo.cs` + `src/Inforest.Infrastructure/Maestros/InsumoRepository.cs` |
| `USP_LISTARINSUMOS` | SP | `ListarInsumosHandler` + `IInsumoRepository.ObtenerTodosAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Maestros/InsumoHandlers.cs` + `src/Inforest.Infrastructure/Maestros/InsumoRepository.cs` |
| `usp_agregarinsumos` | SP | `AgregarInsumoHandler` + `IInsumoRepository.AgregarAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Maestros/InsumoHandlers.cs` + `src/Inforest.Infrastructure/Maestros/InsumoRepository.cs` |
| `USP_MODIFICARINSUMOS` | SP | `ModificarInsumoHandler` + `IInsumoRepository.ModificarAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Maestros/InsumoHandlers.cs` + `src/Inforest.Infrastructure/Maestros/InsumoRepository.cs` |
| `USP_ELIMINARINSUMOS` | SP | `EliminarInsumoHandler` + `IInsumoRepository.EliminarAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Maestros/InsumoHandlers.cs` + `src/Inforest.Infrastructure/Maestros/InsumoRepository.cs` |
| `clsAlmacen.cls` (descargo automatico) | Class | `IInventoryGateway` + `InventoryGateway` (existente — BR-008) | Interface + Class | MIGRATED | `src/Inforest.Infrastructure/Almacen/InventoryGateway.cs` |
| (tests BR-INSUMO-001..004) | — | `InsumoTests` (8 tests domain) + `InsumoHandlerTests` (10 tests handler) | xUnit | MIGRATED | `tests/Inforest.Domain.Tests/InsumoTests.cs` + `tests/Inforest.Application.Tests/Maestros/InsumoHandlerTests.cs` |

## Componentes POS-FUNC-014 — Importación de pedidos externos

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmImportacionRequerimientos.frm` | Form | `FrmImportacionRequerimientos` | WinForm | MIGRATED | `src/Inforest.Desktop/Almacen/FrmImportacionRequerimientos.cs` |
| `frmImportacionRequerimientoDetalle.frm` | Form | `FrmImportacionRequerimientoDetalle` | WinForm | MIGRATED | `src/Inforest.Desktop/Almacen/FrmImportacionRequerimientoDetalle.cs` |
| `vRequerimiento` (ALMACEN DB) | View | `RequerimientoAlmacen` + `IRequerimientoAlmacenRepository` + `RequerimientoAlmacenRepository` | Entity + Interface + Repository | MIGRATED | `src/Inforest.Domain/Entities/Almacen/RequerimientoAlmacen.cs` + `src/Inforest.Infrastructure/Almacen/RequerimientoAlmacenRepository.cs` |
| `vRequerimiento` (detalle) | View | `DetalleRequerimientoAlmacen` | Domain Entity | MIGRATED | `src/Inforest.Domain/Entities/Almacen/DetalleRequerimientoAlmacen.cs` |
| `MREQUERIMIENTO.tPedido` / `lPedido` | Field | `IRequerimientoAlmacenRepository.MarcarImportadoAsync` | Repository Method | MIGRATED | `src/Inforest.Infrastructure/Almacen/RequerimientoAlmacenRepository.cs` |
| `TRUTAAREA.lImportarPV` | Flag | `RequerimientoAlmacenRepository.ObtenerPendientesAsync` (BR-IMPORT-001) | Business Rule | MIGRATED | `src/Inforest.Infrastructure/Almacen/RequerimientoAlmacenRepository.cs` |
| `spIns_MPEDIDO` (importación) | SP | `ImportacionPedidoGateway.CrearPedidoDesdeRequerimientoAsync` | Gateway | MIGRATED | `src/Inforest.Infrastructure/Almacen/ImportacionPedidoGateway.cs` |
| `InsertaProducto()` (DPEDIDO insert) | Function | `ImportacionPedidoGateway` (precio por canal + INSERT DPEDIDO) | Gateway | MIGRATED | `src/Inforest.Infrastructure/Almacen/ImportacionPedidoGateway.cs` |
| `MPEDIDO` update estado '03' (cancelar) | SQL | `ImportacionPedidoGateway.CancelarPedidoAsync` (BR-IMPORT-003) | Business Rule | MIGRATED | `src/Inforest.Infrastructure/Almacen/ImportacionPedidoGateway.cs` |
| Query buscar pedidos pendientes | Logic | `ObtenerRequerimientosPendientesHandler` | Handler | MIGRATED | `src/Inforest.Application/Almacen/ImportacionRequerimientoHandlers.cs` |
| Ver detalle requerimiento | Logic | `ObtenerDetalleRequerimientoHandler` | Handler | MIGRATED | `src/Inforest.Application/Almacen/ImportacionRequerimientoHandlers.cs` |
| Importar requerimiento como pedido | Logic | `ImportarRequerimientoHandler` (BR-IMPORT-001..004) | Handler | MIGRATED | `src/Inforest.Application/Almacen/ImportacionRequerimientoHandlers.cs` |
| `IImportacionPedidoGateway` | Interface | `IImportacionPedidoGateway` | Interface | MIGRATED | `src/Inforest.Application/Interfaces/IImportacionPedidoGateway.cs` |
| (tests BR-IMPORT-001..004) | — | `RequerimientoAlmacenTests` (8 tests domain) + `ImportacionRequerimientoHandlerTests` (9 tests handler) | xUnit | MIGRATED | `tests/Inforest.Domain.Tests/RequerimientoAlmacenTests.cs` + `tests/Inforest.Application.Tests/Almacen/ImportacionRequerimientoHandlerTests.cs` |

## Componentes POS-FUNC-015 — Mensajería cocina/KDS

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmMensajeCocina.frm` | Form | `FrmMensajeCocina` | WinForm | COMPLETED | `src/Inforest.Desktop/Kitchen/FrmMensajeCocina.cs` |
| `frmMensajeCocinaDetalle.frm` | Form | `FrmMensajeCocinaDetalle` + `AgregarMensajeCocinaHandler` + `ModificarMensajeCocinaHandler` + `EliminarMensajeCocinaHandler` | WinForm + Handlers | COMPLETED | `src/Inforest.Desktop/Kitchen/FrmMensajeCocinaDetalle.cs` + `src/Inforest.Application/Kitchen/MensajeCocinaHandlers.cs` |
| `TMENSAJECOCINA` | Table | `MensajeCocina` + `IMensajeCocinaRepository` + `MensajeCocinaRepository` | Entity + Interface + Repository | COMPLETED | `src/Inforest.Domain/Entities/Cocina/MensajeCocina.cs` + `src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs` |
| `USP_LISTARMENSAJES` | SP | `ObtenerMensajesCocinaHandler` + `IMensajeCocinaRepository.ObtenerMensajesAsync` | Handler + Repository | COMPLETED | `src/Inforest.Application/Kitchen/MensajeCocinaHandlers.cs` + `src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs` |
| `USP_AGREGARMENSAJE` | SP | `AgregarMensajeCocinaHandler` + `IMensajeCocinaRepository.AgregarAsync` | Handler + Repository | COMPLETED | `src/Inforest.Application/Kitchen/MensajeCocinaHandlers.cs` + `src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs` |
| `USP_MODIFICARMENSAJE` | SP | `ModificarMensajeCocinaHandler` + `IMensajeCocinaRepository.ModificarAsync` | Handler + Repository | COMPLETED | `src/Inforest.Application/Kitchen/MensajeCocinaHandlers.cs` + `src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs` |
| `USP_ELIMINARRMENSAJES` | SP | `EliminarMensajeCocinaHandler` + `IMensajeCocinaRepository.EliminarAsync` | Handler + Repository | COMPLETED | `src/Inforest.Application/Kitchen/MensajeCocinaHandlers.cs` + `src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs` |
| `USP_CERRAR_MENSAJES_CIERRETURNO` | SP | `CerrarTurnoHandler` + `IMensajeCocinaRepository.CerrarActivosPorCajaAsync` | Handler + Repository | COMPLETED | `src/Inforest.Application/Turno/TurnoHandlers.cs` + `src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs` |
| `frmLiquidacionDetalle.frm` (cierre mensajes cocina) | Form rule | `FrmLiquidacionCierre` + `CerrarTurnoHandler` | WinForm + Handler | COMPLETED | `src/Inforest.Desktop/Caja/FrmLiquidacionCierre.cs` + `src/Inforest.Application/Turno/TurnoHandlers.cs` |
| (tests BR-MSGCOC-001..005) | — | `MensajeCocinaTests` + `MensajeCocinaHandlerTests` + `MensajeCocinaRepositoryTests` + `CerrarTurnoHandlerTests` | xUnit | COMPLETED | `tests/Inforest.Domain.Tests/Cocina/MensajeCocinaTests.cs`, `tests/Inforest.Application.Tests/Kitchen/MensajeCocinaHandlerTests.cs`, `tests/Inforest.Infrastructure.Tests/Kitchen/MensajeCocinaRepositoryTests.cs`, `tests/Inforest.Application.Tests/Turno/CerrarTurnoHandlerTests.cs` |

## Componentes POS-FUNC-018 — Recibo de Ingresos/Egresos

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmReciboIngreso.frm` | Form | `FrmReciboIngreso` | WinForm | COMPLETED | `src/Inforest.Desktop/Caja/Recibos/FrmReciboIngreso.cs` |
| `frmReciboIngresoDetalle.frm` | Form | `FrmReciboIngresoDetalle` | WinForm | COMPLETED | `src/Inforest.Desktop/Caja/Recibos/FrmReciboIngresoDetalle.cs` |
| `frmReciboEgreso.frm` | Form | `FrmReciboEgreso` | WinForm | COMPLETED | `src/Inforest.Desktop/Caja/Recibos/FrmReciboEgreso.cs` |
| `frmReciboEgresoDetalle.frm` | Form | `FrmReciboEgresoDetalle` | WinForm | COMPLETED | `src/Inforest.Desktop/Caja/Recibos/FrmReciboEgresoDetalle.cs` |
| `MINGRESO` | Table | `ReciboIngreso` + `IReciboIngresoRepository` + `ReciboIngresoRepository` | Entity + Interface + Repository | COMPLETED | `src/Inforest.Domain/Entities/Caja/ReciboIngreso.cs` + `src/Inforest.Infrastructure/Caja/ReciboIngresoRepository.cs` |
| `MEGRESO` | Table | `ReciboEgreso` + `IReciboEgresoRepository` + `ReciboEgresoRepository` | Entity + Interface + Repository | COMPLETED | `src/Inforest.Domain/Entities/Caja/ReciboEgreso.cs` + `src/Inforest.Infrastructure/Caja/ReciboEgresoRepository.cs` |
| `frmReciboIngreso.frm` / lógica registrar | Business rule | `RegistrarIngresoHandler` (BR-RECIBO-001..006) | Handler | COMPLETED | `src/Inforest.Application/Caja/ReciboIngresoHandlers.cs` |
| `frmReciboIngreso.frm` / lógica anular | Business rule | `AnularIngresoHandler` (BR-RECIBO-005) | Handler | COMPLETED | `src/Inforest.Application/Caja/ReciboIngresoHandlers.cs` |
| `frmReciboEgreso.frm` / lógica registrar | Business rule | `RegistrarEgresoHandler` (BR-RECIBO-007..012) | Handler | COMPLETED | `src/Inforest.Application/Caja/ReciboEgresoHandlers.cs` |
| `frmReciboEgreso.frm` / lógica anular | Business rule | `AnularEgresoHandler` (BR-RECIBO-011) | Handler | COMPLETED | `src/Inforest.Application/Caja/ReciboEgresoHandlers.cs` |
| `spRep_ReciboEgreso` | SP (report) | `ObtenerEgresosHandler` (reutiliza SP via Dapper) | Handler + Repository | COMPLETED | `src/Inforest.Application/Caja/ReciboEgresoHandlers.cs` + `src/Inforest.Infrastructure/Caja/ReciboEgresoRepository.cs` |
| (tests BR-RECIBO-001..012) | — | `ReciboIngresoTests` (9) + `ReciboEgresoTests` (8) + `ReciboIngresoHandlerTests` (8) + `ReciboEgresoHandlerTests` (8) | xUnit | COMPLETED | `tests/Inforest.Domain.Tests/Caja/ReciboIngresoTests.cs`, `tests/Inforest.Domain.Tests/Caja/ReciboEgresoTests.cs`, `tests/Inforest.Application.Tests/Caja/ReciboIngresoHandlerTests.cs`, `tests/Inforest.Application.Tests/Caja/ReciboEgresoHandlerTests.cs` |
| `frmDivision.frm` | Form | `FrmDivision` | WinForm | COMPLETED | `src/Inforest.Desktop/Pedidos/FrmDivision.cs` |
| `frmDivision.frm` / sesión in-memory | Design pattern | `SesionDivision` + `ItemDivision` | Domain entity | COMPLETED | `src/Inforest.Domain/Entities/Ventas/SesionDivision.cs` + `ItemDivision.cs` |
| `frmDivision.frm` / lógica Disgregar | Business rule | `DisgregarItemHandler` (BR-DIV-003) | Handler | COMPLETED | `src/Inforest.Application/Pedidos/DivisionPedidoHandlers.cs` |
| `frmDivision.frm` / lógica Compartir | Business rule | `CompartirPedidoHandler` (BR-DIV-004) | Handler | COMPLETED | `src/Inforest.Application/Pedidos/DivisionPedidoHandlers.cs` |
| `frmDivision.frm` / lógica Mover | Business rule | `MoverItemHandler` | Handler | COMPLETED | `src/Inforest.Application/Pedidos/DivisionPedidoHandlers.cs` |
| `frmDivision.frm` / confirmar — validación estado | Business rule | `ConfirmarDivisionHandler` (BR-DIV-001) | Handler | COMPLETED | `src/Inforest.Application/Pedidos/DivisionPedidoHandlers.cs` |
| `frmDivision.frm` / confirmar — validación monto máximo | Business rule | `ConfirmarDivisionHandler` (BR-DIV-002) | Handler | COMPLETED | `src/Inforest.Application/Pedidos/DivisionPedidoHandlers.cs` |
| `frmDivision.frm` / confirmar — SQL (DPEDIDO/MPEDIDO) | SQL | `DivisionPedidoRepository.ConfirmarDivisionAsync` (BR-DIV-001..005) | Repository | COMPLETED | `src/Inforest.Infrastructure/Pedidos/DivisionPedidoRepository.cs` |
| `frmDetallePedido.frm` | Form | `FrmDetallePedido` | WinForm | MIGRATED | `src/Inforest.Desktop/Pedidos/FrmDetallePedido.cs` |
| `frmDetallePedido.frm` / vPedidoDetalle (BR-PEDIDO-005) | SQL/View | `IPedidoReadRepository.ObtenerDetalleExtendidoAsync` + `PedidoRepository` impl | Repository | MIGRATED | `src/Inforest.Application/Pedidos/IPedidoReadRepository.cs` + `src/Inforest.Infrastructure/Pedidos/PedidoRepository.cs` |
| `frmDetallePedido.frm` / vPedidoCombo (BR-PEDIDO-006) | SQL/View | `IPedidoReadRepository.ObtenerCombosAsync` + `PedidoRepository` impl | Repository | MIGRATED | `src/Inforest.Application/Pedidos/IPedidoReadRepository.cs` + `src/Inforest.Infrastructure/Pedidos/PedidoRepository.cs` |
| `frmDetallePedido.frm` / handler detalle extendido | Business rule | `ObtenerDetallePedidoExtendidoHandler` (BR-PEDIDO-005/BR-PEDIDO-006) | Handler | MIGRATED | `src/Inforest.Application/Pedidos/DetallePedidoHandlers.cs` |
| `frmDetallePedido.frm` / DTOs vista | Data Transfer | `ItemPedidoVista` + `ComboPedidoVista` + `DetallePedidoExtendidoResultado` | Records | MIGRATED | `src/Inforest.Application/Pedidos/DetallePedidoVista.cs` |
| `frmJuntarMesas.frm` | Form | `FrmJuntarMesas` | WinForm | MIGRATED | `src/Inforest.Desktop/Pedidos/FrmJuntarMesas.cs` |
| `frmJuntarMesas.frm` / resultado para caller | Design pattern | `JuntarMesasResult` record (BR-JUNTAR-002/BR-JUNTAR-003) | Record | MIGRATED | `src/Inforest.Desktop/Pedidos/FrmJuntarMesas.cs` |
| `frmJuntarMesas.frm` / selección visual de mesas | UI Logic | Color-coded buttons por `EstadoMesa` (BR-JUNTAR-001) | WinForms UI | MIGRATED | `src/Inforest.Desktop/Pedidos/FrmJuntarMesas.cs` |
| `spUpdate_DPEDIDO_Ina` | SP | reutilizado vía Dapper en `DivisionPedidoRepository` | Repository call | COMPLETED | `src/Inforest.Infrastructure/Pedidos/DivisionPedidoRepository.cs` |
| (tests BR-DIV-001..005) | — | `DivisionPedidoTests` (11 domain) + `DivisionPedidoHandlerTests` (11 handler) | xUnit | COMPLETED | `tests/Inforest.Domain.Tests/Pedidos/DivisionPedidoTests.cs` + `tests/Inforest.Application.Tests/Pedidos/DivisionPedidoHandlerTests.cs` |

## Componentes POS-FUNC-027 — Cambio de Documento

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmCambio.frm` | Form | `FrmCambioDocumento` | WinForm | MIGRATED | `src/Inforest.Desktop/Ventas/FrmCambioDocumento.cs` |
| `frmCambio.frm` / TTIPODOCUMENTOIMPRESORA + vTipoDocumentoImpresora | SQL | `ICambioDocumentoRepository.ObtenerTiposDocumentoAsync` + `CambioDocumentoRepository` | Interface + Class | MIGRATED | `src/Inforest.Application/Ventas/ICambioDocumentoRepository.cs` + `src/Inforest.Infrastructure/Ventas/CambioDocumentoRepository.cs` |
| `frmCambio.frm` / cmdOpcion_Click(0) — cambio documento | Business Logic | `CambiarDocumentoHandler` + `CambiarDocumentoCommand` (BR-CAMBIO-001..005) | Handler | MIGRATED | `src/Inforest.Application/Ventas/CambiarDocumentoHandlers.cs` |
| `frmCambio.frm` / validación monto máximo | BR-CAMBIO-002 | `CambiarDocumentoHandler.ObtenerMontosValidacionAsync` | Handler logic | MIGRATED | `src/Inforest.Application/Ventas/CambiarDocumentoHandlers.cs` |
| (tests BR-CAMBIO-001..005) | — | `CambiarDocumentoHandlerTests` (8 tests) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Ventas/CambiarDocumentoHandlerTests.cs` |

## Componentes POS-FUNC-028 — Actualizar Datos Pedido

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmUpdateDatosPedido.frm` (`FrmActualizarPedidos`) | Form | `FrmActualizarDatosPedido` | WinForm | MIGRATED | `src/Inforest.Desktop/Ventas/FrmActualizarDatosPedido.cs` |
| `usp_ActualizarCabPeDoc` @opcion='1' (cortesía) | SP | `ActualizarCortesiaPedidoHandler` + `IActualizarDatosPedidoRepository.ActualizarCortesiaAsync` (BR-ACTPED-001/002) | Handler + Interface | MIGRATED | `src/Inforest.Application/Ventas/ActualizarDatosPedidoHandlers.cs` + `src/Inforest.Infrastructure/Ventas/ActualizarDatosPedidoRepository.cs` |
| `usp_ActualizarCabPeDoc` @opcion='2' (canal venta) | SP | `ActualizarCanalVentaPedidoHandler` + `IActualizarDatosPedidoRepository.ActualizarCanalVentaAsync` (BR-ACTPED-001/003) | Handler + Interface | MIGRATED | `src/Inforest.Application/Ventas/ActualizarDatosPedidoHandlers.cs` + `src/Inforest.Infrastructure/Ventas/ActualizarDatosPedidoRepository.cs` |
| (tests BR-ACTPED-001..003) | — | `ActualizarDatosPedidoHandlerTests` (8 tests) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Ventas/ActualizarDatosPedidoHandlerTests.cs` |

---

## Componentes POS-FUNC-029 — Lista Motivos Anulación

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmListaMotivos.frm` | Form | `FrmListaMotivos` | WinForm modal | MIGRATED | `src/Inforest.Desktop/Ventas/FrmListaMotivos.cs` |
| `vMotivoANULACION` / `TTABLA WHERE TTABLA='MOTIVOANULACION'` | View / Table | `MotivoAnulacion` | Domain Record | MIGRATED | `src/Inforest.Domain/Entities/Configuracion/MotivoAnulacion.cs` |
| `vMotivoANULACION` | View | `IMotivoAnulacionRepository` + `MotivoAnulacionRepository` | Interface + Class | MIGRATED | `src/Inforest.Application/Interfaces/IMotivoAnulacionRepository.cs` + `src/Inforest.Infrastructure/Ventas/MotivoAnulacionRepository.cs` |
| `Form_Load` → RecordCount=0 → MsgBox (BR-MOTIVO-002) | Logic | `ObtenerMotivosAnulacionHandler` | Handler | MIGRATED | `src/Inforest.Application/Ventas/MotivoAnulacionHandlers.cs` |
| `cmdEliminacion_Click` → `sCodigo` + `sDescrip` + `wEnter=True` (BR-MOTIVO-003) | Logic | `FrmListaMotivos.MotivoSeleccionado` | Property | MIGRATED | `src/Inforest.Desktop/Ventas/FrmListaMotivos.cs` |
| (tests BR-MOTIVO-001..003) | — | `MotivoAnulacionHandlerTests` (5 tests) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Ventas/MotivoAnulacionHandlerTests.cs` |

---

## Componentes POS-FUNC-021 — Tarjetas RFID / Proximidad

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `FrmTarjetaAproximidad.frm` + `FrmTarjetaAproximidadDetalle.frm` | Form | `FrmTarjetaProximidad` | WinForm | MIGRATED | `src/Inforest.Desktop/Clientes/FrmTarjetaProximidad.cs` |
| `TTARJETASRFID` | Table | `TarjetaProximidad` + `ITarjetaProximidadRepository` + `TarjetaProximidadRepository` | Entity + Interface + Class | MIGRATED | `src/Inforest.Domain/Entities/Delivery/TarjetaProximidad.cs` + `src/Inforest.Application/Delivery/TarjetaProximidadHandlers.cs` + `src/Inforest.Infrastructure/Delivery/TarjetaProximidadRepository.cs` |
| `TMOVIMIENTOTARJETASRFID` | Table | `MovimientoTarjetaProximidad` + `ObtenerMovimientosTarjetaProximidadHandler` | Record + Handler | MIGRATED | `src/Inforest.Domain/Entities/Delivery/TarjetaProximidad.cs` + `src/Inforest.Application/Delivery/TarjetaProximidadHandlers.cs` |
| `TDELIVERY` (búsqueda cliente asociado) | Table | `IClienteDeliveryRepository` + selector modal en `FrmTarjetaProximidad` | Interface + UI | MIGRATED | `src/Inforest.Domain/Repositories/IClienteDeliveryRepository.cs` + `src/Inforest.Desktop/Clientes/FrmTarjetaProximidad.cs` |
| `FrmRecargarTarjeta.frm` | Form | `FrmRecargarTarjeta` | WinForm | MIGRATED | `src/Inforest.Desktop/Clientes/FrmRecargarTarjeta.cs` |
| `FrmRecargarTarjetaDetalle.frm` | Form | `FrmRecargarTarjetaDetalle` | WinForm | MIGRATED | `src/Inforest.Desktop/Clientes/FrmRecargarTarjetaDetalle.cs` |
| `usp_Inforest_ObtieneRecargas` | Stored Procedure | `ObtenerRecargasTarjetaHandler` + `IRecargaTarjetaRepository.ObtenerMovimientosAsync` | Handler + Interface | MIGRATED | `src/Inforest.Application/Delivery/RecargaTarjetaHandlers.cs` + `src/Inforest.Infrastructure/Delivery/RecargaTarjetaRepository.cs` |
| `TMOVIMIENTOTARJETASRFID` (tipo `R`) | Table | `IRecargaTarjetaRepository.RegistrarRecargaAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Delivery/RecargaTarjetaRepository.cs` |
| `TTARJETASRFID.MontoDisponible` | Table | `RegistrarRecargaTarjetaHandler` + `RecargaTarjetaRepository` | Handler + Repository | MIGRATED | `src/Inforest.Application/Delivery/RecargaTarjetaHandlers.cs` + `src/Inforest.Infrastructure/Delivery/RecargaTarjetaRepository.cs` |
| (tests BR-RFID-005..007) | — | `RecargaTarjetaHandlerTests` (3) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Delivery/RecargaTarjetaHandlerTests.cs` |
| (tests BR-RFID-001..004) | — | `TarjetaProximidadTests` (7) + `TarjetaProximidadHandlerTests` (7) | xUnit | MIGRATED | `tests/Inforest.Domain.Tests/Delivery/TarjetaProximidadTests.cs` + `tests/Inforest.Application.Tests/Delivery/TarjetaProximidadHandlerTests.cs` |
| `frmMesas.frm` | Form | `FrmMesas` | WinForms Form | MIGRATED | `src/Inforest.Desktop/Maestros/FrmMesas.cs` |
| `EstadoMesa` (VB6 '01'..'06') | Enum | `EstadoMesa` (Libre/Ocupada/Reservada/Sucia/Bloqueada/FueraDeServicio/EnCuenta) | Enum | MIGRATED | `src/Inforest.Domain/Entities/Maestros/EstadoMesa.cs` |
| `TMESA.tEstado` ('01'..'06') | Table column | `MesaRepository` CASE mapping | Repository | MIGRATED | `src/Inforest.Infrastructure/Maestros/MesaRepository.cs` |
| `MPEDIDO` (activos sin mesa) | Table | `ObtenerPedidosSinMesaHandler` + `IPedidoReadRepository.ObtenerActivosSinMesaAsync` | Handler + Interface | MIGRATED | `src/Inforest.Application/Pedidos/PedidoHandlers.cs` |
| `PedidoSinMesaVista` | — | `PedidoSinMesaVista` record | DTO | MIGRATED | `src/Inforest.Application/Pedidos/DetallePedidoVista.cs` |
| (tests BR-MESAS-001..005) | — | `ObtenerPedidosSinMesaHandlerTests` (3) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Pedidos/ObtenerPedidosSinMesaHandlerTests.cs` |
| `frmNumPad.frm` | Form | `FrmNumPad` | WinForms Dialog | MIGRATED | `src/Inforest.Desktop/Shared/FrmNumPad.cs` |
| `frmCambioPropina.frm` | Form | `FrmCambioPropina` | WinForms Dialog | MIGRATED | `src/Inforest.Desktop/Ventas/FrmCambioPropina.cs` |
| `tPropina` ('01'/'02') + `nPropina` | VB6 global vars | `FrmCambioPropina.TipoPropina` + `PropinaMN`/`PropinaME` | Properties | MIGRATED | `src/Inforest.Desktop/Ventas/FrmCambioPropina.cs` |
| `frmMesaConsulta.frm` | Form | `FrmMesaConsulta` | WinForms Form | MIGRATED | `src/Inforest.Desktop/Maestros/FrmMesaConsulta.cs` |
| `UPDATE TMESA SET tEstadoMesa` | SQL | `IMesaRepository.CambiarEstadoAsync` + `MesaRepository.CambiarEstadoAsync` | Repository | MIGRATED | `src/Inforest.Application/Maestros/IMaestrosRepository.cs` + `src/Inforest.Infrastructure/Maestros/MesaRepository.cs` |
| `sTipo="V"/"M"/default` | VB6 global var | `ModoConsulta` enum (Visual/Mover/Seleccionar) | Enum | MIGRATED | `src/Inforest.Desktop/Maestros/FrmMesaConsulta.cs` |
| (tests BR-MESACONSULTA-001..004) | — | `CambiarEstadoMesaHandlerTests` (4) | xUnit | MIGRATED | `tests/Inforest.Application.Tests/Maestros/CambiarEstadoMesaHandlerTests.cs` |
