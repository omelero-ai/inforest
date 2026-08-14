# Matriz de Trazabilidad — INFOREST Legacy → .NET 8

> Status: IN_PROGRESS — baseline transversal de Fase 3 validado; la matriz refleja equivalencias parciales y gaps controlados.
>
> Última actualización: 2026-08-12

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
| `InfoRest.vbp` / `InfoRest.exe` | VBP Project | — | — | NOT_STARTED | — |
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
| `modKDS.bas` | Integration | `IProduccionCocinaService` + `KdsXmlDispatcher` + `KdsLegacyGateway` | Interface + Classes | IN_PROGRESS | `modern-net8/src/Inforest.Application/Interfaces/IProduccionCocinaService.cs`, `modern-net8/src/Inforest.Infrastructure/Kitchen/KdsXmlDispatcher.cs`, `modern-net8/src/Inforest.Infrastructure/Kitchen/KdsLegacyGateway.cs` |
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
| `MPEDIDO` | Table | — | — | NOT_STARTED | Cabecera pedido |
| `DPEDIDO` | Table | `DetallePedido`, `PedidoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Infrastructure/Pedidos/ |
| `CPEDIDO` | Table | — | — | NOT_STARTED | Combos en pedido |
| `APEDIDO` | Table | — | — | NOT_STARTED | Auditoría anulados |
| `MDOCUMENTO` | Table | `Documento`, `IDocumentoRepository`, `DocumentoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Ventas/ |
| `DDOCUMENTO` | Table | `DetalleDocumento` | Domain Entity | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Ventas/ |
| `MTURNO` | Table | `Turno`, `ITurnoRepository`, `TurnoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Infrastructure/Turno/ |
| `TCAJA` | Table | `ConfiguracionCaja`, `IParametroRepository.ObtenerConfiguracionCajaAsync` | Domain Record + Repository | IN_PROGRESS | modern-net8/src/Inforest.Infrastructure/Configuracion/ |
| `TPRODUCTO` | Table | `ProductoMaestro`, `IProductoMaestroRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TGRUPO` | Table | `GrupoProducto`, `IGrupoProductoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TSUBGRUPO` | Table | `SubGrupoProducto`, `ISubGrupoProductoRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TCLIENTE` | Table | `Cliente`, `IClienteRepository` | Entity + Repository | IN_PROGRESS | modern-net8/src/Inforest.Domain/Entities/Maestros/ |
| `TUSUARIO` | Table | — | — | NOT_STARTED | Usuarios |
| `TPARAMETRO` | Table | `ConfiguracionSistema`, `IParametroRepository`, `ParametroService` | Domain Record + Repository + Service | IN_PROGRESS | modern-net8/src/Inforest.Infrastructure/Configuracion/ |
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
| `frmAcceso.frm` — bloqueo intentos (BR-POS-006-LOCK) | Business Rule | `FrmLoginPOS` (`_intentosFallidos` + `LoginPolicy.MaxIntentosFallidos`) | UI guard | MIGRATED | `src/Inforest.Desktop/POS/FrmLoginPOS.cs`, `src/Inforest.Application/Seguridad/LoginPolicy.cs` |
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
| `frmCentralPedidos.frm` | Form | `CentralPedidosForm` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Delivery/CentralPedidosForm.cs` |
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
| `IPedidoDeliveryRepository` | Interface | `IPedidoDeliveryRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/IPedidoDeliveryRepository.cs` |
| `IMotorizadoRepository` | Interface | `IMotorizadoRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/IMotorizadoRepository.cs` |
| `ILocalRepository` | Interface | `ILocalRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/ILocalRepository.cs` |
| `ICentralPedidosRepository` | Interface | `ICentralPedidosRepository` | Interface | COMPLETED | `src/Inforest.Domain/Repositories/ICentralPedidosRepository.cs` |
| `sp_Inforest_PedidosCentralPedido` | SP | `CentralPedidosRepository.ObtenerPedidosCentralAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Delivery/CentralPedidosRepository.cs` |
| `sp_CD_Modificar_EstadoDelivery_Cabecera` | SP | `CentralPedidosRepository.ModificarEstadoDeliveryAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Delivery/CentralPedidosRepository.cs` |
| `sp_UpdFotoDelivery` | SP | `ActualizarFotoClienteDeliveryHandler` + `ClienteDeliveryRepository.ActualizarFotoAsync` | Handler + Repository | MIGRATED | `src/Inforest.Application/Delivery/DeliveryHandlers.cs`, `src/Inforest.Infrastructure/Delivery/ClienteDeliveryRepository.cs` |
| `spRep_AnaliticoMotorizadoIntegrado` | SP | *(reporte FastReport pendiente)* | Report | NOT_STARTED | — |
| `frmOrdenesConsola.frm` | Form | `ObtenerOrdenesExternasHandler` + `RappiOrderAdapter` | Handler+Service | IN_PROGRESS | `src/Inforest.Application/Delivery/RappiHandlers.cs` |
| `(test etapa 9 domain)` | — | `DeliveryDomainTests` (12 tests) + `MotorizadoDomainTests` (11 tests) | xUnit | IN_PROGRESS | `tests/Inforest.Domain.Tests/Delivery/` |
| `(test etapa 9 application)` | — | `DeliveryHandlersTests` (4 tests) + `MotorizadoHandlersTests` (5 tests) | xUnit | IN_PROGRESS | `tests/Inforest.Application.Tests/` |

| **Etapa 10 — Reportes, Consultas y Salidas Operativas** | | | | | |
| `spRep_Comanda` | SP | `ReporteRepository.ObtenerComandaAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrComandaD.Dsr` / `dsrComandaR.Dsr` | Crystal | `RepComanda.frx` / `RepComandaDetallado.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `frmRepComanda.frm` | Form | `FrmComandaReporte.cs` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Forms/Reportes/FrmComandaReporte.cs` |
| `spRep_Propina` | SP | `ReporteRepository.ObtenerPropinaAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrPropinaD.Dsr` / `dsrPropinaR.Dsr` | Crystal | `RepPropina.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `frmRepPropina.frm` | Form | `FrmPropinaReporte.cs` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Forms/Reportes/FrmPropinaReporte.cs` |
| `spRep_PrincipalCliente` | SP | `ReporteRepository.ObtenerPrincipalCliente*Async` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrPrincipalD.Dsr` / `dsrPrincipalR.Dsr` | Crystal | `RepPrincipalClienteDetalle.frx` / `ResumenRepPrincipalCliente.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
| `spRep_CtaCteIntegrado` | SP | `ReporteRepository.ObtenerCtaCteIntegradoAsync` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Reportes/ReporteRepository.cs` |
| `dsrRepCtaCteIntegradoConsolidado.Dsr` | Crystal | `RepCtaCteIntegradoConsolidado.frx` | FastReport | NOT_STARTED | `src/Inforest.Desktop/reports/templates/` |
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
| `(test etapa 10 application)` | — | `ReportesHandlersTests` (14 tests) | xUnit | COMPLETED | `tests/Inforest.Application.Tests/Reportes/` |

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
| `frmPago.frm`, `frmPagoPinPad.frm` | Form | `FrmPago` / `FrmPagoPinPad` | WinForm | NOT_STARTED | `src/Inforest.Desktop/Forms/` |
| (tests P3-11) | — | `NullHardwareServicesTests` (15 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| (tests P3-11) | — | `FacturacionElectronicaFactoryTests` (7 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| (tests P3-11) | — | `PaisPolicyTests` (8 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| (tests P3-11) | — | `DomainHardwareEntityTests` (8 tests) | xUnit | COMPLETED | `tests/Inforest.Infrastructure.Tests/Hardware/` |
| `frmDocumentoCorrelativo.frm` + `modPuntoVenta.bas` | Form/Module | `CorrelativoDocumento` | Domain Entity | IN_PROGRESS | `src/Inforest.Domain/Entities/Ventas/CorrelativoDocumento.cs` |
| `TTIPODOCUMENTOIMPRESORA` | Table | `ICorrelativoRepository` / `CorrelativoRepository` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Ventas/CorrelativoRepository.cs` |
| `frmDocumentoCorrelativo.frm` | Form | `FrmDocumentoCorrelativo` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/Ventas/FrmDocumentoCorrelativo.cs` |
| `frmPrecuentaImpresora.frm` | Form | `ImprimirPrecuentaHandler` / `FrmPrecuentaImpresora` | Handler + WinForm | IN_PROGRESS | `src/Inforest.Application/Impresion/ImprimirPrecuentaHandler.cs` + `src/Inforest.Desktop/Impresion/FrmPrecuentaImpresora.cs` |
| `TIMPRESORA` | Table | `IImpresoraRepository` / `ImpresoraRepository` | Repository | IN_PROGRESS | `src/Inforest.Infrastructure/Impresion/ImpresoraRepository.cs` |
| `frmCambiarContrasenia.frm` | Form | `CambiarPasswordHandler` + `FrmCambiarContrasenia` | Handler + WinForm | IN_PROGRESS | `src/Inforest.Application/Seguridad/CambiarPasswordHandler.cs` + `src/Inforest.Desktop/POS/FrmCambiarContrasenia.cs` |
| `frmPassword.frm` | Form | `FrmPassword` | WinForm | IN_PROGRESS | `src/Inforest.Desktop/POS/FrmPassword.cs` |
| `AuthService` (cambio password) | Service | `IAuthService.CambiarPasswordAsync` | Service method | IN_PROGRESS | `src/Inforest.Infrastructure/Security/AuthService.cs` |
| `frmLiquidacionDetalle.frm` | Form | `FrmLiquidacionCierre` + `CerrarTurnoHandler` | WinForm + Handler | MIGRATED | `src/Inforest.Desktop/Caja/FrmLiquidacionCierre.cs` + `src/Inforest.Application/Turno/TurnoHandlers.cs` |
| `CierreTurnoBreakdown` (domain) | — | `CierreTurnoBreakdown` | Value Object | MIGRATED | `src/Inforest.Domain/Entities/Caja/CierreTurnoBreakdown.cs` |
| `UPDATE MTURNO (cierre completo)` | SQL | `TurnoRepository.CerrarAsync` | Repository | MIGRATED | `src/Inforest.Infrastructure/Turno/TurnoRepository.cs` |
| `TCAJA.lObligaCierre` | Flag | `CerrarTurnoHandler` (BR-CAJA-001) | Business Rule | MIGRATED | `src/Inforest.Application/Turno/TurnoHandlers.cs` |
| `TPARAMETRO.lActivaConsultaDescargo` | Flag | `CerrarTurnoHandler` (BR-CAJA-002) | Business Rule | MIGRATED | `src/Inforest.Application/Turno/TurnoHandlers.cs` |
