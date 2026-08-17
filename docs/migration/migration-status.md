# Estado de Migración — INFOREST

> Última actualización: 2026-08-17
>
> Estado general: **Fase 4 IN_PROGRESS — W1-W15 completados; POS-FUNC-001..029 parcialmente migrados; baseline transversal + configuración + maestros + turno + pedidos + venta + caja + reportes + módulos POS + servicios de dominio + POS-FUNC-004 (frmSolicitudBoleta.frm MIGRATED) + POS-FUNC-021 (RFID recargas base MIGRATED: `FrmRecargarTarjeta` + `FrmRecargarTarjetaDetalle` + `ObtenerRecargasTarjetaHandler` + `RegistrarRecargaTarjetaHandler` + `RecargaTarjetaRepository` + BR-RFID-005..007 + 3 tests; gap abierto para equivalencia total anticipo/documento) + POS-FUNC-027 (frmCambio.frm MIGRATED) + POS-FUNC-028 (frmUpdateDatosPedido.frm MIGRATED) + POS-FUNC-029 (frmListaMotivos.frm MIGRATED: MotivoAnulacion + IMotivoAnulacionRepository + MotivoAnulacionRepository + ObtenerMotivosAnulacionHandler + FrmListaMotivos + BR-MOTIVO-001..003 + 5 tests) + POS-FUNC-019 (frmCambiarContrasenia.frm + frmPassword.frm COMPLETED) — **570 tests en verde****

---

## Resumen Ejecutivo

| Indicador | Valor |
|---|---|
| Fase actual | 4 (IN_PROGRESS) — W1-W15 completados: Configuración, Maestros, Turno/DíaContable, Pedidos, Venta/Documentos, Caja/Pagos, SEC-006, Reportes FastReport, Módulos POS/Admin/Consultas/CajaRápida/Adición, TaxPolicy/ProductoVisibilidad/InventoryGateway/Email, Caja/Cierre (POS-FUNC-008 COMPLETED), POS-FUNC-006 (Notas de Crédito MIGRATED), POS-FUNC-010 (Cliente y cuentas corrientes MIGRATED) y POS-FUNC-021 (RFID recargas base MIGRATED, equivalencia total IN_PROGRESS) |
| Código .NET 8 existente | IN_PROGRESS — 250+ archivos .cs, 570 tests verdes: baseline + seguridad + KDS + delivery + configuración + maestros + turno + pedidos + venta + caja + reportes FastReport + WinForms POS/Admin/Consultas/CajaRápida/Adición + TaxPolicy + ProductoVisibilidad + InventoryGateway + SmtpEmail + TarjetaProximidad (`FrmTarjetaProximidad`, 14 tests) + RecargaTarjeta (`FrmRecargarTarjeta` + `FrmRecargarTarjetaDetalle` + 3 tests) + RegistrarPagosMultiplesHandler + DocumentoTests (13) + NotaCredito (19 tests) + CuentaCorrienteHandlerTests (7) + CuentaCorrienteTests (8) + CentralPedidosHandlersTests (10) + RequerimientoAlmacenTests (8) + ImportacionRequerimientoHandlerTests (9) + ReportesHandlersTests (22) + CambiarDocumentoHandlerTests (8) + ActualizarDatosPedidoHandlerTests (8) + MotivoAnulacionHandlerTests (5) + SolicitudBoletaHandlersTests (2) |
| Módulos migrados | 0 / 7 |
| Documentación Legacy | IN_PROGRESS — inventarios base completos; trazabilidad y gaps de Fase 3 actualizados al cierre P3-12 |
| Arquitectura Target definida | IN_PROGRESS — ADR-001 a ADR-012 aceptados; gaps funcionales posteriores siguen abiertos |

---

## Estado por Módulo

| Módulo | Ejecutable Legacy | Fase | Estado | Notas |
|---|---|---|---|---|
| Punto de Venta | `InfoRest.exe` | Implementación | IN_PROGRESS | POS core parcial en .NET 8: seguridad, pedidos, venta, caja, reportes y mantenimiento base RFID; faltan recargas RFID y otros gaps de hardware |
| Caja Rápida | `CajaRapida.exe` | Análisis | NOT_STARTED | Comparte gran parte del stack POS/caja |
| Adición | `Adicion.exe` | Análisis | NOT_STARTED | Flujo auxiliar sobre pedidos en curso |
| Administración | `Administracion.exe` | Análisis | NOT_STARTED | Maestros, parámetros y catálogos críticos |
| Consultas | `Consulta.exe` | Análisis | NOT_STARTED | Alta concentración de reportes `spRep_*` + Crystal |
| Despachador | `Despachador.exe` | Implementación | IN_PROGRESS | Domain+Application+Infrastructure+WinForms completo; plantillas FastReport implementadas; Rappi HTTP real pendiente |
| Motorizados | `Motorizado.exe` | Implementación | IN_PROGRESS | Domain+Application+Infrastructure+WinForms completo; plantillas FastReport implementadas |

---

## Estado por Capa

| Capa | Legacy | .NET 8 | Estado |
|---|---|---|---|
| Presentación (UI) | 400 formularios VB6 detectados | `Form1` + formularios base de delivery/motorizado/reportes | IN_PROGRESS |
| Lógica de negocio | 32 módulos + 10 clases | Seguridad, KDS/comanda, delivery, reportes y periféricos encapsulados por slices | IN_PROGRESS |
| Acceso a datos | ADO + `clsComando` + 150 SP | `IDbConnectionFactory` + `ISpExecutor` + `UnitOfWork` | IN_PROGRESS |
| Base de datos | 126T + 116V + 150SP | contrato SQL Legacy reutilizable + repositorios/reportes selectivos | IN_PROGRESS |
| Reportes | 206 Crystal Reports | 13 SPs + handlers + formularios base + plantillas pendientes | IN_PROGRESS |
| Integraciones | COM/DLL/OCX + hardware POS | KDS XML/directorio + PinPad/CashDro/BlueVision + FE por país (stubs controlados) | IN_PROGRESS |
| Seguridad/Auth | `INFSEGURIDAD` + cifrado débil Legacy | `AuthService` + `RbacService` + `SessionService` + `AuditoriaService` + `LicenseService` | IN_PROGRESS |
| Configuración | INI + `TPARAMETRO` + `TCAJA` | lectura tipada de `TPARAMETRO`/`TCAJA`, cache lazy y validación de arranque POS cubiertas por tests; edición administrativa sigue fuera de este corte | COMPLETED |

---

## Cobertura de Migración

> Cobertura todavía parcial: Fase 3 cierra la base transversal y vertical slices de soporte, pero aún no existe un módulo funcional completo listo para producción.

| Componente | Inventario Legacy | Migrado | Cobertura |
|---|---|---|---|
| Formularios | 400 | 0 | NOT_AVAILABLE |
| Módulos BAS | 32 | 0 | NOT_AVAILABLE |
| Clases | 10 | 0 | NOT_AVAILABLE |
| Stored Procedures | 150 | 0 | NOT_AVAILABLE |
| Tablas | 126 | 0 | NOT_AVAILABLE |
| Vistas | 116 | 0 | NOT_AVAILABLE |
| Reportes | 206 | 0 | NOT_AVAILABLE |
| Integraciones externas | 10+ | 0 | NOT_AVAILABLE |

---

## Milestones

| Milestone | Descripción | Estado |
|---|---|---|
| M1 | Documentación Legacy completa | COMPLETED |
| M2 | Arquitectura Target definida | IN_PROGRESS |
| M3 | Proyecto .NET 8 base creado | COMPLETED |
| M3-P3-01 | Re-baseline arquitectura ejecutable (P3-01) | COMPLETED — 2026-08-11 |
| M3-P3-02 | Contrato SQL y ejecución de datos Legacy (P3-02) | COMPLETED — 2026-08-11 |
| M3-P3-03 | Seguridad, sesión, auditoría y licenciamiento | COMPLETED — 2026-08-12 |
| M3-P3-08 | Cocina, comanda, áreas y KDS | COMPLETED — 2026-08-12 |
| M3-P3-10 | Reportes, consultas y salidas operativas | COMPLETED — 2026-08-12 |
| M3-P3-11 | Periféricos especializados e integraciones por país | COMPLETED — 2026-08-12 |
| M3-P3-12 | Validación integral, hardening y cierre de Fase 3 | COMPLETED — 2026-08-12 |
| M4 | Primer módulo migrado (Maestros) | IN_PROGRESS — W2 maestros implementados |
| M4-W1 | Configuración (`TPARAMETRO`/`TCAJA`): `ConfiguracionSistema` + `ConfiguracionCaja` + `ParametroRepository` + `ParametroService` + `ValidarInicioPosHandler` validados con tests | COMPLETED — 2026-08-14 |
| M4-W14 | Servicios de dominio transversales (TaxPolicy, ProductoVisibilidad, InventoryGateway, Email) | COMPLETED — 2026-08-13 |
| M4-W15 | Caja/Cierre completo (POS-FUNC-008): CierreTurnoBreakdown + CerrarTurnoHandler (BR-CAJA-001/002) + TurnoRepository full MTURNO update + FrmLiquidacionCierre + confirmación previa al cierre | COMPLETED — 2026-08-14 |
| M4-POS-FUNC-007 | Correlativos de comprobante (CorrelativoDocumento + CorrelativoHandlers + CorrelativoRepository + FrmDocumentoCorrelativo) | COMPLETED — 2026-08-13 |
| M4-POS-FUNC-016 | Impresión pre-cuenta (ImprimirPrecuentaHandler + ImpresoraRepository + FrmPrecuentaImpresora) | COMPLETED — 2026-08-13 |
| M4-POS-FUNC-019 | Control seguridad/contraseñas (CambiarPasswordHandler + AuthService.CambiarPasswordAsync + FrmCambiarContrasenia + FrmPassword) | COMPLETED — 2026-08-13 |
| M4-POS-FUNC-003 | Gestión de pedidos salón — stage MIGRATED: UpdatePedidoHandlerTests (6) + ObtenerPedidoHandlerTests (6) + BR-PEDIDO-001/002/003/004 documentados. Pendiente: VALIDATING vs Legacy + frmDetallePedido + frmJuntarMesas | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-004 | Registro de venta MIGRATED: EmitirDocumentoHandler + AnularDocumentoHandler + `frmSolicitudBoleta.frm` (`FrmSolicitudBoleta` + `FrmSeleccionClienteBoleta`) + `ObtenerClienteGeneralBoletaHandler` + validación SP `usp_Inforest_ValidaClienteSel` + `ConfiguracionSistema.lBODato` + Documento domain entity + 13 DocumentoTests + 6 `EmitirDocumentoHandlerTests` + 3 `AnularDocumentoHandlerTests` + 2 `SolicitudBoletaHandlersTests` + BR-VENTA-001/002/003 + BR-BOLETA-001/002/003 documentados | MIGRATED — 2026-08-17 |
| M4-POS-FUNC-006 | Facturación y Notas de Crédito (POS-FUNC-006): NotaCredito entity + EmitirNotaCreditoHandler + AnularNotaCreditoHandler + ObtenerNotaCreditoHandler + ObtenerNotasCreditoPorFechaHandler + NotaCreditoRepository + FrmNotaCredito + FrmNotaCreditoDetalle + BR-NC-001..006 documentados + 19 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-010 | Cliente y cuentas corrientes (POS-FUNC-010): CuentaCorriente entity + ICuentaCorrienteRepository + CuentaCorrienteRepository + FrmNuevoCliente + FrmCtaCte + FrmCuentaCobrar + BR-CLIENTE-001..004 + BR-CTACTE-001..003 + 7+8 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-011 | Reservas (POS-FUNC-011): Reserva entity + EstadoReserva enum + IReservaRepository + ReservaRepository + FrmReserva + FrmReservaDetalle + CrearReservaHandler + ModificarReservaHandler + AnularReservaHandler + ObtenerReservaHandler + ObtenerReservasPorFechaHandler + ConvertirReservaAPedidoHandler (spIns_MPEDIDO_RESERVA) + BR-RESERVA-001..004 + 18 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-012 | Delivery dependiente POS (POS-FUNC-012): FrmPedidoDelivery (seguimiento vDespachador) + CentralPedidosForm extendido (ConfirmarEntrega + RevertirEntrega + ModificarFecha) + ConfirmarEntregaCentralHandler + RevertirEntregaCentralHandler + ModificarFechaProgramadaDeliveryHandler + ObtenerPedidosSeguimientoDeliveryHandler + 5 nuevos métodos IPedidoDeliveryRepository + BR-DEL-012/013/014 + 10 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-013 | Insumos/descargo (POS-FUNC-013): Insumo entity + IInsumoRepository + InsumoRepository (USP_LISTARINSUMOS/usp_agregarinsumos/USP_MODIFICARINSUMOS/USP_ELIMINARINSUMOS) + FrmInsumo + FrmInsumoDetalle + ListarInsumosHandler + AgregarInsumoHandler + ModificarInsumoHandler + EliminarInsumoHandler + BR-INSUMO-001..004 + 18 tests (8 domain + 10 handler) | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-014 | Importación de pedidos externos (POS-FUNC-014): RequerimientoAlmacen + DetalleRequerimientoAlmacen (Domain) + IRequerimientoAlmacenRepository + IImportacionPedidoGateway (Application) + RequerimientoAlmacenRepository + ImportacionPedidoGateway (Infrastructure, ALMACEN DB) + ObtenerRequerimientosPendientesHandler + ObtenerDetalleRequerimientoHandler + ImportarRequerimientoHandler + FrmImportacionRequerimientos + FrmImportacionRequerimientoDetalle + BR-IMPORT-001..004 + 17 tests (8 domain + 9 handler) — 411 tests en verde | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-018 | Recibo Ingresos/Egresos (POS-FUNC-018): ReciboIngreso + ReciboEgreso (Domain, BR-RECIBO-001..012) + IReciboIngresoRepository + IReciboEgresoRepository + ReciboIngresoRepository + ReciboEgresoRepository (direct SQL, MINGRESO/MEGRESO) + ObtenerIngresosHandler + RegistrarIngresoHandler + AnularIngresoHandler + ObtenerEgresosHandler + RegistrarEgresoHandler + AnularEgresoHandler + FrmReciboIngreso + FrmReciboIngresoDetalle + FrmReciboEgreso + FrmReciboEgresoDetalle + 33 tests (9+8 domain + 8+8 handler) | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-025 | División de Pedidos (POS-FUNC-025): SesionDivision + ItemDivision (Domain, in-memory reemplaza temp tables SQL) + IDivisionPedidoRepository + IniciarDivisionHandler + DisgregarItemHandler + CompartirPedidoHandler + MoverItemHandler + ConfirmarDivisionHandler + DivisionPedidoRepository (DPEDIDO/CPEDIDO/MPEDIDO + spUpdate_DPEDIDO_Ina) + FrmDivision + BR-DIV-001..005 + 22 tests (11 domain + 11 handler) — **521 tests en verde** | MIGRATED — 2026-08-17 |
| M4-POS-FUNC-026 | Detalle de Pedido + Juntar Mesas (POS-FUNC-026): ItemPedidoVista + ComboPedidoVista + DetallePedidoExtendidoResultado (Application DTOs) + ObtenerDetallePedidoExtendidoHandler + IPedidoReadRepository extendido (ObtenerDetalleExtendidoAsync + ObtenerCombosAsync) + PedidoRepository impl (DPEDIDO JOIN TPRODUCTO / CPEDIDO JOIN TPRODUCTO) + PedidoVistaRows (Dapper) + FrmDetallePedido (read-only, grid items + combos + totales) + FrmJuntarMesas (selector visual de mesas con color-coded buttons + JuntarMesasResult) + BR-PEDIDO-005/006 + BR-JUNTAR-001/002/003 + 8 tests (handler) — **529 tests en verde** | MIGRATED — 2026-08-17 |
| M4-POS-FUNC-027 | Cambio de tipo de documento (POS-FUNC-027): CambiarDocumentoCommand + CambiarDocumentoHandler + ICambioDocumentoRepository + CambioDocumentoRepository (MDOCUMENTO/DDOCUMENTO/DPEDIDO/DPAGODOCUMENTO/TTIPODOCUMENTOIMPRESORA transacción) + ObtenerTiposDocumentoCambioHandler + FrmCambioDocumento (botones tipo documento + selección cliente + motivo + consumo) + BR-CAMBIO-001..005 + 8 tests — **545 tests en verde** | MIGRATED — 2026-08-17 |
| M4-POS-FUNC-028 | Actualizar Datos Pedido (POS-FUNC-028): ActualizarCortesiaPedidoCommand + ActualizarCanalVentaPedidoCommand + ActualizarCortesiaPedidoHandler + ActualizarCanalVentaPedidoHandler + ObtenerDatosPedidoActualizarHandler + IActualizarDatosPedidoRepository + ActualizarDatosPedidoRepository (usp_ActualizarCabPeDoc @opcion='1'/'2' + vCortesia + vTipoPedido) + FrmActualizarDatosPedido + BR-ACTPED-001..003 + 8 tests — **545 tests en verde** | MIGRATED — 2026-08-17 |
| M4-POS-FUNC-029 | Lista Motivos Anulación (POS-FUNC-029): MotivoAnulacion (domain record) + IMotivoAnulacionRepository + ObtenerMotivosAnulacionHandler + MotivoAnulacionRepository (TTABLA/vMotivoAnulacion) + FrmListaMotivos (modal dialog con botones de motivos) + BR-MOTIVO-001..003 + 5 tests — **550 tests en verde** | MIGRATED — 2026-08-17 |
| M4-POS-FUNC-030 | Mapa Visual de Mesas (POS-FUNC-030): `FrmMesas` (selector de mesa por salón con color-coding Libre/Ocupada/Reservada/Sucia/Bloqueada/FueraDeServicio) + `EstadoMesa` enum extendido (Sucia/Bloqueada/FueraDeServicio) + fix `MesaRepository` CASE mapping + `ObtenerPedidosSinMesaHandler` + `PedidoSinMesaVista` DTO + `PedidoRepository.ObtenerActivosSinMesaAsync` (MPEDIDO activos sin mesa) + eventos `MesaOcupadaSeleccionada`/`JuntarMesasSolicitado` + timer auto-refresh 30 s + BR-MESAS-001..005 + 3 tests — **556 tests en verde** | MIGRATED — 2026-08-17 |
| M4-POS-FUNC-021 | Tarjetas RFID / Proximidad (POS-FUNC-021): `TarjetaProximidad` + `MovimientoTarjetaProximidad` (Domain) + `ITarjetaProximidadRepository` + `TarjetaProximidadRepository` (TTARJETASRFID/TMOVIMIENTOTARJETASRFID) + `ObtenerTarjetasProximidadHandler` + `ObtenerMovimientosTarjetaProximidadHandler` + `CrearTarjetaProximidadHandler` + `ActualizarTarjetaProximidadHandler` + `FrmTarjetaProximidad` + `FrmRecargarTarjeta` + `FrmRecargarTarjetaDetalle` + `ObtenerRecargasTarjetaHandler` + `RegistrarRecargaTarjetaHandler` + `IRecargaTarjetaRepository` + `RecargaTarjetaRepository` + BR-RFID-001..007 + 17 tests. Gap abierto: equivalencia total del bloque legacy de anticipo/emisión documental/impresión en recarga. | IN_PROGRESS — 2026-08-17 |
| M4-POS-FUNC-009 | Turno/Día Contable (POS-FUNC-009): FrmDiaContable + AperturarDiaContableHandler + CerrarDiaContableHandler + ObtenerDiaContableHandler + IDiaContableService extendido (CerrarDiaContableAsync + ObtenerFechaMaximaDiaContableAsync) + BR-DC-001/002/003/004 documentados + 10 tests | COMPLETED — 2026-08-14 |
| M4-POS-FUNC-017 | Reportería operativa POS (slices `frmRepCtaCte.frm` + `frmRepAnulado.frm` + `frmRepLiquidacionTicket.frm` + `frmRepPaloteoTicket.frm` + `frmRepDeliveryTicket.frm` + `frmRepReservas.frm` + `frmRepEntrega.frm` + `frmRepComanda.frm` + `frmRepPropina.frm` + `frmRepVentaFecha.frm` + `frmRepLiquidacion.frm` + **`frmRepRegistroVenta.frm`**): `spRep_CtaCteN` + `spRep_Anulacion` + `spRep_LiquidacionSuma` + `(query dinámica frmRepPaloteoTicket)` + `(query dinámica frmRepDeliveryTicket)` + `(query dinámica frmRepReservas)` + `spRep_Entregas` + `spRep_Comanda` + `spRep_Propina` + `spRep_VentaFecha` + `spRep_LiquidacionOutPut` + `spRep_Liquidacion` (tipos 1-5) + **`spRep_RegVenta` + `spRep_RegVentaSunat` + `spRep_RegVentaSunatAD` + `spRep_ComprobanteDetallado`** + handlers/repositorios correspondientes + `FrmRepCtaCteReporte` + `FrmRepAnuladoReporte` + `FrmRepLiquidacionTicketReporte` + `FrmRepPaloteoTicketReporte` + `FrmRepDeliveryTicketReporte` + `FrmRepReservasReporte` + `FrmRepEntregaReporte` + `FrmComandaReporte` + `FrmPropinaReporte` + `FrmRepVentaFechaReporte` + `FrmRepLiquidacionReporte` + **`FrmRepRegistroVentaReporte`** + 21 plantillas FastReport + 24 tests (BR-REP-001/002/013..022) — **248 application tests en verde** | IN_PROGRESS — 2026-08-16 |
| M5 | Punto de Venta migrado | IN_PROGRESS — W4+W5+W6+W11+POS-FUNC-002/004/007/016/019 implementados; POS-FUNC-003 MIGRATED (tests completos) |
| M6 | Caja y Pagos migrados | NOT_STARTED |
| M7 | Todos los módulos migrados | NOT_STARTED |
| M8 | Validación completa | NOT_STARTED |
| M9 | Go-live en producción | NOT_STARTED |

---

## Blockers Actuales

| Blocker | Descripción | Impacto |
|---|---|---|
| HardKey físico | La validación física del dongle sigue reemplazada por validación contractual/SQL; requiere decisión posterior de hardware/licencia | Parcial |
| Biometría SecuGen | OCX de 32 bits sin SDK .NET oficial; servicio Null mantenido como gap controlado | Alto |
| Impresora fiscal Epson | Integración real sigue bloqueada por OCX 32-bit propietario | Alto |
| Plantillas FastReport | Se implementó `frmRepCtaCte` extremo a extremo, pero siguen pendientes plantillas `.frx` para varios reportes restantes | Medio |
| Rappi HTTP real | `RappiOrderAdapter` continúa como contrato/stub hasta definir integración productiva | Medio |

---

## Próximos Pasos

1. **COMPLETED:** POS-FUNC-018 Recibo de Ingresos/Egresos validado (`frmReciboIngreso.frm`, `frmReciboIngresoDetalle.frm`, `frmReciboEgreso.frm`, `frmReciboEgresoDetalle.frm`, MINGRESO/MEGRESO, BR-RECIBO-001..012, 33 tests en verde, 221 application tests en verde, 154 domain tests en verde).
2. **COMPLETED:** `frmRepCtaCte.frm` migrado y validado (`spRep_CtaCteN`, filtros `vTipoCtaCte`/`vSubTipoCtaCte`, `FrmRepCtaCteReporte`, 3 plantillas FastReport, 2 tests nuevos).
3. **COMPLETED:** `frmRepAnulado.frm` migrado (POS-FUNC-017): `spRep_Anulacion` + `AnulacionRow`/`AnulacionParametros` + `ObtenerAnulacionAsync` + `ObtenerReporteAnulacionHandler` + `FrmRepAnuladoReporte` + BR-REP-014 + 2 tests nuevos — 231 application tests en verde.
4. **COMPLETED:** `frmRepLiquidacionTicket.frm` migrado y validado (`spRep_LiquidacionSuma` + `LiquidacionTicketRow`/`LiquidacionTicketParametros` + `ObtenerLiquidacionTicketAsync` + `ObtenerReporteLiquidacionTicketHandler` + `FrmRepLiquidacionTicketReporte` + `RepLiquidacionTicket.frx` + BR-REP-015 + 2 tests).
5. **COMPLETED:** `frmRepPaloteoTicket.frm` migrado (POS-FUNC-017): query dinámica Legacy (`MPEDIDO/DPEDIDO/CPEDIDO/MDOCUMENTO`) + `PaloteoTicketRow`/`PaloteoTicketParametros` + `ObtenerPaloteoTicketAsync` + `ObtenerReportePaloteoTicketHandler` + `FrmRepPaloteoTicketReporte` + `RepPaloteoTicket.frx` + BR-REP-016 + 2 tests nuevos — 235 application tests en verde.
6. **MIGRATED:** `frmRepDeliveryTicket.frm` migrado (POS-FUNC-017): query dinámica "Cierre de Cajeros Delivery" (`MDOCUMENTO/DPREPAGO/MPEDIDO/vMotorizado/vTipoPago/vMoneda/TTARJETACREDITO/vTipoCancelacion`) + `DeliveryTicketRow`/`DeliveryTicketParametros` + `ObtenerDeliveryTicketAsync` + `ObtenerReporteDeliveryTicketHandler` + `FrmRepDeliveryTicketReporte` + `RepDeliveryTicket.frx` + BR-REP-017 + 2 tests nuevos — **237 application tests en verde**.
7. **MIGRATED:** `frmRepReservas.frm` migrado (POS-FUNC-017): query dinámica sobre `TRESERVA + vEstadoReserva` + filtros fecha/hora + estados (Generado/Atendido/Anulado) + `ReservaReporteRow`/`ReservaReporteParametros`/`OrdenReserva` + `ObtenerReservasReporteAsync` + `ObtenerReporteReservasHandler` + `FrmRepReservasReporte` + `RepReservas.frx` + BR-REP-018 + 2 tests nuevos — **239 application tests en verde**.
8. **MIGRATED:** `frmRepEntrega.frm` migrado (POS-FUNC-017): `spRep_Entregas` + `EntregaRow`/`EntregaParametros`/`FormatoReporteEntrega` + `ObtenerEntregasAsync` + `ObtenerReporteEntregaHandler` + `FrmRepEntregaReporte` + `RepEntregaFormato1.frx` + `RepEntregaFormato2.frx` + `RepEntregaResumidoProd.frx` + BR-REP-019 + 2 tests nuevos — **241 application tests en verde**.
9. **COMPLETED:** `frmRepComanda.frm` migrado (POS-FUNC-017): `FrmComandaReporte` actualizado con filtros Mozo/Usuario + checkboxes "Todos" + time pickers (dd/MM/yyyy HH:mm) + RadioButtons Detallado/Resumido + `ConstruirCriterio` (equivalente a sCriterio Legacy) + `RepComanda.frx` (Resumido: columnas reales) + `RepComandaDetallado.frx` (nuevo, landscape, 11 columnas) + BR-REP-001 confirmado + 2 tests existentes — **241 application tests en verde**.
10. **COMPLETED:** `frmRepPropina.frm` migrado (POS-FUNC-017): `FrmPropinaReporte` reescrito con combos Mozo/Motorizado/TipoPedido + checkboxes "Todos" + time pickers (dd/MM/yyyy HH:mm) + RadioButtons Detallado/Resumido + `ConstruirCondicion` (equivalente a Sub `Genera` Legacy: tmozo/tmotorizado/ttipopedido) + `ObtenerReportePropinaQuery.EsDetallado` flag + `RepPropina.frx` (Detallado: 10 columnas, landscape) + `RepPropinaResumido.frx` (Resumido: 5 columnas, portrait) + BR-REP-002 confirmado + 1 test nuevo (Resumido) — **242 application tests en verde**.
11. **COMPLETED:** `frmRepVentaFecha.frm` migrado (POS-FUNC-017): `spRep_VentaFecha` (SQL dinámico con `@sPrecio`/`@sFecha`/`@criterio`/`@tipooper`) + `VentaFechaRow` + `VentaFechaParametros` + `TipoPrecioVentaFecha` + `SubGrupoItem` + `ObtenerVentaFechaAsync` + `ObtenerSubGruposAsync` (`vSubGrupo`) + `ObtenerReporteVentaFechaHandler` + `FrmRepVentaFechaReporte` (año/mes/hora corte/precio/preventa/filtro sub-grupos) + `RepVentaFecha.frx` + BR-REP-020 + 2 tests nuevos — **244 application tests en verde**.
12. **MIGRATED:** `frmRepLiquidacion.frm` migrado (POS-FUNC-017): `spRep_LiquidacionOutPut` (OUTPUT params) + `spRep_Liquidacion` tipos 1-5 (documentos/sumasGrupo/tarjetas/otrosTipos) + `spRep_LiquidacionSuma` (tiposPedido) + `LiquidacionParametros` + `LiquidacionOutput` + `LiquidacionRow` + `LiquidacionSumaGrupoRow` + `LiquidacionTarjetaRow` + `LiquidacionTipoPedidoRow` + `LiquidacionOtroTipoRow` + `LiquidacionResultado` + `ObtenerReporteLiquidacionHandler` + `FrmRepLiquidacionReporte` + `RepLiquidacion.frx` + BR-REP-021 + 2 tests nuevos — **246 application tests en verde**.
13. **MIGRATED:** `frmRepRegistroVenta.frm` migrado (POS-FUNC-017): `spRep_RegVenta` (tipos 1/2/4/6 con flags) + `spRep_RegVentaSunat` (tipo 0) + `spRep_RegVentaSunatAD` (tipo 3) + `spRep_ComprobanteDetallado` (tipo 5) + `TipoReporteRegistroVenta` enum + `RegistroVentaParametros` + `RegistroVentaRow` + `RegistroVentaSunatRow` + `RegistroVentaSunatAdRow` + `RegistroVentaDetalladoRow` + `RegistroVentaResultado` + `ObtenerReporteRegistroVentaHandler` + `FrmRepRegistroVentaReporte` + 5 plantillas FastReport + BR-REP-022 + GAP-REP-022 (`spRep_RegVentaSunat_formaPago` no encontrado) + 2 tests nuevos — **248 application tests en verde**. ⚠️ Todos los `frmRep*` de `InfoRest.vbp` han sido migrados.
14. **MIGRATED:** `frmDivision.frm` migrado (POS-FUNC-025): `SesionDivision` + `ItemDivision` (Domain, in-memory reemplaza temp tables SQL) + `IDivisionPedidoRepository` + `IniciarDivisionHandler` + `DisgregarItemHandler` + `CompartirPedidoHandler` + `MoverItemHandler` + `ConfirmarDivisionHandler` + `DivisionPedidoRepository` (DPEDIDO/CPEDIDO/MPEDIDO + `spUpdate_DPEDIDO_Ina`) + `FrmDivision` + BR-DIV-001..005 + 22 tests nuevos (11 domain + 11 handler) — **521 tests en verde**.
15. **MIGRATED:** `frmDetallePedido.frm` + `frmJuntarMesas.frm` migrados (POS-FUNC-026): `ItemPedidoVista` + `ComboPedidoVista` + `DetallePedidoExtendidoResultado` (Application DTOs) + `ObtenerDetallePedidoExtendidoHandler` + `IPedidoReadRepository` extendido (`ObtenerDetalleExtendidoAsync` + `ObtenerCombosAsync`) + `PedidoRepository` impl (DPEDIDO JOIN TPRODUCTO / CPEDIDO JOIN TPRODUCTO) + `PedidoVistaRows` (Dapper) + `FrmDetallePedido` (read-only, grid items + combos + totales) + `FrmJuntarMesas` (selector visual de mesas con color-coded buttons + `JuntarMesasResult`) + BR-PEDIDO-005/006 + BR-JUNTAR-001/002/003 + 8 tests nuevos (handler) — **529 tests en verde**.
16. **MIGRATED:** `frmCambio.frm` migrado (POS-FUNC-027): `CambiarDocumentoCommand` + `CambiarDocumentoHandler` + `ICambioDocumentoRepository` + `CambioDocumentoRepository` (transacción MDOCUMENTO/DDOCUMENTO/DPEDIDO/DPAGODOCUMENTO/TTIPODOCUMENTOIMPRESORA) + `ObtenerTiposDocumentoCambioHandler` + `FrmCambioDocumento` (botones tipo doc + selección cliente + motivo + consumo) + BR-CAMBIO-001..005 + 8 tests nuevos — **545 tests en verde**.
17. **MIGRATED:** `frmUpdateDatosPedido.frm` migrado (POS-FUNC-028): `ActualizarCortesiaPedidoCommand` + `ActualizarCanalVentaPedidoCommand` + `ActualizarCortesiaPedidoHandler` + `ActualizarCanalVentaPedidoHandler` + `ObtenerDatosPedidoActualizarHandler` + `IActualizarDatosPedidoRepository` + `ActualizarDatosPedidoRepository` (usp_ActualizarCabPeDoc @opcion='1'/'2') + `FrmActualizarDatosPedido` (GroupBox cortesía + GroupBox canal de venta + cliente delivery) + BR-ACTPED-001..003 + 8 tests nuevos — **545 tests en verde**.
18. **MIGRATED:** `frmListaMotivos.frm` migrado (POS-FUNC-029): `MotivoAnulacion` (domain record) + `IMotivoAnulacionRepository` + `ObtenerMotivosAnulacionHandler` + `MotivoAnulacionRepository` (TTABLA/vMotivoAnulacion — filtra lActivo=1, ordena por Codigo) + `FrmListaMotivos` (modal dialog con botones per-motivo, selección devuelve `MotivoSeleccionado`) + BR-MOTIVO-001..003 + 5 tests nuevos — **550 tests en verde**.
19. **COMPLETED:** `frmCambiarContrasenia.frm` + `frmPassword.frm` validados (POS-FUNC-019): `CambiarPasswordHandler` + `AuthService.CambiarPasswordAsync` alineados a Legacy (`TUSUARIO` activo único por `tResumido`, validaciones de cambio, actualización `tPassword/fRegistro/tUsuarioModifica` y sidecar `TUSUARIO_HASH`) + normalización de login en UI + 8 tests existentes en verde (`CambiarPasswordHandlerTests`).
20. **MIGRATED (base):** `FrmRecargarTarjeta.frm` + `FrmRecargarTarjetaDetalle.frm` (POS-FUNC-021) migrados con listado por `usp_Inforest_ObtieneRecargas`, registro transaccional de movimiento `R` en `TMOVIMIENTOTARJETASRFID` y actualización de `TTARJETASRFID.MontoDisponible` + BR-RFID-005..007 + 3 tests nuevos — **570 tests en verde**.
21. **SIGUIENTE PENDIENTE:** Cerrar equivalencia completa del bloque legacy de recarga (anticipo/emisión documental/impresión fiscal) dentro de POS-FUNC-021.
22. Mantener seguimiento de gaps bloqueantes: HardKey físico, SecuGen, Epson fiscal y conectores FE/Rappi reales.
23. Completar plantillas FastReport pendientes para los handlers/reportes de Etapa 10 restantes.

---

## Referencias

- [Estrategia de migración](migration-strategy.md)
- [Inventario Legacy](legacy-inventory.md)
- [Inventario SQL](database/sql-inventory.md)
- [Matriz de lógica de negocio](traceability/business-logic-matrix.md)
- [Decisiones arquitectónicas](../architecture/architecture-decisions.md)
