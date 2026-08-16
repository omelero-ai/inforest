# Estado de Migración — INFOREST

> Última actualización: 2026-08-16
>
> Estado general: **Fase 4 IN_PROGRESS — W1-W15 completados; POS-FUNC-001 (Login) y POS-FUNC-002 (Apertura MDI) MIGRATED; baseline transversal + configuración + maestros + turno + pedidos + venta + caja + reportes + módulos POS + servicios de dominio W14 (TaxPolicy, ProductoVisibilidad, InventoryGateway, Email) + POS-FUNC-007/016/019 + POS-FUNC-008 (Caja/Cierre COMPLETED) + POS-FUNC-003 (Gestión pedidos salón MIGRATED) + POS-FUNC-004 (Registro de venta MIGRATED) + POS-FUNC-005 (Cobro y pagos MIGRATED) + POS-FUNC-006 (Notas de Crédito MIGRATED) + POS-FUNC-010 (Cliente y cuentas corrientes MIGRATED) + POS-FUNC-011 (Reservas MIGRATED) + POS-FUNC-012 (Delivery dependiente POS MIGRATED) + POS-FUNC-013 (Insumos/descargo MIGRATED) + POS-FUNC-014 (Importación de pedidos externos MIGRATED) + **POS-FUNC-017 IN_PROGRESS (frmRepCtaCte COMPLETED + frmRepAnulado COMPLETED + frmRepLiquidacionTicket COMPLETED + frmRepPaloteoTicket COMPLETED + frmRepDeliveryTicket COMPLETED + frmRepReservas MIGRATED: ReservaReporteRow + ReservaReporteParametros + OrdenReserva + ObtenerReservasReporteAsync + ObtenerReporteReservasHandler + FrmRepReservasReporte + RepReservas.frx + BR-REP-018 + 2 tests — 239 application tests en verde)**

---

## Resumen Ejecutivo

| Indicador | Valor |
|---|---|
| Fase actual | 4 (IN_PROGRESS) — W1-W15 completados: Configuración, Maestros, Turno/DíaContable, Pedidos, Venta/Documentos, Caja/Pagos, SEC-006, Reportes FastReport, Módulos POS/Admin/Consultas/CajaRápida/Adición, TaxPolicy/ProductoVisibilidad/InventoryGateway/Email, Caja/Cierre(POS-FUNC-008 COMPLETED) + **POS-FUNC-006 (Notas de Crédito MIGRATED) + **POS-FUNC-010 (Cliente y cuentas corrientes MIGRATED)** |
| Código .NET 8 existente | IN_PROGRESS — 240+ archivos .cs, 477 tests verdes: baseline + seguridad + KDS + delivery + configuración + maestros + turno + pedidos + venta + caja + reportes FastReport + WinForms POS/Admin/Consultas/CajaRápida/Adición + TaxPolicy + ProductoVisibilidad + InventoryGateway + SmtpEmail + RegistrarPagosMultiplesHandler + DocumentoTests (13) + NotaCredito (19 tests) + CuentaCorrienteHandlerTests (7) + CuentaCorrienteTests (8) + CentralPedidosHandlersTests (10) + RequerimientoAlmacenTests (8) + ImportacionRequerimientoHandlerTests (9) + ReportesHandlersTests (18) |
| Módulos migrados | 0 / 7 |
| Documentación Legacy | IN_PROGRESS — inventarios base completos; trazabilidad y gaps de Fase 3 actualizados al cierre P3-12 |
| Arquitectura Target definida | IN_PROGRESS — ADR-001 a ADR-012 aceptados; gaps funcionales posteriores siguen abiertos |

---

## Estado por Módulo

| Módulo | Ejecutable Legacy | Fase | Estado | Notas |
|---|---|---|---|---|
| Punto de Venta | `InfoRest.exe` | Análisis | NOT_STARTED | Baseline funcional identificado; depende de Infrastructure + Configuración + Turno |
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
| M4-POS-FUNC-004 | Registro de venta MIGRATED: EmitirDocumentoHandler + AnularDocumentoHandler + Documento domain entity + 13 DocumentoTests (PE/CO/AN estados, total fórmula, propina, descuento, validaciones) + 5+3 VentaHandlerTests + BR-VENTA-001/002/003 documentados | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-006 | Facturación y Notas de Crédito (POS-FUNC-006): NotaCredito entity + EmitirNotaCreditoHandler + AnularNotaCreditoHandler + ObtenerNotaCreditoHandler + ObtenerNotasCreditoPorFechaHandler + NotaCreditoRepository + FrmNotaCredito + FrmNotaCreditoDetalle + BR-NC-001..006 documentados + 19 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-010 | Cliente y cuentas corrientes (POS-FUNC-010): CuentaCorriente entity + ICuentaCorrienteRepository + CuentaCorrienteRepository + FrmNuevoCliente + FrmCtaCte + FrmCuentaCobrar + BR-CLIENTE-001..004 + BR-CTACTE-001..003 + 7+8 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-011 | Reservas (POS-FUNC-011): Reserva entity + EstadoReserva enum + IReservaRepository + ReservaRepository + FrmReserva + FrmReservaDetalle + CrearReservaHandler + ModificarReservaHandler + AnularReservaHandler + ObtenerReservaHandler + ObtenerReservasPorFechaHandler + ConvertirReservaAPedidoHandler (spIns_MPEDIDO_RESERVA) + BR-RESERVA-001..004 + 18 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-012 | Delivery dependiente POS (POS-FUNC-012): FrmPedidoDelivery (seguimiento vDespachador) + CentralPedidosForm extendido (ConfirmarEntrega + RevertirEntrega + ModificarFecha) + ConfirmarEntregaCentralHandler + RevertirEntregaCentralHandler + ModificarFechaProgramadaDeliveryHandler + ObtenerPedidosSeguimientoDeliveryHandler + 5 nuevos métodos IPedidoDeliveryRepository + BR-DEL-012/013/014 + 10 tests | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-013 | Insumos/descargo (POS-FUNC-013): Insumo entity + IInsumoRepository + InsumoRepository (USP_LISTARINSUMOS/usp_agregarinsumos/USP_MODIFICARINSUMOS/USP_ELIMINARINSUMOS) + FrmInsumo + FrmInsumoDetalle + ListarInsumosHandler + AgregarInsumoHandler + ModificarInsumoHandler + EliminarInsumoHandler + BR-INSUMO-001..004 + 18 tests (8 domain + 10 handler) | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-014 | Importación de pedidos externos (POS-FUNC-014): RequerimientoAlmacen + DetalleRequerimientoAlmacen (Domain) + IRequerimientoAlmacenRepository + IImportacionPedidoGateway (Application) + RequerimientoAlmacenRepository + ImportacionPedidoGateway (Infrastructure, ALMACEN DB) + ObtenerRequerimientosPendientesHandler + ObtenerDetalleRequerimientoHandler + ImportarRequerimientoHandler + FrmImportacionRequerimientos + FrmImportacionRequerimientoDetalle + BR-IMPORT-001..004 + 17 tests (8 domain + 9 handler) — 411 tests en verde | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-018 | Recibo Ingresos/Egresos (POS-FUNC-018): ReciboIngreso + ReciboEgreso (Domain, BR-RECIBO-001..012) + IReciboIngresoRepository + IReciboEgresoRepository + ReciboIngresoRepository + ReciboEgresoRepository (direct SQL, MINGRESO/MEGRESO) + ObtenerIngresosHandler + RegistrarIngresoHandler + AnularIngresoHandler + ObtenerEgresosHandler + RegistrarEgresoHandler + AnularEgresoHandler + FrmReciboIngreso + FrmReciboIngresoDetalle + FrmReciboEgreso + FrmReciboEgresoDetalle + 33 tests (9+8 domain + 8+8 handler) | MIGRATED — 2026-08-14 |
| M4-POS-FUNC-009 | Turno/Día Contable (POS-FUNC-009): FrmDiaContable + AperturarDiaContableHandler + CerrarDiaContableHandler + ObtenerDiaContableHandler + IDiaContableService extendido (CerrarDiaContableAsync + ObtenerFechaMaximaDiaContableAsync) + BR-DC-001/002/003/004 documentados + 10 tests | COMPLETED — 2026-08-14 |
| M4-POS-FUNC-017 | Reportería operativa POS (slices `frmRepCtaCte.frm` + `frmRepAnulado.frm` + `frmRepLiquidacionTicket.frm` + `frmRepPaloteoTicket.frm` + `frmRepDeliveryTicket.frm` + `frmRepReservas.frm`): `spRep_CtaCteN` + `spRep_Anulacion` + `spRep_LiquidacionSuma` + `(query dinámica frmRepPaloteoTicket)` + `(query dinámica frmRepDeliveryTicket)` + `(query dinámica frmRepReservas: TRESERVA+vEstadoReserva)` + handlers/repositorios correspondientes + `FrmRepCtaCteReporte` + `FrmRepAnuladoReporte` + `FrmRepLiquidacionTicketReporte` + `FrmRepPaloteoTicketReporte` + `FrmRepDeliveryTicketReporte` + `FrmRepReservasReporte` + 6 plantillas FastReport + 12 tests nuevos (BR-REP-013..018) — 239 application tests en verde | IN_PROGRESS — 2026-08-16 |
| M5 | Punto de Venta migrado | IN_PROGRESS — W4+W5+W6+W11+POS-FUNC-002/007/016/019 implementados; POS-FUNC-003 MIGRATED (tests completos) |
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
| Build Desktop preexistente | `dotnet build src/Inforest.Desktop/Inforest.Desktop.csproj` sigue fallando por error previo en `Turno/FrmCierreTurno.cs` (`_codigoUsuario` inexistente), ajeno a este slice | Medio |

---

## Próximos Pasos

1. **COMPLETED:** POS-FUNC-018 Recibo de Ingresos/Egresos validado (`frmReciboIngreso.frm`, `frmReciboIngresoDetalle.frm`, `frmReciboEgreso.frm`, `frmReciboEgresoDetalle.frm`, MINGRESO/MEGRESO, BR-RECIBO-001..012, 33 tests en verde, 221 application tests en verde, 154 domain tests en verde).
2. **COMPLETED:** `frmRepCtaCte.frm` migrado y validado (`spRep_CtaCteN`, filtros `vTipoCtaCte`/`vSubTipoCtaCte`, `FrmRepCtaCteReporte`, 3 plantillas FastReport, 2 tests nuevos).
3. **COMPLETED:** `frmRepAnulado.frm` migrado (POS-FUNC-017): `spRep_Anulacion` + `AnulacionRow`/`AnulacionParametros` + `ObtenerAnulacionAsync` + `ObtenerReporteAnulacionHandler` + `FrmRepAnuladoReporte` + BR-REP-014 + 2 tests nuevos — 231 application tests en verde.
4. **COMPLETED:** `frmRepLiquidacionTicket.frm` migrado y validado (`spRep_LiquidacionSuma` + `LiquidacionTicketRow`/`LiquidacionTicketParametros` + `ObtenerLiquidacionTicketAsync` + `ObtenerReporteLiquidacionTicketHandler` + `FrmRepLiquidacionTicketReporte` + `RepLiquidacionTicket.frx` + BR-REP-015 + 2 tests).
5. **COMPLETED:** `frmRepPaloteoTicket.frm` migrado (POS-FUNC-017): query dinámica Legacy (`MPEDIDO/DPEDIDO/CPEDIDO/MDOCUMENTO`) + `PaloteoTicketRow`/`PaloteoTicketParametros` + `ObtenerPaloteoTicketAsync` + `ObtenerReportePaloteoTicketHandler` + `FrmRepPaloteoTicketReporte` + `RepPaloteoTicket.frx` + BR-REP-016 + 2 tests nuevos — 235 application tests en verde.
6. **MIGRATED:** `frmRepDeliveryTicket.frm` migrado (POS-FUNC-017): query dinámica "Cierre de Cajeros Delivery" (`MDOCUMENTO/DPREPAGO/MPEDIDO/vMotorizado/vTipoPago/vMoneda/TTARJETACREDITO/vTipoCancelacion`) + `DeliveryTicketRow`/`DeliveryTicketParametros` + `ObtenerDeliveryTicketAsync` + `ObtenerReporteDeliveryTicketHandler` + `FrmRepDeliveryTicketReporte` + `RepDeliveryTicket.frx` + BR-REP-017 + 2 tests nuevos — **237 application tests en verde**.
7. **MIGRATED:** `frmRepReservas.frm` migrado (POS-FUNC-017): query dinámica sobre `TRESERVA + vEstadoReserva` + filtros fecha/hora + estados (Generado/Atendido/Anulado) + `ReservaReporteRow`/`ReservaReporteParametros`/`OrdenReserva` + `ObtenerReservasReporteAsync` + `ObtenerReporteReservasHandler` + `FrmRepReservasReporte` + `RepReservas.frx` + BR-REP-018 + 2 tests nuevos — **239 application tests en verde**.
8. Siguiente pendiente directo de `InfoRest.vbp` en etapa POS-FUNC-017: `frmRepEntrega.frm` (usa `spRep_Entregas`).
9. Mantener seguimiento de gaps bloqueantes heredados de Fase 3: HardKey físico, SecuGen, Epson fiscal y conectores FE/Rappi reales.
9. Completar plantillas FastReport pendientes para que los handlers/reportes de Etapa 10 restantes puedan operar extremo a extremo.

---

## Referencias

- [Estrategia de migración](migration-strategy.md)
- [Inventario Legacy](legacy-inventory.md)
- [Inventario SQL](database/sql-inventory.md)
- [Matriz de lógica de negocio](traceability/business-logic-matrix.md)
- [Decisiones arquitectónicas](../architecture/architecture-decisions.md)
