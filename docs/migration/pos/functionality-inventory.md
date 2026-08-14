# Punto de Venta — Functionality Inventory

| ID | Funcionalidad | Evidencia Legacy | Descripción | Estado |
|---|---|---|---|---|
| POS-FUNC-001 | Inicio de sesión POS | `frmAcceso.frm + modPuntoVenta.bas` | Autenticar usuario y habilitar MDI POS | MIGRATED |
| POS-FUNC-002 | Apertura MDI Punto de Venta | `mdiPuntoVenta.frm` | Contenedor principal de operaciones POS | MIGRATED |
| POS-FUNC-003 | Gestión de pedidos salón | `frmPedido.frm / frmDetallePedido.frm / frmJuntarMesas.frm` | Crear/editar/agrupar pedidos | MIGRATED |
| POS-FUNC-004 | Registro de venta | `frmVenta.frm` | Emitir venta desde pedido/canal | MIGRATED |
| POS-FUNC-005 | Cobro y pagos | `frmPago.frm / frmPagoMasivo.frm / frmPagoPinPad.frm` | Registrar pagos por distintos medios | MIGRATED |
| POS-FUNC-006 | Facturación y documentos | `frmDocumento.frm / frmFactura.frm / frmNotaCredito*.frm` | Emitir comprobantes y NC | ANALYSIS |
| POS-FUNC-007 | Correlativos de comprobante | `frmDocumentoCorrelativo*.frm / modPuntoVenta.bas` | Control numeración por caja/tipo | COMPLETED |
| POS-FUNC-008 | Caja y cierre | `frmCajaRapida.frm / frmLiquidacionDetalle.frm` | Control de caja y liquidaciones | MIGRATED |
| POS-FUNC-009 | Turno/Día contable | `frmDiaContable.frm / clsDiaContable.cls` | Apertura/cierre día operativo | IN_PROGRESS |
| POS-FUNC-010 | Cliente y cuentas corrientes | `frmNuevoCliente.frm / frmCtaCte*.frm / frmCuentaCobrar*.frm` | Gestión de clientes y deuda | MIGRATED |
| POS-FUNC-011 | Reservas | `frmReserva.frm / frmReservaDetalle.frm` | Registro y seguimiento de reservas | MIGRATED |
| POS-FUNC-012 | Delivery dependiente POS | `frmPedidoDelivery*.frm / frmCentralPedidos.frm` | Pedidos delivery vinculados al POS | IN_PROGRESS |
| POS-FUNC-013 | Insumos/descargo | `frmInsumo*.frm / clsAlmacen.cls` | Ajustes/descargo de insumos | IN_PROGRESS |
| POS-FUNC-014 | Importación de pedidos externos | `frmImportacionRequerimientos*.frm` | Integración pedidos externos | ANALYSIS |
| POS-FUNC-015 | Mensajería cocina/KDS | `frmMensajeCocina*.frm / modKDS.bas` | Envío y control de producción | IN_PROGRESS |
| POS-FUNC-016 | Impresión pre-cuenta/comanda | `frmPrecuentaImpresora.frm / frmRepPaloteoTicket.frm` | Salida térmica operativa | ANALYSIS |
| POS-FUNC-017 | Reportería operativa POS | `frmRep*.frm + Designers DSR` | Consultas y reportes de operación | ANALYSIS |
| POS-FUNC-018 | Recibo de ingresos/egresos | `frmReciboIngreso*.frm / frmReciboEgreso*.frm` | Movimientos de caja adicionales | ANALYSIS |
| POS-FUNC-019 | Control de seguridad y contraseñas | `frmPassword.frm / frmCambiarContrasenia.frm` | Validaciones sensibles de operación | ANALYSIS |
| POS-FUNC-020 | Biometría huella | `frmVerificacionHuella*.frm + OCX` | Autenticación biométrica opcional | BLOCKED |
| POS-FUNC-021 | Tarjetas RFID / proximidad | `FrmTarjetaAproximidad*.frm / FrmRecargarTarjeta*.frm` | Saldo y consumo con tarjetas | ANALYSIS |
| POS-FUNC-022 | Balanza y captura de peso | `frmMsComm.frm / frmWinSock.frm / vBalanza` | Captura de productos por peso | ANALYSIS |
| POS-FUNC-023 | Actualizador cliente POS | `FrmActualizacion.frm / frmNoticias.frm` | Actualización y mensajes del cliente POS | ANALYSIS |
| POS-FUNC-024 | Fiscal Epson (Argentina) | `modImpresoraFiscal.bas + IFEpson.ocx` | Comprobación fiscal hardware | EXCLUDED_BY_SCOPE |