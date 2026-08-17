# Punto de Venta — Functionality Inventory

| ID | Funcionalidad | Evidencia Legacy | Descripción | Estado |
|---|---|---|---|---|
| POS-FUNC-001 | Inicio de sesión POS | `frmAcceso.frm + modPuntoVenta.bas` | Autenticar usuario y habilitar MDI POS | MIGRATED |
| POS-FUNC-002 | Apertura MDI Punto de Venta | `mdiPuntoVenta.frm` | Contenedor principal de operaciones POS | MIGRATED |
| POS-FUNC-003 | Gestión de pedidos salón | `frmPedido.frm / frmDetallePedido.frm / frmJuntarMesas.frm` | Crear/editar/agrupar pedidos | MIGRATED |
| POS-FUNC-004 | Registro de venta | `frmVenta.frm / frmSolicitudBoleta.frm` | Emitir venta desde pedido/canal y resolver boleta con/sin datos | MIGRATED |
| POS-FUNC-005 | Cobro y pagos | `frmPago.frm / frmPagoMasivo.frm / frmPagoPinPad.frm` | Registrar pagos por distintos medios | MIGRATED |
| POS-FUNC-006 | Facturación y documentos | `frmDocumento.frm / frmFactura.frm / frmNotaCredito*.frm` | Emitir comprobantes y NC | ANALYSIS |
| POS-FUNC-007 | Correlativos de comprobante | `frmDocumentoCorrelativo*.frm / modPuntoVenta.bas` | Control numeración por caja/tipo | COMPLETED |
| POS-FUNC-008 | Caja y cierre | `frmCajaRapida.frm / frmLiquidacionDetalle.frm` | Control de caja y liquidaciones | MIGRATED |
| POS-FUNC-009 | Turno/Día contable | `frmDiaContable.frm / clsDiaContable.cls` | Apertura/cierre día operativo | IN_PROGRESS |
| POS-FUNC-010 | Cliente y cuentas corrientes | `frmNuevoCliente.frm / frmCtaCte*.frm / frmCuentaCobrar*.frm` | Gestión de clientes y deuda | MIGRATED |
| POS-FUNC-011 | Reservas | `frmReserva.frm / frmReservaDetalle.frm` | Registro y seguimiento de reservas | MIGRATED |
| POS-FUNC-012 | Delivery dependiente POS | `frmPedidoDelivery*.frm / frmCentralPedidos.frm` | Pedidos delivery vinculados al POS | IN_PROGRESS |
| POS-FUNC-013 | Insumos/descargo | `frmInsumo*.frm / frmInsumoDetalle.frm / clsAlmacen.cls` | Ajustes/descargo de insumos | MIGRATED |
| POS-FUNC-014 | Importación de pedidos externos | `frmImportacionRequerimientos*.frm` | Integración pedidos externos | MIGRATED |
| POS-FUNC-015 | Mensajería cocina/KDS | `frmMensajeCocina*.frm / modKDS.bas` | Envío y control de producción | COMPLETED |
| POS-FUNC-016 | Impresión pre-cuenta/comanda | `frmPrecuentaImpresora.frm` | Salida térmica operativa | COMPLETED |
| POS-FUNC-017 | Reportería operativa POS | `frmRep*.frm + Designers DSR` | Consultas y reportes de operación | IN_PROGRESS |
| POS-FUNC-018 | Recibo de ingresos/egresos | `frmReciboIngreso*.frm / frmReciboEgreso*.frm` | Movimientos de caja adicionales | MIGRATED |
| POS-FUNC-019 | Control de seguridad y contraseñas | `frmPassword.frm / frmCambiarContrasenia.frm` | Validaciones sensibles de operación | COMPLETED |
| POS-FUNC-020 | Biometría huella | `frmVerificacionHuella*.frm + OCX` | Autenticación biométrica opcional | BLOCKED |
| POS-FUNC-021 | Tarjetas RFID / proximidad | `FrmTarjetaAproximidad*.frm / FrmRecargarTarjeta*.frm` | Saldo y consumo con tarjetas | IN_PROGRESS |
| POS-FUNC-022 | Balanza y captura de peso | `frmMsComm.frm / frmWinSock.frm / vBalanza` | Captura de productos por peso | ANALYSIS |
| POS-FUNC-023 | Actualizador cliente POS | `FrmActualizacion.frm / frmNoticias.frm` | Actualización y mensajes del cliente POS | ANALYSIS |
| POS-FUNC-024 | Fiscal Epson (Argentina) | `modImpresoraFiscal.bas + IFEpson.ocx` | Comprobación fiscal hardware | EXCLUDED_BY_SCOPE |
| POS-FUNC-030 | Mapa Visual de Mesas | `frmMesas.frm` | Mapa interactivo de salones y mesas con color coding por estado; selector de mesa para nuevo pedido | MIGRATED |
| POS-FUNC-031 | Teclado Numérico + Cambio de Propina | `frmNumPad.frm / frmCambioPropina.frm` | Teclado numérico modal reutilizable + diálogo de cambio de propina MN/ME | MIGRATED |
| POS-FUNC-032 | Mapa Consulta Mesas | `frmMesaConsulta.frm` | Mapa visual mesas en 3 modos: Visual (cambiar estado), Seleccionar (asignar mesa), Mover (transferir pedido) | MIGRATED |
| POS-FUNC-033 | Junta de Mesas — TPEDIDOMESA | `frmJuntaMesa.frm` | Gestión de mesas físicas asignadas a un pedido; tabla TPEDIDOMESA con transacción Delete+Insert | MIGRATED |
| POS-FUNC-034 | Correlativo de Pedidos | `frmPedidoCorrelativo.frm + frmPedidoDetalle.frm + frmPedidoAnterior.frm` | Consulta de pedidos por rango de fechas con drill-down a ítems (vPedidoDetalle) y documentos (vDocumentoAgrupado) en paneles integrados | MIGRATED |
