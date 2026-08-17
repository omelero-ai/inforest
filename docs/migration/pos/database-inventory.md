# Punto de Venta — Database Inventory (InfoRest.exe)

> SQL auditado desde componentes declarados en `InfoRest.vbp` (`.frm`, `.bas`, `.cls`).

## Resumen

| Tipo SQL | Total en scripts Legacy | Detectado en componentes POS | Estado |
|---|---:|---:|---|
| Stored Procedures | 150 (inventario global) | 80 | IN_PROGRESS |
| Views | 116 (inventario global) | 67 | IN_PROGRESS |
| Tables | 126 (inventario global) | 91 | IN_PROGRESS |
| SQL dinámico (líneas detectadas) | N/A | 5723 | IN_PROGRESS |

## Stored Procedures detectados (InfoRest)

| ID | Stored Procedure | Evidencia VB6 (muestras) | Clasificación | Estado |
|---:|---|---|---|---|
| 1 | `SP_Dictionary_Inforest` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 2 | `SP_EJECUTA_ACTUALIZA_FE` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 3 | `SP_INFOREST_VALPRODASOCIADOS` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmVenta.frm` | SQL_ONLY | ANALYSIS |
| 4 | `SP_InsertaDescargoInsumo` | `frmDescargo.frm` | SQL_ONLY | ANALYSIS |
| 5 | `SP_NOTACREDITO_PARCIAL` | `frmNotaCreditoDetalle.frm` | SQL_ONLY | ANALYSIS |
| 6 | `SP_NOTICIAS_INFOREST` | `mdiPuntoVenta.frm` | SQL_ONLY | ANALYSIS |
| 7 | `USP_ADD_MARCACION` | `frmVerificarHuellaMarcacion.frm` | SQL_ONLY | ANALYSIS |
| 8 | `USP_AGREGARINSUMOS` | `frmInsumoDetalle.frm` | SQL_ONLY | ANALYSIS |
| 9 | `USP_AGREGARMENSAJE` | `frmMensajeCocinaDetalle.frm` | SQL_ONLY | MIGRATED |
| 10 | `USP_CALCULA_PRECIO` | `frmDivision.frm` | SQL_ONLY | ANALYSIS |
| 11 | `USP_CERRAR_MENSAJES_CIERRETURNO` | `frmLiquidacionDetalle.frm` | SQL_ONLY | MIGRATED |
| 12 | `USP_ELIMINARINSUMOS` | `frmInsumoDetalle.frm` | SQL_ONLY | ANALYSIS |
| 13 | `USP_ELIMINARRMENSAJES` | `frmMensajeCocinaDetalle.frm` | SQL_ONLY | MIGRATED |
| 14 | `USP_INF_verificaProductoAreaPantalla` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 15 | `USP_KDS_GrabarTiempoSalidaDPedido` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 16 | `USP_KDS_ObtenerArea` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 17 | `USP_KDS_ObtenerAreaImpresionKDS` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 18 | `USP_KDS_ObtenerCategoria` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 19 | `USP_KDS_ObtenerDetalleCombo` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 20 | `USP_KDS_ObtenerDetallePedido` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 21 | `USP_KDS_ObtenerDetallePedidox` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 22 | `USP_KDS_ObtenerDetallePedidoxKDS` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 23 | `USP_KDS_ObtenerNombreMesaXCodigo` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 24 | `USP_KDS_ObtenerProductoArea` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 25 | `USP_KDS_ObtenerProductoPedido` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 26 | `USP_KDS_ObtenerProductoPedidoDeCombo` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 27 | `USP_KDS_ObtenerProductoPedidoImpresos` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 28 | `USP_KDS_ObtenerPropiedadesProducto` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 29 | `USP_KDS_ObtenerTipoPedido` | `modKDS.bas` | SQL_ONLY | ANALYSIS |
| 30 | `USP_LISTADOMENSAJES` | `frmMozoUsuario.frm` | SQL_ONLY | ANALYSIS |
| 31 | `USP_LISTARINSUMOS` | `frmInsumo.frm` | SQL_ONLY | ANALYSIS |
| 32 | `USP_LISTARMENSAJES` | `frmLiquidacionDetalle.frm`, `frmMensajeCocina.frm` | SQL_ONLY | MIGRATED |
| 33 | `USP_MODIFICARINSUMOS` | `frmInsumoDetalle.frm` | SQL_ONLY | ANALYSIS |
| 34 | `USP_MODIFICARMENSAJE` | `frmMensajeCocinaDetalle.frm` | SQL_ONLY | MIGRATED |
| 35 | `Usp_ActualizarPreciosTransferenciaAlmacen` | `frmAcceso.frm`, `frmFlash.frm` | SQL_ONLY | ANALYSIS |
| 36 | `Usp_GuiaTransporteXml` | `clsxml.cls` | SQL_ONLY | ANALYSIS |
| 37 | `spIns_DPEDIDO` | `frmCargoMozo.frm` | SQL_ONLY | ANALYSIS |
| 38 | `spIns_MPEDIDO` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | SQL_ONLY | ANALYSIS |
| 39 | `spIns_MPEDIDO_RESERVA` | `frmReservaDetalle.frm` | SQL_ONLY | ANALYSIS |
| 40 | `spIns_TipoCambio` | `frmAcceso.frm`, `frmFlash.frm`, `frmInicio.frm`… | SQL_ONLY | ANALYSIS |
| 41 | `spRep_CtaCteN` | `frmRepCtaCte.frm` | SQL_ONLY | MIGRATED |
| 42 | `spRep_Entregas` | `frmRepEntrega.frm` | SQL_ONLY | ANALYSIS |
| 43 | `spRep_LiquidacionSuma` | `frmRepLiquidacionTicket.frm` | SQL_ONLY | MIGRATED |
| 44 | `spRep_ReciboEgreso` | `frmReciboEgreso.frm` | SQL_ONLY | ANALYSIS |
| 45 | `spUpd_MPEDIDO` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmVenta.frm` | SQL_ONLY | ANALYSIS |
| 46 | `sp_AsignaAnticipo_Pedido` | `frmVenta.frm` | SQL_ONLY | ANALYSIS |
| 47 | `sp_VinculacionSAP` | `frmInsumo.frm` | SQL_ONLY | ANALYSIS |
| 48 | `ups_ObtieneFechaHora` | `clsDiaContable.cls`, `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 49 | `usp_ActualizarCabPeDoc` | `frmUpdateDatosPedido.frm` | SQL_ONLY | ANALYSIS |
| 50 | `usp_ControlServicioWindows` | `clsAlmacen.cls` | SQL_ONLY | ANALYSIS |
| 51 | `usp_FE_factObtieneCodigoBHQ` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 52 | `usp_GenInsertarDiaContable` | `clsDiaContable.cls` | SQL_ONLY | ANALYSIS |
| 53 | `usp_GenObtieneDiaContable` | `clsDiaContable.cls`, `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`… | SQL_ONLY | ANALYSIS |
| 54 | `usp_Inforest_ActualizaCostos` | `frmAcceso.frm`, `frmFlash.frm` | SQL_ONLY | ANALYSIS |
| 55 | `usp_Inforest_DescargoVenta` | `frmDescargo.frm` | SQL_ONLY | ANALYSIS |
| 56 | `usp_Inforest_DescargoVentaPorInsumo` | `frmDescargo.frm` | SQL_ONLY | ANALYSIS |
| 57 | `usp_Inforest_Impresion` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCtaCteDetalle.frm`… | SQL_ONLY | ANALYSIS |
| 58 | `usp_Inforest_InicializaCostos` | `frmAcceso.frm`, `frmFlash.frm` | SQL_ONLY | ANALYSIS |
| 59 | `usp_Inforest_InsertaDocumento` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 60 | `usp_Inforest_InsertaLogDescargo` | `clsAlmacen.cls` | SQL_ONLY | ANALYSIS |
| 61 | `usp_Inforest_InsertarLogErrores` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 62 | `usp_Inforest_ObtenerClientesFrecuente` | `frmBusquedaSocio.frm` | SQL_ONLY | ANALYSIS |
| 63 | `usp_Inforest_ObtieneCentralPedidos_Adicicion` | `frmBusquedaCentralPedido.frm` | SQL_ONLY | ANALYSIS |
| 64 | `usp_Inforest_ObtieneClientesFactura` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmGeneracion.frm`… | SQL_ONLY | ANALYSIS |
| 65 | `usp_Inforest_ObtieneCodigoQR_Bol` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 66 | `usp_Inforest_ObtieneCodigoQR_CB_HASH_FACT` | `modProcedimiento.bas` | SQL_ONLY | ANALYSIS |
| 67 | `usp_Inforest_ObtieneDocumentos` | `FrmRecargarTarjeta.frm`, `frmDocumentoCorrelativo.frm` | SQL_ONLY | ANALYSIS |
| 68 | `usp_Inforest_ObtieneDocumentos_NC_bizlink` | `frmNotaCredito.frm` | SQL_ONLY | ANALYSIS |
| 69 | `usp_Inforest_ObtieneDocumentos_bizlink` | `frmDocumentoCorrelativo.frm` | SQL_ONLY | ANALYSIS |
| 70 | `usp_Inforest_ObtienePropiedadesBusqueda` | `frmCargoMozo.frm`, `frmVenta.frm` | SQL_ONLY | ANALYSIS |
| 71 | `usp_Inforest_ObtienePropiedadesComboBusqueda` | `frmCargoMozo.frm`, `frmVenta.frm` | SQL_ONLY | ANALYSIS |
| 72 | `usp_Inforest_ObtieneRecargas` | `FrmRecargarTarjeta.frm` | SQL_ONLY | ANALYSIS |
| 73 | `usp_Inforest_ObtieneRptGuia` | `frmCtaCteDetalle.frm` | SQL_ONLY | ANALYSIS |
| 74 | `usp_Inforest_PedidosCentralPedido` | `frmCentralPedidos.frm` | SQL_ONLY | ANALYSIS |
| 75 | `usp_Inforest_ValidaClienteSel` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCambio.frm`, `frmVenta.frm`, `frmSolicitudBoleta.frm`… | SQL_ONLY | MIGRATED |
| 76 | `usp_ListDocumentosFE` | `frmDocumentoElectronicoCorrelativo.frm`, `frmLiquidacionDetalle.frm` | SQL_ONLY | ANALYSIS |
| 77 | `usp_RepInforest_DescargoVenta` | `frmDescargo.frm` | SQL_ONLY | ANALYSIS |
| 78 | `usp_Seg_verConexiones` | `frmConexionesActuales.frm` | SQL_ONLY | ANALYSIS |
| 79 | `usp_TicketComandaVistaPrevia` | `frmVenta.frm` | SQL_ONLY | ANALYSIS |
| 80 | `usp_WS_Inforest` | `frmVenta.frm` | SQL_ONLY | ANALYSIS |

## Views detectadas (InfoRest)

| ID | View | Evidencia VB6 (muestras) | Estado |
|---:|---|---|---|
| 1 | `VTIPOEGRESO` | `frmReciboEgresoDetalle.frm` | ANALYSIS |
| 2 | `vArea` | `frmAreasImpresora.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 3 | `vAreaImpresora` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmGeneracion.frm`… | ANALYSIS |
| 4 | `vBalanza` | `modPuntoVenta.bas` | ANALYSIS |
| 5 | `vCliente` | `FrmRecargarTarjetaDetalle.frm`, `frmBusquedaDelivery.frm`, `frmBusquedaRapida.frm`… | ANALYSIS |
| 6 | `vCompania` | `frmCajaRapida.frm`, `frmCtaCte.frm`, `frmCtaCteDetalle.frm`… | ANALYSIS |
| 7 | `vCortesia` | `frmCajaRapida.frm`, `frmCtaCteDetalle.frm`, `frmDetallePedido.frm`… | ANALYSIS |
| 8 | `vCtaCte` | `frmCajaRapida.frm`, `frmCtaCteDetalle.frm`, `frmVenta.frm` | ANALYSIS |
| 9 | `vDelivery` | `frmBusquedaDelivery.frm`, `frmBusquedaSocio.frm`, `frmCajaRapida.frm`… | ANALYSIS |
| 10 | `vDespachador` | `frmDespachador.frm`, `frmPedidoDelivery.frm`, `frmPedidoDeliveryNo.frm` | MIGRATED |
| 11 | `vDistrito` | `frmNuevoDelivery.frm` | ANALYSIS |
| 12 | `vDocumento` | `frmCuentaCobrarDetalle.frm`, `modImpresoraFiscal.bas`, `modProcedimiento.bas` | ANALYSIS |
| 13 | `vDocumentoAgrupado` | `frmBusquedaCentralPedido.frm`, `frmCentralPedidos.frm`, `frmPedidoAnterior.frm`… | ANALYSIS |
| 14 | `vDocumentoCorrelativoDetalle` | `frmDocumentoCorrelativoDetalle.frm` | ANALYSIS |
| 15 | `vDocumentoGrilla` | `frmCuentaCobrar.frm` | ANALYSIS |
| 16 | `vDocumentoImpresora` | `frmCambio.frm`, `frmCuentaCobrarDetalle.frm`, `modProcedimiento.bas` | ANALYSIS |
| 17 | `vDocumentoImpresoraAgrupado` | `modProcedimiento.bas` | ANALYSIS |
| 18 | `vDocumentoPago` | `frmDocumentoPago.frm` | ANALYSIS |
| 19 | `vEgreso` | `frmGuiaTransporte.frm`, `frmReciboEgreso.frm`, `frmReciboEgresoDetalle.frm`… | ANALYSIS |
| 20 | `vEstadoDocumento` | `frmDocumentoCompra.frm`, `frmDocumentoCorrelativo.frm`, `frmRepLiquidacionTicket.frm` | ANALYSIS |
| 21 | `vEstadoFrecuente` | `frmBusquedaSocio.frm`, `frmNuevoDelivery.frm`, `frmVenta.frm` | ANALYSIS |
| 22 | `vEstadoReserva` | `frmRepReservas.frm`, `frmReserva.frm` | ANALYSIS |
| 23 | `vGrupo` | `frmCargoMozo.frm`, `frmRepEntrega.frm`, `frmRepPaloteoTicket.frm`… | ANALYSIS |
| 24 | `vGrupoUsuario` | `frmAcceso.frm`, `frmMesa.frm`, `frmMozoUsuario.frm`… | ANALYSIS |
| 25 | `vGuiaTransporte` | `frmCtaCteDetalle.frm`, `frmGuiaTransporte.frm`, `frmGuiaTransporteCorrelativoDetalle.frm` | ANALYSIS |
| 26 | `vIngreso` | `frmPago.frm`, `frmPagoMasivo.frm`, `frmPrePago.frm`… | ANALYSIS |
| 27 | `vInvitado` | `frmBusquedaDelivery.frm`, `frmBusquedaSocio.frm` | ANALYSIS |
| 28 | `vLocal` | `frmRepPaloteoTicket.frm`, `modProcedimiento.bas` | ANALYSIS |
| 29 | `vMoneda` | `FrmSeleccionProductos.frm`, `frmDocumentoCompra.frm`, `frmInicio.frm`… | ANALYSIS |
| 30 | `vMotivoAnulacion` | `frmDocumento.frm`, `frmListaMotivos.frm` | ANALYSIS |
| 31 | `vMotivoDescuento` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmDivision.frm`… | ANALYSIS |
| 32 | `vMotivoEliminacion` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmVenta.frm` | ANALYSIS |
| 33 | `vMotorizado` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmDocumento.frm`… | ANALYSIS |
| 34 | `vMozo` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmMesa.frm`… | ANALYSIS |
| 35 | `vNotaCredito` | `frmGeneracion.frm`, `frmNotaCredito.frm`, `frmPago.frm`… | ANALYSIS |
| 36 | `vNotaCreditoSunat` | `modProcedimiento.bas` | ANALYSIS |
| 37 | `vOperador` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmDetallePedido.frm`… | ANALYSIS |
| 38 | `vOrigenVenta` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmRepPaloteoTicket.frm`… | ANALYSIS |
| 39 | `vPariente` | `frmCargoMozo.frm`, `frmVenta.frm` | ANALYSIS |
| 40 | `vPedido` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmPedido.frm`… | ANALYSIS |
| 41 | `vPedidoCabecera` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmGeneracion.frm`… | ANALYSIS |
| 42 | `vPedidoCombo` | `frmCargoMozo.frm`, `frmDetallePedido.frm`, `frmPedido.frm`… | ANALYSIS |
| 43 | `vPedidoCombox` | `frmCargoMozo.frm`, `frmVenta.frm` | ANALYSIS |
| 44 | `vPedidoCorrelativo` | `frmPedidoCorrelativo.frm` | ANALYSIS |
| 45 | `vPedidoDetalle` | `frmCargoMozo.frm`, `frmDetallePedido.frm`, `frmDivision.frm`… | ANALYSIS |
| 46 | `vPedidoGrilla` | `frmCajaRapida.frm`, `frmDivision.frm`, `frmDocumento.frm`… | ANALYSIS |
| 47 | `vPreCuenta` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmPedidoCorrelativo.frm`… | ANALYSIS |
| 48 | `vPreCuentaDelivery` | `frmCargoMozo.frm`, `frmDocumento.frm`, `frmPedidoCorrelativo.frm`… | ANALYSIS |
| 49 | `vPreCuentaDetallada` | `frmCargoMozo.frm`, `frmPedidoCorrelativo.frm`, `frmVenta.frm` | ANALYSIS |
| 50 | `vPrecuentaAgrupada` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmPedidoCorrelativo.frm`… | ANALYSIS |
| 51 | `vProducto` | `frmBusquedaRapida.frm`, `frmBusquedaValeConsumo.frm`, `frmCajaRapida.frm`… | ANALYSIS |
| 52 | `vSalon` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmIr.frm`… | ANALYSIS |
| 53 | `vSubGrupo` | `frmCargoMozo.frm`, `frmRepEntrega.frm`, `frmRepPaloteoTicket.frm`… | ANALYSIS |
| 54 | `vSubTipoCtaCte` | `frmNuevoDelivery.frm`, `frmRepCtaCte.frm` | IN_PROGRESS |
| 55 | `vTienda` | `frmBusquedaDelivery.frm`, `frmCtaCteDetalle.frm` | ANALYSIS |
| 56 | `vTipoCancelacion` | `frmLiquidacionDetalle.frm`, `frmPago.frm`, `frmPagoMasivo.frm`… | ANALYSIS |
| 57 | `vTipoClienteFrecuente` | `frmClienteDeliveryFrecuente.frm`, `frmNuevoDelivery.frm`, `frmVenta.frm` | ANALYSIS |
| 58 | `vTipoCtaCte` | `frmNuevoDelivery.frm`, `frmRepCtaCte.frm` | IN_PROGRESS |
| 59 | `vTipoDocumento` | `FrmRecargarTarjetaDetalle.frm`, `FrmSeleccionProductos.frm`, `frmCajaRapida.frm`… | ANALYSIS |
| 60 | `vTipoIdentidad` | `frmClienteDeliveryFrecuente.frm`, `frmNuevoCliente.frm`, `frmNuevoDelivery.frm`… | ANALYSIS |
| 61 | `vTipoPago` | `frmRepDeliveryTicket.frm`, `frmRepLiquidacionTicket.frm`, `modProcedimiento.bas` | ANALYSIS |
| 62 | `vTipoPedido` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmCentralPedidos.frm`… | ANALYSIS |
| 63 | `vTipoProducto` | `frmRepPaloteoTicket.frm` | ANALYSIS |
| 64 | `vTipoResponsable` | `frmNuevoCliente.frm`, `modImpresoraFiscal.bas` | ANALYSIS |
| 65 | `vTipodocumentoImpresora` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCambio.frm`… | ANALYSIS |
| 66 | `vVehiculo` | `frmCtaCteDetalle.frm` | ANALYSIS |
| 67 | `vZona` | `frmBusquedaDelivery.frm`, `frmNuevoDelivery.frm` | ANALYSIS |

## Tablas detectadas (InfoRest)

| ID | Tabla | Evidencia VB6 (muestras) | Estado |
|---:|---|---|---|
| 1 | `APEDIDO` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmVenta.frm` | ANALYSIS |
| 2 | `CPEDIDO` | `frmBusquedaCentralPedido.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 3 | `DDOCUMENTO` | `FrmRecargarTarjetaDetalle.frm`, `FrmSeleccionProductos.frm`, `frmCajaRapida.frm`… | ANALYSIS |
| 4 | `DGUIATRANSPORTE` | `frmCtaCteDetalle.frm` | ANALYSIS |
| 5 | `DINGRESO` | `frmPago.frm`, `frmReciboIngreso.frm` | ANALYSIS |
| 6 | `DNOTACREDITO` | `frmNotaCreditoDetalle.frm` | ANALYSIS |
| 7 | `DPAGODOCUMENTO` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCambio.frm`… | ANALYSIS |
| 8 | `DPAGODOCUMENTO_VC` | `frmCuentaCobrar.frm`, `frmDocumento.frm`, `frmDocumentoCorrelativo.frm`… | ANALYSIS |
| 9 | `DPAGOTARJETA` | `frmCajaRapida.frm`, `frmDocumento.frm`, `frmDocumentoCorrelativo.frm`… | ANALYSIS |
| 10 | `DPEDIDO` | `FrmRecargarTarjetaDetalle.frm`, `frmBusquedaDelivery.frm`, `frmCajaRapida.frm`… | ANALYSIS |
| 11 | `DPREPAGO` | `frmPago.frm`, `frmPagoMasivo.frm`, `frmPagoRapido.frm`… | ANALYSIS |
| 12 | `INFOVISOR` | `FrmMensajeVisor.frm`, `frmBusquedaRapida.frm`, `frmBusquedaValeConsumo.frm`… | ANALYSIS |
| 13 | `INTEGRACION_CASHDRO` | `frmPago.frm` | ANALYSIS |
| 14 | `LOG_INFOREST` | `clsTrama.cls`, `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`… | ANALYSIS |
| 15 | `LOG_SESIONES` | `mdiPuntoVenta.frm` | ANALYSIS |
| 16 | `MCIERRE` | `frmDescargo.frm`, `modProcedimiento.bas` | ANALYSIS |
| 17 | `MDOCUMENTO` | `FrmRecargarTarjetaDetalle.frm`, `FrmSeleccionProductos.frm`, `frmBusquedaDelivery.frm`… | ANALYSIS |
| 18 | `MEGRESO` | `frmLiquidacionDetalle.frm`, `frmReciboEgresoDetalle.frm` | ANALYSIS |
| 19 | `MGUIATRANSPORTE` | `frmCtaCteDetalle.frm`, `frmDocumento.frm`, `frmNotaCreditoDetalle.frm` | ANALYSIS |
| 20 | `MINGRESO` | `frmCuentaCobrar.frm`, `frmDocumento.frm`, `frmDocumentoCorrelativo.frm`… | ANALYSIS |
| 21 | `MNOTACREDITO` | `frmCuentaCobrar.frm`, `frmDocumento.frm`, `frmDocumentoCompra.frm`… | ANALYSIS |
| 22 | `MPEDIDO` | `FrmRecargarTarjetaDetalle.frm`, `frmBusquedaCentralPedido.frm`, `frmBusquedaDelivery.frm`… | ANALYSIS |
| 23 | `MPROPINA` | `frmCajaRapida.frm`, `frmPedidoCorrelativo.frm`, `frmVenta.frm` | ANALYSIS |
| 24 | `MTURNO` | `frmDocumentoCorrelativo.frm`, `frmInicio.frm`, `frmLiquidacionDetalle.frm`… | ANALYSIS |
| 25 | `PEDIDO_PAGOEFECTIVO` | `frmVenta.frm` | ANALYSIS |
| 26 | `TACCESO` | `modProcedimiento.bas`, `modSeguridadInfhotel.bas` | ANALYSIS |
| 27 | `TAREASUBGRUPO` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmImportacionRequerimientos.frm`… | ANALYSIS |
| 28 | `TCAJA` | `FrmMensajeVisor.frm`, `FrmRecargarTarjetaDetalle.frm`, `frmBusquedaRapida.frm`… | ANALYSIS |
| 29 | `TCAJACANALVENTA` | `frmVenta.frm`, `mdiPuntoVenta.frm` | ANALYSIS |
| 30 | `TCAJAORIGEN_BLOQUEO` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmVenta.frm` | ANALYSIS |
| 31 | `TCAJATERMINAL` | `frmDocumento.frm`, `frmLiquidacionDetalle.frm`, `frmPago.frm`… | ANALYSIS |
| 32 | `TCANALVENTA` | `frmCajaRapida.frm`, `frmDocumento.frm`, `frmPago.frm`… | ANALYSIS |
| 33 | `TCENTROCOSTO` | `FrmSeleccionProductos.frm`, `frmDocumentoCompra.frm`, `frmReciboEgresoDetalle.frm` | ANALYSIS |
| 34 | `TCLIENTE` | `FrmRecargarTarjetaDetalle.frm`, `frmBusquedaRapida.frm`, `frmBusquedaValeConsumo.frm`… | ANALYSIS |
| 35 | `TCLIENTEPRODUCTO` | `frmBusquedaRapida.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 36 | `TCOMBO` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmVenta.frm`… | ANALYSIS |
| 37 | `TCOMBOPROPIEDAD` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmDetallePedido.frm`… | ANALYSIS |
| 38 | `TCOMPANIA` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargo.frm`… | ANALYSIS |
| 39 | `TDELIVERY` | `FrmRecargarTarjetaDetalle.frm`, `FrmTarjetaAproximidad.frm`, `FrmTarjetaAproximidadDetalle.frm`… | ANALYSIS |
| 40 | `TDELIVERYCLIENTE` | `FrmRecargarTarjetaDetalle.frm`, `frmReciboIngresoDetalle.frm`, `frmVenta.frm` | ANALYSIS |
| 41 | `TDELIVERYINVITADO` | `frmBusquedaDelivery.frm`, `frmBusquedaSocio.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 42 | `TDESCARGOINSUMO` | `frmDescargo.frm` | ANALYSIS |
| 43 | `TDIACONTABLE` | `frmDiaContable.frm` | ANALYSIS |
| 44 | `TESTADOBIZLINK` | `frmNotaCreditoDetalle.frm` | ANALYSIS |
| 45 | `TESTADOINFOFACT` | `frmCtaCteDetalle.frm` | ANALYSIS |
| 46 | `TGRUPO` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargo.frm`… | ANALYSIS |
| 47 | `TGRUPOACCESO` | `frmReciboEgresoDetalle.frm`, `modProcedimiento.bas` | ANALYSIS |
| 48 | `TGRUPOUSUARIO` | `frmCargoMozo.frm`, `frmDocumentoCompra.frm`, `frmDocumentoCorrelativo.frm`… | ANALYSIS |
| 49 | `TIMPRESORA` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCambio.frm`… | ANALYSIS |
| 50 | `TINSUMO` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmInsumoDetalle.frm`… | ANALYSIS |
| 51 | `TINTEGRACIONES` | `frmDocumentoCorrelativo.frm`, `frmPago.frm` | ANALYSIS |
| 52 | `TLOCAL` | `frmAcceso.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 53 | `TLOG_IMPRESION` | `frmCajaRapida.frm`, `modProcedimiento.bas` | ANALYSIS |
| 54 | `TLOG_MODPRECIO` | `frmPreciosDetalle.frm` | ANALYSIS |
| 55 | `TMENSAJE` | `frmDescargo.frm`, `frmMensaje.frm`, `frmPedido.frm` | ANALYSIS |
| 56 | `TMENSAJECOCINA` | `frmMensajeCocinaDetalle.frm` | MIGRATED |
| 57 | `TMENSAJEUSUARIO` | `frmPedido.frm` | ANALYSIS |
| 58 | `TMESA` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 59 | `TMODULO` | `modAuditoriaIntegral.bas`, `modProcedimiento.bas` | ANALYSIS |
| 60 | `TMOVIMIENTOTARJETASRFID` | `FrmRecargarTarjetaDetalle.frm`, `FrmTarjetaAproximidad.frm` | IN_PROGRESS |
| 61 | `TOFERTA` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 62 | `TOPERADOR` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmDetallePedido.frm`… | ANALYSIS |
| 63 | `TORIGENCODIGOCONTROL` | `modProcedimiento.bas` | ANALYSIS |
| 64 | `TORIGENVENTA` | `modProcedimiento.bas` | ANALYSIS |
| 65 | `TPARAMETRO` | `claCorreoElectronico.cls`, `FrmRecargarTarjetaDetalle.frm`, `FrmSeleccionProductos.frm`… | ANALYSIS |
| 66 | `TPARIENTE` | `frmBusquedaSocio.frm`, `modProcedimiento.bas` | ANALYSIS |
| 67 | `TPEDIDO` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 68 | `TPEDIDOMESA` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmJuntaMesa.frm`… | ANALYSIS |
| 69 | `TPRODUCTO` | `FrmSeleccionProductos.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 70 | `TPRODUCTOAREA` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmVenta.frm`… | ANALYSIS |
| 71 | `TPRODUCTOPROPIEDAD` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmDetallePedido.frm`… | ANALYSIS |
| 72 | `TPROGRAMAPRECIOS_CAB` | `modProcedimiento.bas` | ANALYSIS |
| 73 | `TPROGRAMAPRECIOS_DETA` | `modProcedimiento.bas` | ANALYSIS |
| 74 | `TPROPIEDAD` | `frmCajaRapida.frm`, `frmCargoMozo.frm`, `frmDetallePedido.frm`… | ANALYSIS |
| 75 | `TRESERVA` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 76 | `TSUBGRUPO` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCargoMozo.frm`… | ANALYSIS |
| 77 | `TTABLA` | `claCorreoElectronico.cls`, `frmCajaRapida.frm`, `frmDescargo.frm`… | ANALYSIS |
| 78 | `TTARJETACREDITO` | `FrmRecargarTarjetaDetalle.frm`, `frmDocumento.frm`, `frmDocumentoCorrelativo.frm`… | ANALYSIS |
| 79 | `TTARJETASRFID` | `FrmRecargarTarjetaDetalle.frm`, `FrmTarjetaAproximidad.frm`, `FrmTarjetaAproximidadDetalle.frm` | IN_PROGRESS |
| 80 | `TTERMINAL` | `frmDocumento.frm`, `frmLiquidacionDetalle.frm`, `frmPago.frm`… | ANALYSIS |
| 81 | `TTIENDA` | `FrmRecargarTarjetaDetalle.frm`, `frmBusquedaDelivery.frm`, `frmCajaRapida.frm`… | ANALYSIS |
| 82 | `TTIPOCAMBIO` | `frmAcceso.frm`, `frmDocumentoCompra.frm`, `frmDocumentoCorrelativo.frm`… | ANALYSIS |
| 83 | `TTIPODOCUMENTO` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCambio.frm`… | ANALYSIS |
| 84 | `TTIPODOCUMENTOIMPRESORA` | `FrmRecargarTarjetaDetalle.frm`, `frmCajaRapida.frm`, `frmCambio.frm`… | ANALYSIS |
| 85 | `TTIPOIDENTIDAD` | `frmClienteDeliveryFrecuente.frm`, `frmNuevoCliente.frm`, `frmNuevoDelivery.frm`… | ANALYSIS |
| 86 | `TUBIGEO` | `frmNuevoCliente.frm`, `modImpresoraFiscal.bas`, `modProcedimiento.bas`… | ANALYSIS |
| 87 | `TUSUARIO` | `clsDiaContable.cls`, `FrmRecargarTarjetaDetalle.frm`, `frmAcceso.frm`… | ANALYSIS |
| 88 | `TVISIBILIDADTARJETACREDITOXCANAL` | `frmPago.frm`, `frmPrePago.frm` | ANALYSIS |
| 89 | `VALE_CONSUMO` | `frmBusquedaValeConsumo.frm`, `frmCuentaCobrar.frm`, `frmDocumento.frm`… | ANALYSIS |
| 90 | `VISIBILIDADPROPIEDADXCANAL` | `frmCajaRapida.frm` | ANALYSIS |
| 91 | `VISOR_DPEDIDO` | `frmCajaRapida.frm` | ANALYSIS |

## Trazas VB6 → SQL (ejemplos críticos)

| Formulario | Evento/Flujo | VB6 | SQL/Vista/SP | Tabla(s) | Estado |
|---|---|---|---|---|---|
| `frmVenta.frm` | Registro de venta/pedido | `modProcedimiento.bas`, `clsComando.cls` | `spIns_MPEDIDO`, `spUpd_MPEDIDO`, `spIns_DPEDIDO` | `MPEDIDO`, `DPEDIDO`, `TPEDIDOMESA` | ANALYSIS |
| `frmPago.frm` | Cobro documento | `modProcedimiento.bas` | `vDocumentoPago`, `DPAGODOCUMENTO`, `DPAGOTARJETA` | `DPAGODOCUMENTO`, `DPAGOTARJETA` | ANALYSIS |
| `frmDiaContable.frm` | Apertura/cierre día | `clsDiaContable.cls` | `usp_GenObtieneDiaContable`, `usp_GenInsertarDiaContable` | `TDIACONTABLE`, `MTURNO` | ANALYSIS |
| `frmCentralPedidos.frm` | Central de pedidos | `modPuntoVenta.bas` + delivery modules | `usp_Inforest_PedidosCentralPedido` | `TPEDIDO`, `MPEDIDO` | ANALYSIS |
| `frmInsumo.frm` | Gestión/descargo insumos | `clsAlmacen.cls` | `USP_MODIFICARINSUMOS`, `usp_Inforest_DescargoVenta` | `TINSUMO`, `TDESCARGOINSUMO` | ANALYSIS |

## SQL dinámico y riesgos

- Líneas con SQL dinámico detectadas: **5723**.
- Patrones observados: concatenación de `SELECT/INSERT/UPDATE/DELETE/EXEC` con parámetros de UI/estado global.
- Riesgos: SQL injection, lógica de negocio implícita en cadenas, acoplamiento UI-DB, transacciones no explícitas en código de formulario.
- Acción recomendada: cada query dinámica crítica debe convertirse en contrato explícito (SP parametrizado o query parametrizada Dapper) antes de marcar `COMPLETED`.