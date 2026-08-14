# Punto de Venta — Missing Functionality

## Resumen de faltantes detectados

| Métrica | Valor |
|---|---:|
| Formularios legacy POS (`InfoRest.vbp`) | 120 |
| Formularios con referencia explícita en .NET | 14 |
| Formularios sin referencia explícita en .NET | 106 |
| Módulos BAS legacy POS | 18 |
| Módulos BAS mencionados en .NET | 10 |
| Clases legacy POS | 10 |
| Clases mencionadas en .NET | 8 |

## Formularios POS faltantes en trazabilidad .NET explícita

```text
frmAbout.frm
FrmActualizacion.frm
frmAreasImpresora.frm
frmBusca.frm
frmBusquedaDelivery.frm
frmBusquedaRapida.frm
frmBusquedaRapidaInvitado.frm
frmBusquedaSocio.frm
frmBusquedaValeConsumo.frm
frmCambiarContrasenia.frm
frmCambio.frm
frmCambioPropina.frm
frmCargo.frm
frmCargoMozo.frm
frmConexionesActuales.frm
frmConsultaSaldo.frm
frmCrearIniTVS.frm
frmCtaCte.frm
frmCtaCteDetalle.frm
frmCuentaCobrar.frm
frmCuentaCobrarDetalle.frm
frmDescargo.frm
frmDetallePedido.frm
frmDiaContable.frm
frmDivision.frm
frmDocumentoCompra.frm
frmDocumentoCorrelativo.frm
frmDocumentoCorrelativoDetalle.frm
frmDocumentoElectronicoCorrelativo.frm
frmDocumentoPago.frm
frmEmite.frm
frmFactura.frm
frmFechaEmision.frm
frmFiltroRecibo.frm
frmFlash.frm
frmGeneracion.frm
frmGuiaTransporte.frm
frmGuiaTransporteCorrelativoDetalle.frm
frmImportacionRequerimientoDetalle.frm
frmImportacionRequerimientos.frm
frmInicio.frm
frmInsumo.frm
frmInsumoDetalle.frm
frmIr.frm
frmJuntaMesa.frm
frmJuntarMesas.frm
frmKeyboard.frm
frmLiquidacionDetalle.frm
frmListaMotivos.frm
frmMarca.frm
frmMensaje.frm
frmMensajeCocina.frm
frmMensajeCocinaDetalle.frm
frmMensajeFeSpring.frm
FrmMensajeVisor.frm
frmMesaConsulta.frm
FrmMovimientoTarjetas.frm
frmMozoUsuario.frm
frmMsComm.frm
frmMsgBox.frm
frmNotaCredito.frm
frmNotaCreditoDetalle.frm
frmNoticias.frm
frmNuevaTienda.frm
frmNuevo.frm
frmNuevoCliente.frm
frmNuevoInvitado.frm
frmNumPad.frm
frmPagoMasivo.frm
frmPagoPinPad.frm
frmPagoRapido.frm
frmPassword.frm
frmPedidoAnterior.frm
frmPedidoCorrelativo.frm
frmPedidoDeliveryNo.frm
frmPedidoDetalle.frm
frmPedidosVerPagos.frm
frmPrecios.frm
frmPreciosDetalle.frm
frmPrecuentaImpresora.frm
frmPrePago.frm
frmPrograma.frm
FrmRecargarTarjeta.frm
FrmRecargarTarjetaDetalle.frm
frmReciboEgreso.frm
frmReciboEgresoDetalle.frm
frmReciboIngreso.frm
frmReciboIngresoDetalle.frm
frmRepCtaCte.frm
frmRepDeliveryTicket.frm
frmRepEntrega.frm
frmRepLiquidacionTicket.frm
frmRepPaloteoTicket.frm
frmRepReservas.frm
frmReserva.frm
frmReservaDetalle.frm
FrmSeleccionProductos.frm
frmServidorEnlace.frm
frmSolicitudBoleta.frm
FrmTarjetaAproximidad.frm
FrmTarjetaAproximidadDetalle.frm
frmUpdateDatosPedido.frm
frmVerificacionHuella.frm
frmVerificacionHuellaSup.frm
frmVerificarHuellaMarcacion.frm
frmWinSock.frm
```

## Dependencias críticas pendientes

| Dependencia | Estado | Acción |
|---|---|---|
| Biometría SecuGen / DigitalPersona | BLOCKED | Definir sustituto .NET o servicio externo |
| PinPad DLL3500 validación real | PARTIAL | Validar en hardware real y contract tests |
| CashDro integración operativa real | PARTIAL | Validar endpoint/protocolo productivo |
| Impresión térmica por área | PARTIAL | Completar adapter y pruebas de impresión real |
| Fiscal Epson (`IFEpson.ocx`) | EXCLUDED_BY_SCOPE | Mantener fuera de alcance por requerimiento actual |

## UNKNOWN_BEHAVIOR registrados

- `frmSolicitudBoleta.frm`: flujo funcional exacto pendiente de lectura completa de eventos.
- `frmMarca.frm`: reglas de marcación no totalmente determinadas.
- `frmNoticias.frm` + `FrmActualizacion.frm`: criterio de actualización/versionado en operación real pendiente.
- Flujos RFID (`FrmTarjetaAproximidad*`, `FrmRecargarTarjeta*`): reglas operativas de negocio incompletas.
- Integración hotelera (`INFHOTEL.INI`): reglas cross-sistema no cerradas para POS 100%.