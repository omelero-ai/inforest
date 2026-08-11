# Inventario Legacy — INFOREST
> Inventario maestro de artefactos Legacy VB6 + SQL Server basado únicamente en evidencia del repositorio.
> Última actualización: 2026-08-11
> Nota: este inventario sigue los conteos reales extraídos del árbol actual (`400` formularios `.frm`, `206` reportes `.Dsr`, `116` vistas y `150` stored procedures). Cuando difieren de documentación previa, prevalece este inventario porque se generó desde los archivos fuente reales.

---

## Resumen Ejecutivo

| Área | Conteo | Evidencia |
|---|---:|---|
| Proyectos VB6 (`.vbp`) | 7 | `legacy-restaurant/restaurant-vb6/*.vbp` |
| Ejecutables definidos en VBP | 7 | `ExeName32=` en proyectos VB6 |
| Formularios (`.frm`) | 400 | `legacy-restaurant/restaurant-vb6/Formularios/` |
| Módulos (`.bas`) | 32 | `legacy-restaurant/restaurant-vb6/Modulos/` |
| Clases (`.cls`) | 10 | `legacy-restaurant/restaurant-vb6/Clases/` |
| Reportes Crystal (`.Dsr`) | 206 | `legacy-restaurant/restaurant-vb6/Reportes/` |
| Artefactos Crystal auxiliares (`.DCA`) | 206 | `Reportes/` |
| Artefactos Crystal auxiliares (`.dsx`) | 205 | `Reportes/` |
| Tablas SQL Server | 126 | `1. Estructura.sql` |
| Vistas SQL Server | 116 | `4. Vistas.sql` |
| Stored procedures | 150 | `5. SP.sql` |
| OCX/DLL físicos en raíz VB6 | 5 | `restaurant-vb6/*.ocx`, `*.dll` |
| Referencias COM/ActiveX únicas en VBP | 43 | `Object=` y `Reference=` |

## 1. Aplicación Legacy VB6

### 1.1 Proyectos y ejecutables definidos

| Proyecto VBP | EXE definido | Startup | Formularios | Módulos | Clases |
|---|---|---|---:|---:|---:|
| `Adicion.vbp` | `Adicion.exe` | `Sub Main` | 37 | 12 | 6 |
| `Administracion.vbp` | `Administracion.exe` | `Sub Main` | 151 | 13 | 7 |
| `CajaRapida.vbp` | `CajaRapida.exe` | `Sub Main` | 100 | 16 | 8 |
| `Consulta.vbp` | `Consulta.exe` | `Sub Main` | 124 | 15 | 9 |
| `Despachador.vbp` | `Despachador.exe` | `Sub Main` | 25 | 10 | 3 |
| `InfoRest.vbp` | `InfoRest.exe` | `Sub Main` | 120 | 18 | 10 |
| `Motorizados.vbp` | `Motorizado.exe` | `Sub Main` | 2 | 7 | 3 |

### 1.2 OCX, DLL, ActiveX y referencias COM

**Archivos físicos detectados en `restaurant-vb6/`:**

- OCX: `IFEpson.ocx`
- DLL: `CAJA_PINPAD.dll`, `dzstactx.dll`, `hkey-w32.dll`, `qrcodelib.dll`
- INI: `Almacen.ini`, `DLL3500.ini`, `Facturacion.ini`, `Infhotel.ini`, `Inforest.ini`, `Tiempo.ini`

**Objetos `Object=` únicos declarados en los `.vbp`:**

```text
CRViewer.dll
MCI32.OCX
MSCHRT20.OCX
MSCOMCTL.OCX
MSCOMM32.OCX
MSDATLST.OCX
MSINET.OCX
MSMAPI32.OCX
MSWINSCK.OCX
TABCTL32.OCX
comdlg32.ocx
crviewer9.dll
duzactx.dll
dzactx.dll
dzstactx.dll
ifepson.ocx
mscomct2.ocx
sgfplibx.ocx
todg8.ocx
wmp.dll
```

**Referencias `Reference=` únicas declaradas en los `.vbp`:**

```text
BlueVision_Core_TVS
Chilkat Certificate
Chilkat Mail v7.9.0
Chilkat Util
Crystal Data Object
Crystal Report 6.0 ActiveX Designer Run Time Library
Crystal Reports 6.0 Library
Crystal Reports 9 ActiveX Designer Design and Runtime Library
DigitalPersona One Touch for Windows Control 1.0
DigitalPersona One Touch for Windows Device components 1.0
DigitalPersona One Touch for Windows Engine components 1.0
DigitalPersona One Touch for Windows Shared components 1.0
Libreria16
Microsoft ActiveX Data Objects (Multi-dimensional) 2.8 Library
Microsoft ActiveX Data Objects 2.6 Library
Microsoft ActiveX Data Objects 2.8 Library
Microsoft DAO 3.6 Object Library
Microsoft Data Binding Collection VB 6.0 (SP4)
Microsoft Data Formatting Object Library 6.0 (SP6)
Microsoft Data Report Designer 6.0 (SP4)
Microsoft Scripting Runtime
Microsoft XML, v6.0
OLE Automation
```

### 1.3 Inventario completo de formularios (`.frm`)

```text
Form1.frm
Form2.frm
FrmActualizacion.frm
FrmActualizacionAdmin.frm
FrmConsultaRUC.frm
FrmDispensadores.frm
FrmDispensadoresDetalle.frm
FrmMensajeVisor.frm
FrmMovimientoTarjetas.frm
FrmOptimizar.frm
FrmPreview.frm
FrmRecargarTarjeta.frm
FrmRecargarTarjetaDetalle.frm
FrmRepInventariosIntegrado.frm
FrmRepMovimientosIntegrado.frm
FrmRepMozos.frm
FrmSeleccionProductos.frm
FrmTarjetaAproximidad.frm
FrmTarjetaAproximidadDetalle.frm
FrmUpdate.frm
FrmUpdateAlmacen.frm
fmrPedidosEntregado.frm
frmAbout.frm
frmAcceso.frm
frmActualizar.frm
frmAdicion.frm
frmAgregarCliente.frm
frmAgrupacion.frm
frmAgrupacionDetalle.frm
frmAlmacen.frm
frmAlmacenDOS.frm
frmAlmacenVIS.frm
frmArea.frm
frmAreaDetalle.frm
frmAreasImpresora.frm
frmAsignacionMotorizado.frm
frmBackup.frm
frmBusca.frm
frmBuscaAlmacen.frm
frmBuscaPedido.frm
frmBuscarPropiedad.frm
frmBusquedaCentralPedido.frm
frmBusquedaDelivery.frm
frmBusquedaRapida.frm
frmBusquedaRapidaInvitado.frm
frmBusquedaSocio.frm
frmBusquedaValeConsumo.frm
frmCaja.frm
frmCajaCodigoControl.frm
frmCajaCodigoControlDetalle.frm
frmCajaDetalle.frm
frmCajaDetalle2.frm
frmCajaRapida.frm
frmCajaRapidaC.frm
frmCajon.frm
frmCajonDetalle.frm
frmCalendario.frm
frmCambiarContrasenia.frm
frmCambio.frm
frmCambioPropina.frm
frmCargo.frm
frmCargoMozo.frm
frmCentralPedidos.frm
frmCentroCosto.frm
frmCentroCostoDetalle.frm
frmCheffControl.frm
frmChofer.frm
frmChoferDetalle.frm
frmCierrePeriodo.frm
frmCliente.frm
frmClienteDelivery.frm
frmClienteDeliveryDetalle.frm
frmClienteDeliveryFrecuente.frm
frmClienteDetalle.frm
frmClienteFactura.frm
frmClienteFacturaDetalle.frm
frmClienteProducto.frm
frmCodigoAutorizacion.frm
frmCodigoAutorizacionDetalle.frm
frmCombosDetalle.frm
frmComparativoConsumoInsumos.frm
frmConexionesActuales.frm
frmConfiguraMensaje.frm
frmConfiguraMensajeImpresora.frm
frmConfiguracionParametro.frm
frmConsultaSaldo.frm
frmCortesia.frm
frmCortesiaDetalle.frm
frmCrGuias.frm
frmCrearIniTVS.frm
frmCtaCte.frm
frmCtaCteDetalle.frm
frmCuentaCobrar.frm
frmCuentaCobrarDetalle.frm
frmCuentaContable.frm
frmCuentaPago.frm
frmDatoTransportista.frm
frmDatoTransportistaDetalle.frm
frmDescargo.frm
frmDespachador.frm
frmDetallePedido.frm
frmDiaContable.frm
frmDistrito.frm
frmDistritoDetalle.frm
frmDivision.frm
frmDocumento.frm
frmDocumentoCompra.frm
frmDocumentoCorrelativo.frm
frmDocumentoCorrelativoDetalle.frm
frmDocumentoElectronicoCorrelativo.frm
frmDocumentoPago.frm
frmDocumentosAnticipos.frm
frmEmite.frm
frmEmpacadora.frm
frmEmpacadoraDetalle.frm
frmEnvioDatosRappi.frm
frmEstadoClienteFrecuente.frm
frmEstadoClienteFrecuenteDetalle.frm
frmFactor.frm
frmFactura.frm
frmFecha.frm
frmFechaEmision.frm
frmFiltroProducto.frm
frmFiltroPropiedad.frm
frmFiltroRecibo.frm
frmFlash.frm
frmFormatoPedido.frm
frmGeneracion.frm
frmGeneracionPV.frm
frmGenerarsunat.frm
frmGrupo.frm
frmGrupoAcceso.frm
frmGrupoAccesoDetalle.frm
frmGrupoDetalle.frm
frmGrupoReplica.frm
frmGuiaTransporte.frm
frmGuiaTransporteCorrelativoDetalle.frm
frmGuiaTransporteDetalle.frm
frmGuias.frm
frmGuiasDetalle.frm
frmImportacion.frm
frmImportacionRequerimientoDetalle.frm
frmImportacionRequerimientos.frm
frmImpresora.frm
frmImpresoraDetalle.frm
frmInicio.frm
frmInicioAdicion.frm
frmInsumo.frm
frmInsumoDetalle.frm
frmInventarioConsolidado.frm
frmIr.frm
frmJuntaMesa.frm
frmJuntarMesas.frm
frmKDSConfiguracion.frm
frmKeyboard.frm
frmLiquidacion.frm
frmLiquidacionAutorizaDetalle.frm
frmLiquidacionDetalle.frm
frmLiquidacionMozo.frm
frmListaMotivos.frm
frmLlegadaSalida.frm
frmLocal.frm
frmLocalDetalle.frm
frmMaitre.frm
frmMaitreDetalle.frm
frmMantenimientoMesa.frm
frmMarca.frm
frmMarcar.frm
frmMensaje.frm
frmMensajeCocina.frm
frmMensajeCocinaDetalle.frm
frmMensajeFeSpring.frm
frmMesa.frm
frmMesaConsulta.frm
frmMesas.frm
frmMotivo.frm
frmMotivoAnulacion.frm
frmMotivoAnulacionDetalle.frm
frmMotivoDescuento.frm
frmMotivoDescuentoDetalle.frm
frmMotivoDetalle.frm
frmMotivoEliminacion.frm
frmMotivoEliminacionDetalle.frm
frmMotivoReserva.frm
frmMotivoReservaDetalle.frm
frmMotorizado.frm
frmMotorizadoDetalle.frm
frmMozo.frm
frmMozoDetalle.frm
frmMozoUsuario.frm
frmMsComm.frm
frmMsgBox.frm
frmNotaCredito.frm
frmNotaCreditoDetalle.frm
frmNoticias.frm
frmNuevaTienda.frm
frmNuevo.frm
frmNuevoCliente.frm
frmNuevoDelivery.frm
frmNuevoInvitado.frm
frmNumPad.frm
frmOferta.frm
frmOfertaDetalle.frm
frmOperador.frm
frmOperadorDetalle.frm
frmOrdenesConsola.frm
frmPago.frm
frmPagoMasivo.frm
frmPagoPinPad.frm
frmPagoRapido.frm
frmPapeCuadratura.frm
frmParametro.frm
frmParametroCheffControl.frm
frmPassword.frm
frmPedido.frm
frmPedidoAnterior.frm
frmPedidoCorrelativo.frm
frmPedidoDelivery.frm
frmPedidoDeliveryNo.frm
frmPedidoDetalle.frm
frmPedidosVerPagos.frm
frmPorciones.frm
frmPrePago.frm
frmPrecios.frm
frmPreciosDetalle.frm
frmPrecuenta.frm
frmPrecuentaImpresora.frm
frmProducto.frm
frmProductoBoton.frm
frmProductoDetalle.frm
frmProductoxProductoCantidad.frm
frmProductoxProductoDetalle.frm
frmProductoxproducto.frm
frmPrograma.frm
frmProgramacionPrecios.frm
frmProgramacionPreciosDetalle.frm
frmProperty.frm
frmPropertyDetalle.frm
frmPropiedad.frm
frmPropiedadDetalle.frm
frmPropina.frm
frmReasignacionMotorizado.frm
frmRecargoConsumo.frm
frmRecetaVentaDetalle.frm
frmReciboEgreso.frm
frmReciboEgresoDetalle.frm
frmReciboIngreso.frm
frmReciboIngresoDetalle.frm
frmReciboIngresoDetalleAdicion.frm
frmRepAnaliticoMotorizado.frm
frmRepAnaliticoMotorizadoIntegrado.frm
frmRepAnaliticoMozo.frm
frmRepAnaliticoMozoIntegrado.frm
frmRepAnulado.frm
frmRepAsistencia.frm
frmRepCancelacion.frm
frmRepClieFrecuentes.frm
frmRepClienteCtaCte.frm
frmRepClienteDelivery.frm
frmRepClienteDeuda.frm
frmRepCobranzaFecha.frm
frmRepComanda.frm
frmRepContacto.frm
frmRepControlEnviosAutorizados.frm
frmRepControlMotorizado.frm
frmRepCortesia.frm
frmRepCtaCte.frm
frmRepCtaCteIntegrado.frm
frmRepDeliveryTicket.frm
frmRepDescargoVentas.frm
frmRepDescuento.frm
frmRepDiarioVentaIntegrado.frm
frmRepDiferencia.frm
frmRepDiferenciaTiempoDelivery.frm
frmRepEntrega.frm
frmRepEntregaRegistro.frm
frmRepFormaPagoIntegrado.frm
frmRepGuias.frm
frmRepInsumoVentas.frm
frmRepInsumoVentasIntegrado.frm
frmRepLiquidacion.frm
frmRepLiquidacionTicket.frm
frmRepMensajesUsuario.frm
frmRepOcupabilidad.frm
frmRepPaloteo.frm
frmRepPaloteoComparativo.frm
frmRepPaloteoOfertas.frm
frmRepPaloteoPropiedades.frm
frmRepPaloteoSubProd.frm
frmRepPaloteoTicket.frm
frmRepPaloteoVentaIntegrado.frm
frmRepPedido.frm
frmRepPlanillaMovilidadMotorizado.frm
frmRepPrincipal.frm
frmRepProduccionMozo.frm
frmRepProductoMes.frm
frmRepProductosNoEnlazados.frm
frmRepPropina.frm
frmRepRanking.frm
frmRepRankingIntegrado.frm
frmRepRegistroComprasIntegrado.frm
frmRepRegistroVenta.frm
frmRepRegistroVentaIntegrado.frm
frmRepReservas.frm
frmRepResultadoOperativo.frm
frmRepRotacionMesa.frm
frmRepSolicitadoVsEntregado.frm
frmRepTiempoChefControl.frm
frmRepTiempoDelivery.frm
frmRepTiempoDeliveryIntegrado.frm
frmRepTiempoKDS.frm
frmRepTiempoSalon.frm
frmRepVentaAcumulada.frm
frmRepVentaCompAnual.frm
frmRepVentaCompMensual.frm
frmRepVentaFecha.frm
frmRepVentaMensualCanalesIntegrado.frm
frmRepVentaMensualIntegrado.frm
frmRepVentaTipoProductoIntegrado.frm
frmRepVentasTurno.frm
frmReportNotaCredito.frm
frmReserva.frm
frmReservaAdicion.frm
frmReservaDetalle.frm
frmReservaDetalleAdicion.frm
frmRestore.frm
frmSector.frm
frmSectorDetalle.frm
frmSectorVenta.frm
frmSectorVentaDetalle.frm
frmSeguridad.frm
frmServidorEnlace.frm
frmServidores.frm
frmSolicitudBoleta.frm
frmSucursal.frm
frmSucursalDetalle.frm
frmTarifaMotorizado.frm
frmTarjeta.frm
frmTarjetaDetalle.frm
frmTarjetaNumero.frm
frmTerminal.frm
frmTerminalDetalle.frm
frmTienda.frm
frmTiendaDetalle.frm
frmTipoCambio.frm
frmTipoCambioDetalle.frm
frmTipoCancelacion.frm
frmTipoCancelacionDetalle.frm
frmTipoCliente.frm
frmTipoClienteDetalle.frm
frmTipoCtaCte.frm
frmTipoCtaCteDetalle.frm
frmTipoDocumento.frm
frmTipoDocumentoDetalle.frm
frmTipoEgreso.frm
frmTipoEgresoDetalle.frm
frmTipoGrupoCliente.frm
frmTipoGrupoClienteDetalle.frm
frmTipoIdentidad.frm
frmTipoIdentidadDetalle.frm
frmTipoPedido.frm
frmTipoPedidoDetalle.frm
frmTipoProducto.frm
frmTipoProductoDetalle.frm
frmTransfer.frm
frmUnidadNegocio.frm
frmUnidadNegocioDetalle.frm
frmUpdateDatosPedido.frm
frmUsuario.frm
frmUsuarioDetalle.frm
frmUsuarioGrupo.frm
frmUsuarioGrupoDetalle.frm
frmUsuarioModulo.frm
frmVehiculo.frm
frmVehiculoDetalle.frm
frmVenta.frm
frmVentaPV.frm
frmVerificacionHuella.frm
frmVerificacionHuellaSup.frm
frmVerificarHuellaMarcacion.frm
frmWinSock.frm
frmZona.frm
frmZonaDetalle.frm
mdiAdministracion.frm
mdiCajaRapida.frm
mdiCheffControl.frm
mdiConsulta.frm
mdiConsultaIntegrada.frm
mdiDespachador.frm
mdiFalso.frm
mdiFastFood.frm
mdiGuias.frm
mdiMotorizado.frm
mdiPuntoVenta.frm
mdiPvCorp.frm
mdiUtilitario.frm
xfrmCajaRapida.frm
xfrmMozoUsuario.frm
xfrmRepOcupabilidad.frm
xfrmVenta.frm
```

### 1.4 Inventario completo de módulos (`.bas`)

```text
CodigoControl.bas
DLL3500.bas
FpLibX_Const.bas
ModPictureBoxCustom.bas
ModuloHardKey.bas
VBZipBas.bas
modAdicion.bas
modAdministracion.bas
modAuditoria.bas
modAuditoriaEquipo.bas
modAuditoriaIntegral.bas
modBarcode.bas
modBlueVision.bas
modCajaRapida.bas
modCheffControl.bas
modConexionIp.bas
modConsulta.bas
modConsultaIntregrada.bas
modCrearInis.bas
modDeclaracion.bas
modDespachador.bas
modGuias.bas
modImpresoraFiscal.bas
modKDS.bas
modMasticar.bas
modMotorizado.bas
modProcedimiento.bas
modProcedimientoNuevo.bas
modPuntoVenta.bas
modPvCorp.bas
modSeguridadInfhotel.bas
modTime.bas
```

### 1.5 Inventario completo de clases (`.cls`)

```text
ClsDocumento.cls
ClsSeguridad.cls
License.cls
Mapping.cls
claCorreoElectronico.cls
clsAlmacen.cls
clsComando.cls
clsDiaContable.cls
clsTrama.cls
clsxml.cls
```

## 2. Base de datos Legacy SQL Server

### 2.1 Resumen

| Script | Objetos extraídos | Observación |
|---|---:|---|
| `1. Estructura.sql` | 126 tablas | Esquema principal INFOREST |
| `4. Vistas.sql` | 116 vistas | Vistas operativas/reportes |
| `5. SP.sql` | 150 stored procedures | Lógica transaccional, reportes e integraciones |

### 2.2 Tablas extraídas de `1. Estructura.sql`

```text
APEDIDO
CPEDIDO
DDOCUMENTO
DPAGODOCUMENTO
DPEDIDO
DPREPAGO
MCIERRE
MDOCUMENTO
MEGRESO
MINGRESO
MNOTACREDITO
MPEDIDO
MTURNO
MGUIATRANSPORTE
DGUIATRANSPORTE
TAREAIMPRESORA
TCLIENTE
TCOMBO
TCOMPANIA
TDELIVERY
TGRUPO
TIMPRESORA
TMESA
TMODULO
TPARAMETRO
TPRODUCTO
TPROPIEDAD
TRESERVA
TSUBGRUPO
TTABLA
TTARJETACREDITO
TTIENDA
TTIPOCAMBIO
TTIPODOCUMENTO
TTIPODOCUMENTOIMPRESORA
TUSUARIO
TCAJA
MENVIO
TPRODUCTOAREA
TOFERTA
MPROPINA
TGRUPOUSUARIO
TMOTIVODESCUENTO
TPRODUCTOPROPIEDAD
TCOMBOPROPIEDAD
TMENSAJE
DPAGOTARJETA
TLOG
TOPERADOR
TACCESO
TGRUPOACCESO
TPEDIDO
TPRODUCTOXPRODUCTO
TPEDIDOMESA
TLOCAL
TACCESOENVIA
DPEDIDOKDS
TORIGENCODIGOCONTROL
TMENSAJECOCINA
TINSUMO
TDIACONTABLE
TAREAPANTALLA
TAREAPANTALLA1
TAREAPANTALLADESPACHO
TTIPOPEDIDODETALLE
TCLIENTEPRODUCTO
TCANALVENTA
TMOTORIZADODATOS
TIMPORTACION
TIMPORTACIONLOG
TTRAMITE
TSOLICITUD
TSOLICITUDDETALLE
TDELIVERYINVITADO
TPARIENTE
TDELIVERYCLIENTE
TAREASUBGRUPO
TAREACHEF
TMENSAJEUSUARIO
TLOG_IMPRESION
TLOG_MODPRECIO
TLISTAESPERA
TUBIGEO
TIMPRESORAIMPRESION
INFOVISOR
LOG_INFOREST
TORIGENVENTA
TCONFIGURAPERIFERICO
VISOR_DPEDIDO
TPRODUCTOSUNAT
TSEGMENTOSUNAT
TCLASESUNAT
TFAMILIASUNAT
TDESCARGOINSUMO
TTIPOIDENTIDAD
TINTEGRACIONES
TDETALLEASISTENCIA
TCAJACANALVENTA
TDISPENSADOR
TTARJETASRFID
TMOVIMIENTOTARJETASRFID
TTIPOMOVIMIENTO
TESTADOBIZLINK
NOTICIAS
HISTORIAL_NOTICIAS
LOG_OPTIMIZACION
LOG_PEDIDO_DOCUMENTO
LOG_SESIONES
ruc_temp
PEDIDO_PAGOEFECTIVO
TPROGRAMAPRECIOS_CAB
TPROGRAMAPRECIOS_DETA
TESTADOINFOFACT
DICTIONARY_INFOREST
DNOTACREDITO
TCENTROCOSTO
VISIBILIDADPROPIEDADXCANAL
INTEGRACION_CASHDRO
TCAJAORIGEN_BLOQUEO
DINGRESO
TMENSAJEIMPRESORA
TVISIBILIDADTARJETACREDITOXCANAL
VALE_CONSUMO
DPAGODOCUMENTO_VC
TTERMINAL
TCAJATERMINAL
```

### 2.3 Vistas extraídas de `4. Vistas.sql`

```text
dbo.vTipoResponsable
dbo.vMotivoAnulacion
vTipoGrupoCliente
vInvitado
VTIPOEGRESO
vOperador
vPropiedad
dbo.vPaloteoProduccionPropiedades
dbo.vPaloteoProduccionPropiedadesCombos
VSector
vFrecuencia
vTablasCentralizada
vTipoCliente
dbo.vEstadoFrecuente
vArea
vSubTipoCtaCte
vTipoCtaCte
vCortesia
vEstadoDocumento
vEstadoMesa
vEstadoPedido
vEstadoGuia
vEstadoReserva
vLocal
vMoneda
vMotorizado
vMozo
vSalon
vTipoAtencion
vTipoDescargo
dbo.vTarjetaCredito
dbo.vTipoDocumento
vTipoIdentidad
vTipoPago
vTipoPedido
vUnidadNegocio
vDistrito
dbo.vTipoClienteFrecuente
vMotivoDescuento
vTipoProducto
vTipodocumentoImpresora
vZona
vAreaImpresora
vCliente
dbo.vCompania
vCtaCte
dbo.vDelivery
vFormulario
vGrupo
vSubGrupo
vProducto
vProductoArea
vCombo
vComboDetalle
dbo.vEgreso
vFacturacionDetalle
vIngreso
vLiquidacion
vNotaCredito
vPreCuenta
vPreCuentaDelivery
vPreCuentaDetallada
vDocumento
vDocumentoImpresora
vDocumentoAgrupado
vDocumentoConsolidado
dbo.vDocumentoGrilla
vDocumentoPago
vDocumentoResultado
vPedido
vPedidoAgrupado
vPedidoCabecera
vPedidoCombo
vPedidoCombox
vPedidoResultado
vPedidoCorrelativo
vPedidoDetalle
vPedidoGrilla
vOferta
vMotivoEliminacion
vTipoCancelacion
vChofer
vVehiculo
vMotivoTraslado
vPrecuentaAgrupada
vGrupoUsuario
vDocumentoImpresoraAgrupado
vEmpacador
vDespachador
vPRODUCTOXPRODUCTO
vDocumentoCorrelativoDetalle
vDocumentoImpresoraAgrupadoAlternativa
vDocumentoImpresoraAlternativa
vPaisOrigen
vCajaCodigoControl
vSucursal
dbo.vMaitre
dbo.vTienda
dbo.vGuiaTransporte
dbo.vEstadoSolicitud
dbo.vEstadoSolicitudDetalle
vPariente
vSectorVenta
vSectorVentaCajaR
vAreaSubGrupo
vAreaChef
vMotivoNotaCredito
vNotaCreditoImpresora
vMotivoReserva
dbo.vOrigenVenta
dbo.vBalanza
dbo.vProductoSunat
dbo.vTPROGRAMAPRECIOS_DETA
dbo.vNotaCreditoSunat
dbo.vDocumentoRegistroVentas
dbo.vDocumentoRegistroVentas_TransGratuita
```

### 2.4 Stored procedures extraídos de `5. SP.sql`

```text
usp_Rec_Consumo
spRep_AnaliticoMozoIntegrado
spIns_MPEDIDO
spIns_TipoCambio
spUpd_MPEDIDO
spRep_Anulacion
spRep_Diferencia
spRep_Rotacion
spRep_Ocupabilidad
usp_InforestCon_ObtenerReporteLiquidacionVentas
spRep_Cortesia
spRep_CtaCteN
spRep_CuentasCobrar
spRep_Descuento
spRep_Liquidacion
spRep_LiquidacionOutPut
spRep_LiquidacionSuma
spRep_PaloteoInsumo
spRep_PaloteoProduccion
spRep_PaloteoPropiedad
spRep_Pedido
spRep_PaloteoOferta
spRep_AnaliticoMotorizado
spRep_AnaliticoMozo
spRep_Cancelacion
spRep_CobranzaFecha
spRep_Ranking
spRep_RegVenta
spRep_RepClieFrecuentes
spRep_ResultadoOperativo
spRep_TiempoDelivery
spRep_TiempoSalon
spRep_VentaCompMensual
spRep_VentaFecha
sp_ActualizaTablas
spRep_FormaPagoIntegrado
spRep_PaloteoInsumoIntegrado
spRep_RegVentaIntegrado
spRep_VentaMensualCanalesIntegrado
spRep_PaloteoProduccionPorMes
sp_CopiaArchivosRemotos
spRep_TiempoDeliveryIntegrado
sp_ComparativoConsumo
spRep_TipoProductoVentaIntegrado
USP_KDS_ResporteTiempoPedido
USP_KDS_ResporteTiempoProducto
spRep_RegVentaSunat
USP_LISTADOMENSAJES
USP_CERRAR_MENSAJES_CIERRETURNO
USP_AGREGARMENSAJE
USP_MODIFICARMENSAJE
USP_LISTARMENSAJES
USP_ELIMINARRMENSAJES
USP_MODIFICARINSUMOS
USP_LISTARINSUMOS
USP_AGREGARINSUMOS
USP_ELIMINARINSUMOS
ups_ObtieneFechaHora
spRep_RegVentaSunatAD
spRep_PlanillaMovilidad
spRep_PlanillaMovilidadGeneral
sp_CargaPedidoReqExternoDirectosInforest
sp_CargaPedidoReqExternoInforest
usp_AdmCen_ObtenerCamposxTablas
usp_AdmCen_ObtenerTablas
usp_AdmCen_ObtenerTablasPrincipales
usp_AdmCen_ObtenerPermisosUsuario
usp_ObtenerTipoDocumento
usp_ObtenerCliente
usp_ObtenerEstadoDocumento
usp_ObtenerCaja
spRep_ControlDocumentos
spRep_ComprobantesVentas
spRep_AutorizacionAutoimpresion
usp_AutSol_ObtenerTipoDocumento
usp_AutSol_ObtenerCliente
usp_AutSol_ObtenerEstadoDocumento
usp_AutSol_ObtenerCaja
usp_AutSol_ObtenerSolicitud
usp_AutSol_ObtenerTramite
usp_AutSol_ObtenerEstadoTramite
sp_TraeDatosVentas_SyBase
sp_TraeDatosPagos_SyBase
sp_InsUptSocioDelivery_SyBASE
sp_CreaTemporalSocio_SyBASE
spRep_VentaIntervaloIntegrado
usp_Inforest_ObtieneClientesFactura
sp_UpdFotoProducto
spRep_AnulacionPedidoIntegrado
spRep_AnulacionDocumentoIntegrado
spRep_MensajeUsuario
sp_UpdImagenCaja
spRep_Entregas
usp_Inforest_ObtienePropiedadesBusqueda
usp_Inforest_ObtienePropiedadesComboBusqueda
spRep_RegVentaDetallado
usp_Inforest_InicializaCostos
usp_Inforest_ActualizaCostos
spRep_ComprobanteDetallado
USP_CALCULA_PRECIO
spIns_DPEDIDO
usp_Inforest_InsertaLogDescargo
spRep_Liquidacion_NC
spRep_LiquidacionOutPut_NC
spRep_LiquidacionSuma_NC
usp_InforestCon_ObtenerReporteLiquidacionVentas_NC
Usp_ActualizarPreciosTransferenciaAlmacen
Usp_GuiaTransporteXml
spRep_LiquidacionOrigenVenta
spRep_RegVentaPagos
spRep_SUNATtxt
spRep_ProductosNoEnlazados
usp_ListDocumentosFE
usp_RepInforest_DescargoVenta
usp_Inforest_Impresion
usp_Inforest_InsertarLogErrores
usp_Inforest_InsertaDocumento
usp_Inforest_DescargoVentaPorInsumo
sp_RepLiquidacion3
usp_Inforest_DescargoVenta
usp_Anticipos
SP_InsertaDescargoInsumo
USP_RD_VALIDANDORUC
usp_Inforest_ListaConectado
usp_Inforest_killListaConec
SP_INFOREST_VALPRODASOCIADOS
USP_ADD_MARCACION
spRep_Asistencia
usp_WS_Inforest
spIns_MPEDIDO_RESERVA
usp_Inforest_ObtieneRecargas
spRep_RegVentaSunat_formaPago
usp_Inforest_ObtieneRptGuia
USP_KDS_ObtenerDetallePedidox
USP_KDS_ObtenerDetallePedidoxKDS
SP_NOTICIAS_INFOREST
usp_Inforest_ObtieneDocumentos_bizlink
SP_EJECUTA_ACTUALIZA_FE
usp_Inforest_ObtieneDocumentos_NC_bizlink
sp_OptimizarBD
usp_ActualizarCabPeDoc
usp_TicketComandaVistaPrevia
spRep_Pedido_GC
USP_RPT_DETA_COMBO
SP_Dictionary_Inforest
sp_AsignaAnticipo_Pedido
SP_NOTACREDITO_PARCIAL
spRep_ControlMotorizado
spRep_ReciboEgreso
usp_ControlServicioWindows
```

## 3. Integraciones detectadas

| Integración | Evidencia | Tecnología Legacy | Alcance inferido |
|---|---|---|---|
| KDS / cocina | `modKDS.bas`, `USP_KDS_ObtenerDetallePedidox`, `frmCheffControl.frm` | XML en directorio + SP SQL | Envío y seguimiento de comandas a cocina |
| Facturación electrónica / InfoFact / Bizlink | `8. InfoFact.sql`, `TESTADOINFOFACT`, `SP_EJECUTA_ACTUALIZA_FE`, `usp_Inforest_ObtieneDocumentos_bizlink` | SQL + integraciones externas | Documentos electrónicos y estados FE |
| PinPad / pagos tarjeta | `DLL3500.bas`, `DLL3500.ini`, `CAJA_PINPAD.dll` | DLL Win32 | Autorización/captura de pagos con tarjeta |
| CashDro | `modProcedimientoNuevo.bas`, tabla `INTEGRACION_CASHDRO`, `frmPago` timer | HTTP/API + polling | Cajón automático de efectivo |
| BlueVision / TVS | `modBlueVision.bas`, referencia `BlueVision_Core_TVS` | COM ActiveX | Visor/display de cliente y señalización |
| Rappi / pedidos externos | `modDespachador.bas`, `frmEnvioDatosRappi.frm`, `sp_CargaPedidoReqExterno*` | SP + integración delivery | Importación/sincronización de pedidos externos |
| Biometría | `FpLibX_Const.bas`, `sgfplibx.ocx`, referencias DigitalPersona | OCX/SDK | Validación de huella y autenticación de usuario |
| Impresora fiscal Epson | `modImpresoraFiscal.bas`, `IFEpson.ocx` | OCX Win32 | Flujo fiscal Argentina |
| Correo / alertas | `claCorreoElectronico.cls`, referencias Chilkat | COM | Alertas, licencias, notificaciones |
| Administración centralizada | `modPuntoVenta.bas`, `INFOREST.INI`, `usp_AdmCen_*` | INI + SQL Server | Sincronización de maestros multi-local |
| Balanza / captura de peso | vista `vBalanza`, flags `TCAJA.nBalanzaPuerto`, `lCapturaPeso` | Puerto serial / hardware POS | Pesaje y captura de productos por peso |

## 4. Hardware POS detectado

| Hardware | Evidencia | Observación de migración |
|---|---|---|
| Impresoras térmicas / comandas | `TIMPRESORA`, `TTIPODOCUMENTOIMPRESORA`, `frmImpresora*.frm` | Requiere abstracción de impresión .NET 8 |
| Cajón de dinero | `frmCajon.frm`, flujo caja POS, hardware POS Windows | Debe mantenerse soporte de apertura de cajón |
| CashDro | `INTEGRACION_CASHDRO`, `modProcedimientoNuevo.bas` | Integración propietaria / polling |
| PinPad DLL3500 | `CAJA_PINPAD.dll`, `DLL3500.bas` | Dependencia Win32 crítica |
| Impresora fiscal Epson | `IFEpson.ocx`, `modImpresoraFiscal.bas` | Requisito fiscal específico de Argentina |
| Biometría SecuGen / DigitalPersona | `sgfplibx.ocx`, referencias DigitalPersona | Validación de huella y marcación |
| KDS / pantallas cocina | `modKDS.bas`, `frmCheffControl.frm` | Interfaz operativa de cocina |
| BlueVision / visor cliente | `modBlueVision.bas`, `FrmMensajeVisor.frm` | Dispositivo/display externo |
| Balanza | `vBalanza`, `nBalanzaPuerto`, `lCapturaPeso` | Captura de peso por puerto configurado |
| Tarjetas RFID / proximidad | `FrmTarjetaAproximidad.frm`, `TTARJETASRFID`, `TMOVIMIENTOTARJETASRFID` | Identificación/carga de saldo con tarjetas |
| Lectores de código de barras | `tBarra` en `TPRODUCTO`, `frmMsComm.frm`, `frmWinSock.frm` | Entrada de productos/etiquetas en POS |

## 5. Reportes Crystal Reports

- Reportes `.Dsr`: **206**
- Artefactos `.DCA`: **206**
- Artefactos `.dsx`: **205**
- Otros artefactos temporales/log en `Reportes/`: **23** (`.log`/`.tmp`)

### 5.1 Lista completa de reportes `.Dsr`

```text
CrtComprasxTipoDocumentoIntegrado.Dsr
CrtComprasxTipoProductoIntegrado.Dsr
CrtInsSolicitadosVsEntregadoDetallado.Dsr
CrtInsSolicitadosVsEntregadoResumido.Dsr
CrtRegComprasIntegrado.Dsr
CrtRepInventarioIntegrado.Dsr
CrtRepMovimientosDetalladoIntegrado.Dsr
CrtRepMovimientosIntegrado.Dsr
DataReport1.Dsr
dsrAnaliticoMotorizado.Dsr
dsrAnaliticoMotorizadoC.Dsr
dsrAnaliticoMotorizadoIntegradoConsolidado.Dsr
dsrAnaliticoMotorizadoIntegradoConsolidadoRES.Dsr
dsrAnaliticoMotorizadoIntegradoConsolidadoXM.Dsr
dsrAnaliticoMotorizadoIntegradoDetallado.Dsr
dsrAnaliticoMotorizadoIntegradoDetalladoXM.Dsr
dsrAnaliticoMotorizador.Dsr
dsrAnaliticoMozo.Dsr
dsrAnaliticoMozoC.Dsr
dsrAnticipoAmbos.Dsr
dsrAnticipoAmbosConsumos.Dsr
dsrAnuladoD.Dsr
dsrAnuladoR.Dsr
dsrArea.Dsr
dsrAsistencia.Dsr
dsrAsistencia2.Dsr
dsrBoleta.Dsr
dsrBoletaC.Dsr
dsrCaja.Dsr
dsrCancelacionD.Dsr
dsrCancelacionR.Dsr
dsrCliente.Dsr
dsrClienteCtaCte.Dsr
dsrClienteDelivery.Dsr
dsrClienteDeudaD.Dsr
dsrClienteDeudaHistorico.Dsr
dsrClienteDeudaR.Dsr
dsrClienteFactura.Dsr
dsrCobranzaFecha.Dsr
dsrCobranzaMensualHorizontal.Dsr
dsrComandaD.Dsr
dsrComandaR.Dsr
dsrCompConsumoInsumos.Dsr
dsrCompConsumoInsumosValor.Dsr
dsrComparativoConsumoInsumos.Dsr
dsrComparativoConsumoInsumosValor.Dsr
dsrContacto.Dsr
dsrControlEnviosAutorizados.Dsr
dsrControlMotorizado.Dsr
dsrCortesia.Dsr
dsrCortesiaA.Dsr
dsrCortesiaAR.Dsr
dsrCortesiaD.Dsr
dsrCortesiaR.Dsr
dsrCtaCteC.Dsr
dsrCtaCteD.Dsr
dsrCtaCteR.Dsr
dsrCuentaCobrar.Dsr
dsrDeliveryDetallado.Dsr
dsrDeliveryDetalladoxDia.Dsr
dsrDeliveryGrupo.Dsr
dsrDeliveryTiempo.Dsr
dsrDeliveryTiempoCD.Dsr
dsrDescuento.Dsr
dsrDescuentoA.Dsr
dsrDescuentoD.Dsr
dsrDiferencia.Dsr
dsrDiferenciaTiempo.Dsr
dsrFormaPagoDiarioIntegradoLocales.Dsr
dsrFormaPagoDiarioIntegradoSector.Dsr
dsrFormaPagoIntegrado.Dsr
dsrGuia.Dsr
dsrGuiaCorrelativo.Dsr
dsrImpresora.Dsr
dsrInsumosCD.Dsr
dsrInsumosD.Dsr
dsrInsumosIntegrados.Dsr
dsrInsumosR.Dsr
dsrInventarioConsolidado.Dsr
dsrLiquidacion.Dsr
dsrLiquidacion3.Dsr
dsrLiquidacionVenta.Dsr
dsrMensajeUsuario.Dsr
dsrMotivoDescuento.Dsr
dsrMovimiento.Dsr
dsrMozoProduccion.Dsr
dsrNewFormatoGuia.Dsr
dsrNotaCredito.Dsr
dsrNotaCreditoObservacion.Dsr
dsrOcupabilidadD.Dsr
dsrOcupabilidadR.Dsr
dsrOferta.Dsr
dsrPaloteoA.Dsr
dsrPaloteoC.Dsr
dsrPaloteoC2.Dsr
dsrPaloteoCR.Dsr
dsrPaloteoCResumido.Dsr
dsrPaloteoComparativo.Dsr
dsrPaloteoD.Dsr
dsrPaloteoItemCombo.Dsr
dsrPaloteoOfertaD.Dsr
dsrPaloteoOfertaP.Dsr
dsrPaloteoOfertaR.Dsr
dsrPaloteoP.Dsr
dsrPaloteoPP.Dsr
dsrPaloteoPropiedadA.Dsr
dsrPaloteoPropiedadD.Dsr
dsrPaloteoPropiedadG.Dsr
dsrPaloteoPropiedadP.Dsr
dsrPaloteoR.Dsr
dsrPaloteoSPD.Dsr
dsrPaloteoVentaIntegradoComparativo.Dsr
dsrPaloteoVentaIntegradoS.Dsr
dsrPaloteoVentaIntegradol.Dsr
dsrPedidosD.Dsr
dsrPedidosFF.Dsr
dsrPedidosFFyKDS.Dsr
dsrPedidosR.Dsr
dsrPedidosX.Dsr
dsrPlanillaMovilidadGeneral.Dsr
dsrPlanillaMovilidadPorDiaDetallado.Dsr
dsrPlanillaMovilidadPorDiaResumido.Dsr
dsrPlato.Dsr
dsrPlatoCombo.Dsr
dsrPrecios.Dsr
dsrPrincipalD.Dsr
dsrPrincipalR.Dsr
dsrProduccionMozo.Dsr
dsrProducto.Dsr
dsrProductoNoEnlazados.Dsr
dsrProductoXMes.Dsr
dsrProductoXMesMonto.Dsr
dsrProductoxproducto.Dsr
dsrPropiedad.Dsr
dsrPropiedad2.Dsr
dsrPropinaD.Dsr
dsrPropinaMaitre.Dsr
dsrPropinaR.Dsr
dsrRankingD.Dsr
dsrRankingIntegradoComparativo.Dsr
dsrRankingIntegradol.Dsr
dsrRankingR.Dsr
dsrRecibo.Dsr
dsrReciboEgreso.Dsr
dsrReciboEgresoCompra.Dsr
dsrRegistroVentaC.Dsr
dsrRegistroVentaD.Dsr
dsrRegistroVentaD2.Dsr
dsrRegistroVentaE.Dsr
dsrRegistroVentaF.Dsr
dsrRegistroVentaIntegradoE.Dsr
dsrRegistroVentaIntegradoL.Dsr
dsrRegistroVentaIntegradoS.Dsr
dsrRegistroVentaR.Dsr
dsrRegistroVentaSunat.Dsr
dsrRegistroVentaSunatAD.Dsr
dsrRegistroVentaSunatPago.Dsr
dsrRepCtaCteIntegradoConsolidado.Dsr
dsrRepCtaCteIntegradoConsolidadoXC.Dsr
dsrRepCtaCteIntegradoDetallado.Dsr
dsrRepCtaCteIntegradoDetalladoXC.Dsr
dsrRepDescargo.Dsr
dsrRepDescargoResumido.Dsr
dsrRepInsumos.Dsr
dsrRepInsumosPtoVenta.Dsr
dsrRepNotaCredito.Dsr
dsrReporte.Dsr
dsrReporte4.Dsr
dsrReporte5.Dsr
dsrReporte6.Dsr
dsrReporte7.Dsr
dsrReporteEmpacador.Dsr
dsrReporteEntrega.Dsr
dsrReporteEntregaFormato2.Dsr
dsrReporteEntregaRegistro.Dsr
dsrReporteEntregaResumidoProd.Dsr
dsrReporteMensajeCocina.Dsr
dsrReporteTipoCambio.Dsr
dsrReprecargoconsumo.Dsr
dsrReprecargoconsumodet.Dsr
dsrReserva.Dsr
dsrReservas.Dsr
dsrResultadosOperativos.Dsr
dsrRotacionMesa.Dsr
dsrSemana.Dsr
dsrSubGrupo.Dsr
dsrTiempoChefDeta.Dsr
dsrTiempoChefResumido.Dsr
dsrTiempoDeliveryIntegrado.Dsr
dsrTiempoKdsP.Dsr
dsrTiempoKdsR.Dsr
dsrTiempoSalon.Dsr
dsrTipoProductoVentaIntegrado.Dsr
dsrVentaAcumulada.Dsr
dsrVentaAnualGrafico.Dsr
dsrVentaCompAnual.Dsr
dsrVentaCompMensualDia.Dsr
dsrVentaDiariaIntegrada.Dsr
dsrVentaDiariaIntegradaMetas.Dsr
dsrVentaDiariaIntegradaS.Dsr
dsrVentaFecha.Dsr
dsrVentaMensualCanalesIntegrada.Dsr
dsrVentaMensualIntegrada.Dsr
dsrVentaMensualIntegradaS.Dsr
dsrVentaTurno.Dsr
xxx.Dsr
```

---

*Documento generado a partir de listados de archivos del repositorio y parsing directo de scripts SQL.*
