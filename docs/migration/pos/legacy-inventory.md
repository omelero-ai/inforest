# Punto de Venta — Legacy Inventory (InfoRest.exe)

> Fuente: `legacy-restaurant/restaurant-vb6/InfoRest.vbp` + componentes declarados directos.  \
> Fecha de auditoría: 2026-08-13.

## Alcance exclusivo POS

- Ejecutable auditado: `InfoRest.exe`
- Proyecto VB6 auditado: `InfoRest.vbp`
- Regla de alcance: no se migra automáticamente otros ejecutables; solo se documentan como dependencia.
- **Nueva restricción aceptada:** `IFEpson.ocx` / impresora fiscal Epson queda **EXCLUIDA** del alcance POS actual (Argentina opera FE); no debe usarse en nuevos componentes .NET mientras no exista aprobación explícita.

## Resumen cuantitativo

| Métrica | Valor | Evidencia |
|---|---:|---|
| Formularios `.frm` declarados en `InfoRest.vbp` | 120 | `Form=` |
| Recursos `.frx` asociados encontrados | 110 | `Formularios/*.frx` |
| Módulos `.bas` declarados | 18 | `Module=` |
| Clases `.cls` declaradas | 10 | `Class=` |
| Diseñadores de reportes `.Dsr` | 21 | `Designer=` |
| Referencias COM (`Reference=`) | 20 | `InfoRest.vbp` |
| Objetos ActiveX/OCX (`Object=`) | 19 | `InfoRest.vbp` |
| Componentes inventariados en tabla (total) | 296 | Tabla completa inferior |

## Flujo ejecutable (Startup Sub Main)

1. Valida instancia única (`App.PrevInstance`) y aborta duplicados.
2. Carga configuración de `INFOREST.INI` (servidor, BD, caja, salón, empresa, flags POS).
3. Inicializa conexiones SQL (`INFSEGURIDAD`, `INFOREST`, centralizada opcional).
4. Lee `TPARAMETRO` y carga masivamente flags globales de operación.
5. Ejecuta validación de versión/actualizador + validación de licencias.
6. Inicializa conexiones condicionadas (`FACTURACION.INI`, `ALMACEN.INI`, `INFHOTEL.INI`).
7. Carga configuración por caja desde `TCAJA` (80+ flags operativos).
8. Valida correlativos (`TTIPODOCUMENTOIMPRESORA`) y facturación electrónica TCP/IP cuando aplica.
9. Presenta `frmAcceso` (o bypass con `lMCPV`) y abre `mdiPuntoVenta`.
10. Maneja errores de conectividad con etiquetas `InforestIni`, `FacturacionIni`, `AlmacenIni`, `InfhotelIni`.

## Inventario completo de componentes POS (InfoRest.vbp)

| ID | Componente | Tipo | Archivo | Responsabilidad | Dependencias | Estado |
|---:|---|---|---|---|---|---|
| 1 | `Formularios\frmAbout.frm` | FORM | `Formularios\frmAbout.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 2 | `frmAbout.frx` | FRX | `Formularios/frmAbout.frx` | Recurso binario del formulario Formularios\frmAbout.frm | Formularios\frmAbout.frm | USED |
| 3 | `Formularios\frmAcceso.frm` | FORM | `Formularios\frmAcceso.frm` | Autenticación y seguridad | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 4 | `frmAcceso.frx` | FRX | `Formularios/frmAcceso.frx` | Recurso binario del formulario Formularios\frmAcceso.frm | Formularios\frmAcceso.frm | USED |
| 5 | `Formularios\frmLiquidacionDetalle.frm` | FORM | `Formularios\frmLiquidacionDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 6 | `frmLiquidacionDetalle.frx` | FRX | `Formularios/frmLiquidacionDetalle.frx` | Recurso binario del formulario Formularios\frmLiquidacionDetalle.frm | Formularios\frmLiquidacionDetalle.frm | USED |
| 7 | `Formularios\frmEmite.frm` | FORM | `Formularios\frmEmite.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 8 | `frmEmite.frx` | FRX | `Formularios/frmEmite.frx` | Recurso binario del formulario Formularios\frmEmite.frm | Formularios\frmEmite.frm | USED |
| 9 | `Formularios\frmFlash.frm` | FORM | `Formularios\frmFlash.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 10 | `frmFlash.frx` | FRX | `Formularios/frmFlash.frx` | Recurso binario del formulario Formularios\frmFlash.frm | Formularios\frmFlash.frm | USED |
| 11 | `Formularios\frmGeneracion.frm` | FORM | `Formularios\frmGeneracion.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 12 | `frmGeneracion.frx` | FRX | `Formularios/frmGeneracion.frx` | Recurso binario del formulario Formularios\frmGeneracion.frm | Formularios\frmGeneracion.frm | USED |
| 13 | `Formularios\frmInicio.frm` | FORM | `Formularios\frmInicio.frm` | Configuración del sistema | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 14 | `frmInicio.frx` | FRX | `Formularios/frmInicio.frx` | Recurso binario del formulario Formularios\frmInicio.frm | Formularios\frmInicio.frm | USED |
| 15 | `Formularios\frmIr.frm` | FORM | `Formularios\frmIr.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 16 | `frmIr.frx` | FRX | `Formularios/frmIr.frx` | Recurso binario del formulario Formularios\frmIr.frm | Formularios\frmIr.frm | USED |
| 17 | `Formularios\frmPrograma.frm` | FORM | `Formularios\frmPrograma.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 18 | `frmPrograma.frx` | FRX | `Formularios/frmPrograma.frx` | Recurso binario del formulario Formularios\frmPrograma.frm | Formularios\frmPrograma.frm | USED |
| 19 | `Formularios\mdiPuntoVenta.frm` | FORM | `Formularios\mdiPuntoVenta.frm` | Operación de venta POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 20 | `mdiPuntoVenta.frx` | FRX | `Formularios/mdiPuntoVenta.frx` | Recurso binario del formulario Formularios\mdiPuntoVenta.frm | Formularios\mdiPuntoVenta.frm | USED |
| 21 | `Formularios\frmCuentaCobrar.frm` | FORM | `Formularios\frmCuentaCobrar.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 22 | `frmCuentaCobrar.frx` | FRX | `Formularios/frmCuentaCobrar.frx` | Recurso binario del formulario Formularios\frmCuentaCobrar.frm | Formularios\frmCuentaCobrar.frm | USED |
| 23 | `Formularios\frmCuentaCobrarDetalle.frm` | FORM | `Formularios\frmCuentaCobrarDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 24 | `frmCuentaCobrarDetalle.frx` | FRX | `Formularios/frmCuentaCobrarDetalle.frx` | Recurso binario del formulario Formularios\frmCuentaCobrarDetalle.frm | Formularios\frmCuentaCobrarDetalle.frm | USED |
| 25 | `Formularios\frmJuntaMesa.frm` | FORM | `Formularios\frmJuntaMesa.frm` | Gestión de mesas/salón | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 26 | `frmJuntaMesa.frx` | FRX | `Formularios/frmJuntaMesa.frx` | Recurso binario del formulario Formularios\frmJuntaMesa.frm | Formularios\frmJuntaMesa.frm | USED |
| 27 | `Formularios\frmFactura.frm` | FORM | `Formularios\frmFactura.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 28 | `frmFactura.frx` | FRX | `Formularios/frmFactura.frx` | Recurso binario del formulario Formularios\frmFactura.frm | Formularios\frmFactura.frm | USED |
| 29 | `Formularios\frmCambio.frm` | FORM | `Formularios\frmCambio.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 30 | `frmCambio.frx` | FRX | `Formularios/frmCambio.frx` | Recurso binario del formulario Formularios\frmCambio.frm | Formularios\frmCambio.frm | USED |
| 31 | `Formularios\frmPedidoAnterior.frm` | FORM | `Formularios\frmPedidoAnterior.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 32 | `frmPedidoAnterior.frx` | FRX | `Formularios/frmPedidoAnterior.frx` | Recurso binario del formulario Formularios\frmPedidoAnterior.frm | Formularios\frmPedidoAnterior.frm | USED |
| 33 | `Formularios\frmReserva.frm` | FORM | `Formularios\frmReserva.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 34 | `frmReserva.frx` | FRX | `Formularios/frmReserva.frx` | Recurso binario del formulario Formularios\frmReserva.frm | Formularios\frmReserva.frm | USED |
| 35 | `Formularios\frmDocumentoPago.frm` | FORM | `Formularios\frmDocumentoPago.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 36 | `frmDocumentoPago.frx` | FRX | `Formularios/frmDocumentoPago.frx` | Recurso binario del formulario Formularios\frmDocumentoPago.frm | Formularios\frmDocumentoPago.frm | USED |
| 37 | `Formularios\frmPedidoCorrelativo.frm` | FORM | `Formularios\frmPedidoCorrelativo.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 38 | `frmPedidoCorrelativo.frx` | FRX | `Formularios/frmPedidoCorrelativo.frx` | Recurso binario del formulario Formularios\frmPedidoCorrelativo.frm | Formularios\frmPedidoCorrelativo.frm | USED |
| 39 | `Formularios\frmCtaCteDetalle.frm` | FORM | `Formularios\frmCtaCteDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 40 | `frmCtaCteDetalle.frx` | FRX | `Formularios/frmCtaCteDetalle.frx` | Recurso binario del formulario Formularios\frmCtaCteDetalle.frm | Formularios\frmCtaCteDetalle.frm | USED |
| 41 | `Formularios\frmPrecios.frm` | FORM | `Formularios\frmPrecios.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 42 | `frmPrecios.frx` | FRX | `Formularios/frmPrecios.frx` | Recurso binario del formulario Formularios\frmPrecios.frm | Formularios\frmPrecios.frm | USED |
| 43 | `Formularios\frmNuevoCliente.frm` | FORM | `Formularios\frmNuevoCliente.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 44 | `frmNuevoCliente.frx` | FRX | `Formularios/frmNuevoCliente.frx` | Recurso binario del formulario Formularios\frmNuevoCliente.frm | Formularios\frmNuevoCliente.frm | USED |
| 45 | `Formularios\frmDocumentoCorrelativo.frm` | FORM | `Formularios\frmDocumentoCorrelativo.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 46 | `frmDocumentoCorrelativo.frx` | FRX | `Formularios/frmDocumentoCorrelativo.frx` | Recurso binario del formulario Formularios\frmDocumentoCorrelativo.frm | Formularios\frmDocumentoCorrelativo.frm | USED |
| 47 | `Formularios\frmGuiaTransporteCorrelativoDetalle.frm` | FORM | `Formularios\frmGuiaTransporteCorrelativoDetalle.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 48 | `frmGuiaTransporteCorrelativoDetalle.frx` | FRX | `Formularios/frmGuiaTransporteCorrelativoDetalle.frx` | Recurso binario del formulario Formularios\frmGuiaTransporteCorrelativoDetalle.frm | Formularios\frmGuiaTransporteCorrelativoDetalle.frm | USED |
| 49 | `Formularios\frmRepCtaCte.frm` | FORM | `Formularios\frmRepCtaCte.frm` | Salida de reportes | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 50 | `frmRepCtaCte.frx` | FRX | `Formularios/frmRepCtaCte.frx` | Recurso binario del formulario Formularios\frmRepCtaCte.frm | Formularios\frmRepCtaCte.frm | USED |
| 51 | `Formularios\frmKeyboard.frm` | FORM | `Formularios\frmKeyboard.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 52 | `frmKeyboard.frx` | FRX | `Formularios/frmKeyboard.frx` | Recurso binario del formulario Formularios\frmKeyboard.frm | Formularios\frmKeyboard.frm | USED |
| 53 | `Formularios\frmPassword.frm` | FORM | `Formularios\frmPassword.frm` | Autenticación y seguridad | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 54 | `frmPassword.frx` | FRX | `Formularios/frmPassword.frx` | Recurso binario del formulario Formularios\frmPassword.frm | Formularios\frmPassword.frm | USED |
| 55 | `Formularios\frmReservaDetalle.frm` | FORM | `Formularios\frmReservaDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 56 | `frmReservaDetalle.frx` | FRX | `Formularios/frmReservaDetalle.frx` | Recurso binario del formulario Formularios\frmReservaDetalle.frm | Formularios\frmReservaDetalle.frm | USED |
| 57 | `Formularios\frmCajaRapida.frm` | FORM | `Formularios\frmCajaRapida.frm` | Operación de caja | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 58 | `frmCajaRapida.frx` | FRX | `Formularios/frmCajaRapida.frx` | Recurso binario del formulario Formularios\frmCajaRapida.frm | Formularios\frmCajaRapida.frm | USED |
| 59 | `Formularios\frmGuiaTransporte.frm` | FORM | `Formularios\frmGuiaTransporte.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 60 | `frmGuiaTransporte.frx` | FRX | `Formularios/frmGuiaTransporte.frx` | Recurso binario del formulario Formularios\frmGuiaTransporte.frm | Formularios\frmGuiaTransporte.frm | USED |
| 61 | `Formularios\frmNotaCredito.frm` | FORM | `Formularios\frmNotaCredito.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 62 | `frmNotaCredito.frx` | FRX | `Formularios/frmNotaCredito.frx` | Recurso binario del formulario Formularios\frmNotaCredito.frm | Formularios\frmNotaCredito.frm | USED |
| 63 | `Formularios\frmNotaCreditoDetalle.frm` | FORM | `Formularios\frmNotaCreditoDetalle.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 64 | `frmNotaCreditoDetalle.frx` | FRX | `Formularios/frmNotaCreditoDetalle.frx` | Recurso binario del formulario Formularios\frmNotaCreditoDetalle.frm | Formularios\frmNotaCreditoDetalle.frm | USED |
| 65 | `Formularios\frmPrePago.frm` | FORM | `Formularios\frmPrePago.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 66 | `frmPrePago.frx` | FRX | `Formularios/frmPrePago.frx` | Recurso binario del formulario Formularios\frmPrePago.frm | Formularios\frmPrePago.frm | USED |
| 67 | `Formularios\frmReciboIngresoDetalle.frm` | FORM | `Formularios\frmReciboIngresoDetalle.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 68 | `frmReciboIngresoDetalle.frx` | FRX | `Formularios/frmReciboIngresoDetalle.frx` | Recurso binario del formulario Formularios\frmReciboIngresoDetalle.frm | Formularios\frmReciboIngresoDetalle.frm | USED |
| 69 | `Formularios\frmReciboIngreso.frm` | FORM | `Formularios\frmReciboIngreso.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 70 | `frmReciboIngreso.frx` | FRX | `Formularios/frmReciboIngreso.frx` | Recurso binario del formulario Formularios\frmReciboIngreso.frm | Formularios\frmReciboIngreso.frm | USED |
| 71 | `Formularios\frmDivision.frm` | FORM | `Formularios\frmDivision.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 72 | `frmDivision.frx` | FRX | `Formularios/frmDivision.frx` | Recurso binario del formulario Formularios\frmDivision.frm | Formularios\frmDivision.frm | USED |
| 73 | `Formularios\frmPedidoDetalle.frm` | FORM | `Formularios\frmPedidoDetalle.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 74 | `frmPedidoDetalle.frx` | FRX | `Formularios/frmPedidoDetalle.frx` | Recurso binario del formulario Formularios\frmPedidoDetalle.frm | Formularios\frmPedidoDetalle.frm | USED |
| 75 | `Formularios\frmBusquedaDelivery.frm` | FORM | `Formularios\frmBusquedaDelivery.frm` | Flujos delivery/dispatch dependientes del POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 76 | `frmBusquedaDelivery.frx` | FRX | `Formularios/frmBusquedaDelivery.frx` | Recurso binario del formulario Formularios\frmBusquedaDelivery.frm | Formularios\frmBusquedaDelivery.frm | USED |
| 77 | `Formularios\frmPago.frm` | FORM | `Formularios\frmPago.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 78 | `frmPago.frx` | FRX | `Formularios/frmPago.frx` | Recurso binario del formulario Formularios\frmPago.frm | Formularios\frmPago.frm | USED |
| 79 | `Formularios\frmVenta.frm` | FORM | `Formularios\frmVenta.frm` | Operación de venta POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 80 | `frmVenta.frx` | FRX | `Formularios/frmVenta.frx` | Recurso binario del formulario Formularios\frmVenta.frm | Formularios\frmVenta.frm | USED |
| 81 | `Formularios\frmMesaConsulta.frm` | FORM | `Formularios\frmMesaConsulta.frm` | Gestión de mesas/salón | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 82 | `frmMesaConsulta.frx` | FRX | `Formularios/frmMesaConsulta.frx` | Recurso binario del formulario Formularios\frmMesaConsulta.frm | Formularios\frmMesaConsulta.frm | USED |
| 83 | `Formularios\frmPedido.frm` | FORM | `Formularios\frmPedido.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 84 | `frmPedido.frx` | FRX | `Formularios/frmPedido.frx` | Recurso binario del formulario Formularios\frmPedido.frm | Formularios\frmPedido.frm | USED |
| 85 | `Formularios\frmRepLiquidacionTicket.frm` | FORM | `Formularios\frmRepLiquidacionTicket.frm` | Salida de reportes | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 86 | `frmRepLiquidacionTicket.frx` | FRX | `Formularios/frmRepLiquidacionTicket.frx` | Recurso binario del formulario Formularios\frmRepLiquidacionTicket.frm | Formularios\frmRepLiquidacionTicket.frm | USED |
| 87 | `Formularios\frmRepPaloteoTicket.frm` | FORM | `Formularios\frmRepPaloteoTicket.frm` | Salida de reportes | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 88 | `frmRepPaloteoTicket.frx` | FRX | `Formularios/frmRepPaloteoTicket.frx` | Recurso binario del formulario Formularios\frmRepPaloteoTicket.frm | Formularios\frmRepPaloteoTicket.frm | USED |
| 89 | `Formularios\frmAreasImpresora.frm` | FORM | `Formularios\frmAreasImpresora.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 90 | `frmAreasImpresora.frx` | FRX | `Formularios/frmAreasImpresora.frx` | Recurso binario del formulario Formularios\frmAreasImpresora.frm | Formularios\frmAreasImpresora.frm | USED |
| 91 | `Formularios\frmCargo.frm` | FORM | `Formularios\frmCargo.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 92 | `frmCargo.frx` | FRX | `Formularios/frmCargo.frx` | Recurso binario del formulario Formularios\frmCargo.frm | Formularios\frmCargo.frm | USED |
| 93 | `Formularios\frmNuevoDelivery.frm` | FORM | `Formularios\frmNuevoDelivery.frm` | Flujos delivery/dispatch dependientes del POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 94 | `frmNuevoDelivery.frx` | FRX | `Formularios/frmNuevoDelivery.frx` | Recurso binario del formulario Formularios\frmNuevoDelivery.frm | Formularios\frmNuevoDelivery.frm | USED |
| 95 | `Formularios\frmDocumento.frm` | FORM | `Formularios\frmDocumento.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 96 | `frmDocumento.frx` | FRX | `Formularios/frmDocumento.frx` | Recurso binario del formulario Formularios\frmDocumento.frm | Formularios\frmDocumento.frm | USED |
| 97 | `Formularios\frmCambioPropina.frm` | FORM | `Formularios\frmCambioPropina.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 98 | `frmCambioPropina.frx` | FRX | `Formularios/frmCambioPropina.frx` | Recurso binario del formulario Formularios\frmCambioPropina.frm | Formularios\frmCambioPropina.frm | USED |
| 99 | `Formularios\frmBusquedaRapida.frm` | FORM | `Formularios\frmBusquedaRapida.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 100 | `frmBusquedaRapida.frx` | FRX | `Formularios/frmBusquedaRapida.frx` | Recurso binario del formulario Formularios\frmBusquedaRapida.frm | Formularios\frmBusquedaRapida.frm | USED |
| 101 | `Formularios\frmBusca.frm` | FORM | `Formularios\frmBusca.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 102 | `frmBusca.frx` | FRX | `Formularios/frmBusca.frx` | Recurso binario del formulario Formularios\frmBusca.frm | Formularios\frmBusca.frm | USED |
| 103 | `Formularios\frmMsComm.frm` | FORM | `Formularios\frmMsComm.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 104 | `Formularios\frmMensaje.frm` | FORM | `Formularios\frmMensaje.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 105 | `frmMensaje.frx` | FRX | `Formularios/frmMensaje.frx` | Recurso binario del formulario Formularios\frmMensaje.frm | Formularios\frmMensaje.frm | USED |
| 106 | `Formularios\frmWinSock.frm` | FORM | `Formularios\frmWinSock.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 107 | `Formularios\frmMsgBox.frm` | FORM | `Formularios\frmMsgBox.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 108 | `frmMsgBox.frx` | FRX | `Formularios/frmMsgBox.frx` | Recurso binario del formulario Formularios\frmMsgBox.frm | Formularios\frmMsgBox.frm | USED |
| 109 | `Formularios\frmPedidoDelivery.frm` | FORM | `Formularios\frmPedidoDelivery.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 110 | `frmPedidoDelivery.frx` | FRX | `Formularios/frmPedidoDelivery.frx` | Recurso binario del formulario Formularios\frmPedidoDelivery.frm | Formularios\frmPedidoDelivery.frm | USED |
| 111 | `Formularios\frmPedidoDeliveryNo.frm` | FORM | `Formularios\frmPedidoDeliveryNo.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 112 | `frmPedidoDeliveryNo.frx` | FRX | `Formularios/frmPedidoDeliveryNo.frx` | Recurso binario del formulario Formularios\frmPedidoDeliveryNo.frm | Formularios\frmPedidoDeliveryNo.frm | USED |
| 113 | `Formularios\frmCtaCte.frm` | FORM | `Formularios\frmCtaCte.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 114 | `frmCtaCte.frx` | FRX | `Formularios/frmCtaCte.frx` | Recurso binario del formulario Formularios\frmCtaCte.frm | Formularios\frmCtaCte.frm | USED |
| 115 | `Formularios\frmPreciosDetalle.frm` | FORM | `Formularios\frmPreciosDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 116 | `frmPreciosDetalle.frx` | FRX | `Formularios/frmPreciosDetalle.frx` | Recurso binario del formulario Formularios\frmPreciosDetalle.frm | Formularios\frmPreciosDetalle.frm | USED |
| 117 | `Formularios\frmPrecuentaImpresora.frm` | FORM | `Formularios\frmPrecuentaImpresora.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 118 | `frmPrecuentaImpresora.frx` | FRX | `Formularios/frmPrecuentaImpresora.frx` | Recurso binario del formulario Formularios\frmPrecuentaImpresora.frm | Formularios\frmPrecuentaImpresora.frm | USED |
| 119 | `Formularios\frmRepDeliveryTicket.frm` | FORM | `Formularios\frmRepDeliveryTicket.frm` | Flujos delivery/dispatch dependientes del POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 120 | `frmRepDeliveryTicket.frx` | FRX | `Formularios/frmRepDeliveryTicket.frx` | Recurso binario del formulario Formularios\frmRepDeliveryTicket.frm | Formularios\frmRepDeliveryTicket.frm | USED |
| 121 | `Formularios\frmMesa.frm` | FORM | `Formularios\frmMesa.frm` | Gestión de mesas/salón | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 122 | `frmMesa.frx` | FRX | `Formularios/frmMesa.frx` | Recurso binario del formulario Formularios\frmMesa.frm | Formularios\frmMesa.frm | USED |
| 123 | `Formularios\frmRepReservas.frm` | FORM | `Formularios\frmRepReservas.frm` | Salida de reportes | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 124 | `frmRepReservas.frx` | FRX | `Formularios/frmRepReservas.frx` | Recurso binario del formulario Formularios\frmRepReservas.frm | Formularios\frmRepReservas.frm | USED |
| 125 | `Formularios\frmMozoUsuario.frm` | FORM | `Formularios\frmMozoUsuario.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 126 | `frmMozoUsuario.frx` | FRX | `Formularios/frmMozoUsuario.frx` | Recurso binario del formulario Formularios\frmMozoUsuario.frm | Formularios\frmMozoUsuario.frm | USED |
| 127 | `Formularios\frmCargoMozo.frm` | FORM | `Formularios\frmCargoMozo.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 128 | `frmCargoMozo.frx` | FRX | `Formularios/frmCargoMozo.frx` | Recurso binario del formulario Formularios\frmCargoMozo.frm | Formularios\frmCargoMozo.frm | USED |
| 129 | `Formularios\frmServidorEnlace.frm` | FORM | `Formularios\frmServidorEnlace.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 130 | `frmServidorEnlace.frx` | FRX | `Formularios/frmServidorEnlace.frx` | Recurso binario del formulario Formularios\frmServidorEnlace.frm | Formularios\frmServidorEnlace.frm | USED |
| 131 | `Formularios\frmFiltroRecibo.frm` | FORM | `Formularios\frmFiltroRecibo.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 132 | `frmFiltroRecibo.frx` | FRX | `Formularios/frmFiltroRecibo.frx` | Recurso binario del formulario Formularios\frmFiltroRecibo.frm | Formularios\frmFiltroRecibo.frm | USED |
| 133 | `Formularios\frmCrearIniTVS.frm` | FORM | `Formularios\frmCrearIniTVS.frm` | Configuración del sistema | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 134 | `Formularios\frmPagoRapido.frm` | FORM | `Formularios\frmPagoRapido.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 135 | `frmPagoRapido.frx` | FRX | `Formularios/frmPagoRapido.frx` | Recurso binario del formulario Formularios\frmPagoRapido.frm | Formularios\frmPagoRapido.frm | USED |
| 136 | `Formularios\frmMensajeCocina.frm` | FORM | `Formularios\frmMensajeCocina.frm` | Producción/cocina | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 137 | `frmMensajeCocina.frx` | FRX | `Formularios/frmMensajeCocina.frx` | Recurso binario del formulario Formularios\frmMensajeCocina.frm | Formularios\frmMensajeCocina.frm | USED |
| 138 | `Formularios\frmMensajeCocinaDetalle.frm` | FORM | `Formularios\frmMensajeCocinaDetalle.frm` | Producción/cocina | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 139 | `frmMensajeCocinaDetalle.frx` | FRX | `Formularios/frmMensajeCocinaDetalle.frx` | Recurso binario del formulario Formularios\frmMensajeCocinaDetalle.frm | Formularios\frmMensajeCocinaDetalle.frm | USED |
| 140 | `Formularios\frmInsumo.frm` | FORM | `Formularios\frmInsumo.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 141 | `frmInsumo.frx` | FRX | `Formularios/frmInsumo.frx` | Recurso binario del formulario Formularios\frmInsumo.frm | Formularios\frmInsumo.frm | USED |
| 142 | `Formularios\frmInsumoDetalle.frm` | FORM | `Formularios\frmInsumoDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 143 | `frmInsumoDetalle.frx` | FRX | `Formularios/frmInsumoDetalle.frx` | Recurso binario del formulario Formularios\frmInsumoDetalle.frm | Formularios\frmInsumoDetalle.frm | USED |
| 144 | `Formularios\frmDiaContable.frm` | FORM | `Formularios\frmDiaContable.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 145 | `frmDiaContable.frx` | FRX | `Formularios/frmDiaContable.frx` | Recurso binario del formulario Formularios\frmDiaContable.frm | Formularios\frmDiaContable.frm | USED |
| 146 | `Formularios\frmConexionesActuales.frm` | FORM | `Formularios\frmConexionesActuales.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 147 | `frmConexionesActuales.frx` | FRX | `Formularios/frmConexionesActuales.frx` | Recurso binario del formulario Formularios\frmConexionesActuales.frm | Formularios\frmConexionesActuales.frm | USED |
| 148 | `Formularios\frmNuevaTienda.frm` | FORM | `Formularios\frmNuevaTienda.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 149 | `frmNuevaTienda.frx` | FRX | `Formularios/frmNuevaTienda.frx` | Recurso binario del formulario Formularios\frmNuevaTienda.frm | Formularios\frmNuevaTienda.frm | USED |
| 150 | `Formularios\frmNuevo.frm` | FORM | `Formularios\frmNuevo.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 151 | `frmNuevo.frx` | FRX | `Formularios/frmNuevo.frx` | Recurso binario del formulario Formularios\frmNuevo.frm | Formularios\frmNuevo.frm | USED |
| 152 | `Formularios\frmReciboEgreso.frm` | FORM | `Formularios\frmReciboEgreso.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 153 | `frmReciboEgreso.frx` | FRX | `Formularios/frmReciboEgreso.frx` | Recurso binario del formulario Formularios\frmReciboEgreso.frm | Formularios\frmReciboEgreso.frm | USED |
| 154 | `Formularios\frmReciboEgresoDetalle.frm` | FORM | `Formularios\frmReciboEgresoDetalle.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 155 | `frmReciboEgresoDetalle.frx` | FRX | `Formularios/frmReciboEgresoDetalle.frx` | Recurso binario del formulario Formularios\frmReciboEgresoDetalle.frm | Formularios\frmReciboEgresoDetalle.frm | USED |
| 156 | `Formularios\frmDocumentoCorrelativoDetalle.frm` | FORM | `Formularios\frmDocumentoCorrelativoDetalle.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 157 | `frmDocumentoCorrelativoDetalle.frx` | FRX | `Formularios/frmDocumentoCorrelativoDetalle.frx` | Recurso binario del formulario Formularios\frmDocumentoCorrelativoDetalle.frm | Formularios\frmDocumentoCorrelativoDetalle.frm | USED |
| 158 | `Formularios\frmBusquedaRapidaInvitado.frm` | FORM | `Formularios\frmBusquedaRapidaInvitado.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 159 | `frmBusquedaRapidaInvitado.frx` | FRX | `Formularios/frmBusquedaRapidaInvitado.frx` | Recurso binario del formulario Formularios\frmBusquedaRapidaInvitado.frm | Formularios\frmBusquedaRapidaInvitado.frm | USED |
| 160 | `Formularios\frmNuevoInvitado.frm` | FORM | `Formularios\frmNuevoInvitado.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 161 | `frmNuevoInvitado.frx` | FRX | `Formularios/frmNuevoInvitado.frx` | Recurso binario del formulario Formularios\frmNuevoInvitado.frm | Formularios\frmNuevoInvitado.frm | USED |
| 162 | `Formularios\frmConsultaSaldo.frm` | FORM | `Formularios\frmConsultaSaldo.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 163 | `frmConsultaSaldo.frx` | FRX | `Formularios/frmConsultaSaldo.frx` | Recurso binario del formulario Formularios\frmConsultaSaldo.frm | Formularios\frmConsultaSaldo.frm | USED |
| 164 | `Formularios\frmNumPad.frm` | FORM | `Formularios\frmNumPad.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 165 | `frmNumPad.frx` | FRX | `Formularios/frmNumPad.frx` | Recurso binario del formulario Formularios\frmNumPad.frm | Formularios\frmNumPad.frm | USED |
| 166 | `Formularios\frmBusquedaSocio.frm` | FORM | `Formularios\frmBusquedaSocio.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 167 | `frmBusquedaSocio.frx` | FRX | `Formularios/frmBusquedaSocio.frx` | Recurso binario del formulario Formularios\frmBusquedaSocio.frm | Formularios\frmBusquedaSocio.frm | USED |
| 168 | `Formularios\frmDescargo.frm` | FORM | `Formularios\frmDescargo.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 169 | `frmDescargo.frx` | FRX | `Formularios/frmDescargo.frx` | Recurso binario del formulario Formularios\frmDescargo.frm | Formularios\frmDescargo.frm | USED |
| 170 | `Formularios\frmVerificacionHuella.frm` | FORM | `Formularios\frmVerificacionHuella.frm` | Autenticación y seguridad | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 171 | `frmVerificacionHuella.frx` | FRX | `Formularios/frmVerificacionHuella.frx` | Recurso binario del formulario Formularios\frmVerificacionHuella.frm | Formularios\frmVerificacionHuella.frm | USED |
| 172 | `Formularios\frmVerificacionHuellaSup.frm` | FORM | `Formularios\frmVerificacionHuellaSup.frm` | Autenticación y seguridad | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 173 | `frmVerificacionHuellaSup.frx` | FRX | `Formularios/frmVerificacionHuellaSup.frx` | Recurso binario del formulario Formularios\frmVerificacionHuellaSup.frm | Formularios\frmVerificacionHuellaSup.frm | USED |
| 174 | `Formularios\frmImportacionRequerimientos.frm` | FORM | `Formularios\frmImportacionRequerimientos.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 175 | `frmImportacionRequerimientos.frx` | FRX | `Formularios/frmImportacionRequerimientos.frx` | Recurso binario del formulario Formularios\frmImportacionRequerimientos.frm | Formularios\frmImportacionRequerimientos.frm | USED |
| 176 | `Formularios\frmImportacionRequerimientoDetalle.frm` | FORM | `Formularios\frmImportacionRequerimientoDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 177 | `frmImportacionRequerimientoDetalle.frx` | FRX | `Formularios/frmImportacionRequerimientoDetalle.frx` | Recurso binario del formulario Formularios\frmImportacionRequerimientoDetalle.frm | Formularios\frmImportacionRequerimientoDetalle.frm | USED |
| 178 | `Formularios\frmCentralPedidos.frm` | FORM | `Formularios\frmCentralPedidos.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 179 | `frmCentralPedidos.frx` | FRX | `Formularios/frmCentralPedidos.frx` | Recurso binario del formulario Formularios\frmCentralPedidos.frm | Formularios\frmCentralPedidos.frm | USED |
| 180 | `Formularios\frmPedidosVerPagos.frm` | FORM | `Formularios\frmPedidosVerPagos.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 181 | `frmPedidosVerPagos.frx` | FRX | `Formularios/frmPedidosVerPagos.frx` | Recurso binario del formulario Formularios\frmPedidosVerPagos.frm | Formularios\frmPedidosVerPagos.frm | USED |
| 182 | `Formularios\frmBusquedaCentralPedido.frm` | FORM | `Formularios\frmBusquedaCentralPedido.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 183 | `frmBusquedaCentralPedido.frx` | FRX | `Formularios/frmBusquedaCentralPedido.frx` | Recurso binario del formulario Formularios\frmBusquedaCentralPedido.frm | Formularios\frmBusquedaCentralPedido.frm | USED |
| 184 | `Formularios\frmRepEntrega.frm` | FORM | `Formularios\frmRepEntrega.frm` | Salida de reportes | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 185 | `frmRepEntrega.frx` | FRX | `Formularios/frmRepEntrega.frx` | Recurso binario del formulario Formularios\frmRepEntrega.frm | Formularios\frmRepEntrega.frm | USED |
| 186 | `Formularios\frmDocumentoElectronicoCorrelativo.frm` | FORM | `Formularios\frmDocumentoElectronicoCorrelativo.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 187 | `frmDocumentoElectronicoCorrelativo.frx` | FRX | `Formularios/frmDocumentoElectronicoCorrelativo.frx` | Recurso binario del formulario Formularios\frmDocumentoElectronicoCorrelativo.frm | Formularios\frmDocumentoElectronicoCorrelativo.frm | USED |
| 188 | `Formularios\frmPagoMasivo.frm` | FORM | `Formularios\frmPagoMasivo.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 189 | `frmPagoMasivo.frx` | FRX | `Formularios/frmPagoMasivo.frx` | Recurso binario del formulario Formularios\frmPagoMasivo.frm | Formularios\frmPagoMasivo.frm | USED |
| 190 | `Formularios\frmClienteDeliveryFrecuente.frm` | FORM | `Formularios\frmClienteDeliveryFrecuente.frm` | Flujos delivery/dispatch dependientes del POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 191 | `frmClienteDeliveryFrecuente.frx` | FRX | `Formularios/frmClienteDeliveryFrecuente.frx` | Recurso binario del formulario Formularios\frmClienteDeliveryFrecuente.frm | Formularios\frmClienteDeliveryFrecuente.frm | USED |
| 192 | `Formularios\FrmMensajeVisor.frm` | FORM | `Formularios\FrmMensajeVisor.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 193 | `Formularios\frmMensajeFeSpring.frm` | FORM | `Formularios\frmMensajeFeSpring.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 194 | `frmMensajeFeSpring.frx` | FRX | `Formularios/frmMensajeFeSpring.frx` | Recurso binario del formulario Formularios\frmMensajeFeSpring.frm | Formularios\frmMensajeFeSpring.frm | USED |
| 195 | `Formularios\frmFechaEmision.frm` | FORM | `Formularios\frmFechaEmision.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 196 | `frmFechaEmision.frx` | FRX | `Formularios/frmFechaEmision.frx` | Recurso binario del formulario Formularios\frmFechaEmision.frm | Formularios\frmFechaEmision.frm | USED |
| 197 | `Formularios\FrmRecargarTarjetaDetalle.frm` | FORM | `Formularios\FrmRecargarTarjetaDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 198 | `FrmRecargarTarjetaDetalle.frx` | FRX | `Formularios/FrmRecargarTarjetaDetalle.frx` | Recurso binario del formulario Formularios\FrmRecargarTarjetaDetalle.frm | Formularios\FrmRecargarTarjetaDetalle.frm | USED |
| 199 | `Formularios\FrmRecargarTarjeta.frm` | FORM | `Formularios\FrmRecargarTarjeta.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 200 | `FrmRecargarTarjeta.frx` | FRX | `Formularios/FrmRecargarTarjeta.frx` | Recurso binario del formulario Formularios\FrmRecargarTarjeta.frm | Formularios\FrmRecargarTarjeta.frm | USED |
| 201 | `Formularios\FrmMovimientoTarjetas.frm` | FORM | `Formularios\FrmMovimientoTarjetas.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 202 | `FrmMovimientoTarjetas.frx` | FRX | `Formularios/FrmMovimientoTarjetas.frx` | Recurso binario del formulario Formularios\FrmMovimientoTarjetas.frm | Formularios\FrmMovimientoTarjetas.frm | USED |
| 203 | `Formularios\FrmTarjetaAproximidad.frm` | FORM | `Formularios\FrmTarjetaAproximidad.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 204 | `FrmTarjetaAproximidad.frx` | FRX | `Formularios/FrmTarjetaAproximidad.frx` | Recurso binario del formulario Formularios\FrmTarjetaAproximidad.frm | Formularios\FrmTarjetaAproximidad.frm | USED |
| 205 | `Formularios\FrmTarjetaAproximidadDetalle.frm` | FORM | `Formularios\FrmTarjetaAproximidadDetalle.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 206 | `FrmTarjetaAproximidadDetalle.frx` | FRX | `Formularios/FrmTarjetaAproximidadDetalle.frx` | Recurso binario del formulario Formularios\FrmTarjetaAproximidadDetalle.frm | Formularios\FrmTarjetaAproximidadDetalle.frm | USED |
| 207 | `Formularios\frmNoticias.frm` | FORM | `Formularios\frmNoticias.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 208 | `Formularios\FrmActualizacion.frm` | FORM | `Formularios\FrmActualizacion.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 209 | `Formularios\frmUpdateDatosPedido.frm` | FORM | `Formularios\frmUpdateDatosPedido.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 210 | `Formularios\frmListaMotivos.frm` | FORM | `Formularios\frmListaMotivos.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 211 | `frmListaMotivos.frx` | FRX | `Formularios/frmListaMotivos.frx` | Recurso binario del formulario Formularios\frmListaMotivos.frm | Formularios\frmListaMotivos.frm | USED |
| 212 | `Formularios\frmSolicitudBoleta.frm` | FORM | `Formularios\frmSolicitudBoleta.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 213 | `Formularios\frmVerificarHuellaMarcacion.frm` | FORM | `Formularios\frmVerificarHuellaMarcacion.frm` | Autenticación y seguridad | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 214 | `frmVerificarHuellaMarcacion.frx` | FRX | `Formularios/frmVerificarHuellaMarcacion.frx` | Recurso binario del formulario Formularios\frmVerificarHuellaMarcacion.frm | Formularios\frmVerificarHuellaMarcacion.frm | USED |
| 215 | `Formularios\frmMarca.frm` | FORM | `Formularios\frmMarca.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 216 | `Formularios\frmCambiarContrasenia.frm` | FORM | `Formularios\frmCambiarContrasenia.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 217 | `frmCambiarContrasenia.frx` | FRX | `Formularios/frmCambiarContrasenia.frx` | Recurso binario del formulario Formularios\frmCambiarContrasenia.frm | Formularios\frmCambiarContrasenia.frm | USED |
| 218 | `Formularios\frmDocumentoCompra.frm` | FORM | `Formularios\frmDocumentoCompra.frm` | Emisión/documentación comercial | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 219 | `frmDocumentoCompra.frx` | FRX | `Formularios/frmDocumentoCompra.frx` | Recurso binario del formulario Formularios\frmDocumentoCompra.frm | Formularios\frmDocumentoCompra.frm | USED |
| 220 | `Formularios\FrmSeleccionProductos.frm` | FORM | `Formularios\FrmSeleccionProductos.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 221 | `FrmSeleccionProductos.frx` | FRX | `Formularios/FrmSeleccionProductos.frx` | Recurso binario del formulario Formularios\FrmSeleccionProductos.frm | Formularios\FrmSeleccionProductos.frm | USED |
| 222 | `Formularios\frmBusquedaValeConsumo.frm` | FORM | `Formularios\frmBusquedaValeConsumo.frm` | Componente declarado en InfoRest.vbp | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 223 | `frmBusquedaValeConsumo.frx` | FRX | `Formularios/frmBusquedaValeConsumo.frx` | Recurso binario del formulario Formularios\frmBusquedaValeConsumo.frm | Formularios\frmBusquedaValeConsumo.frm | USED |
| 224 | `Formularios\frmMesas.frm` | FORM | `Formularios\frmMesas.frm` | Gestión de mesas/salón | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 225 | `frmMesas.frx` | FRX | `Formularios/frmMesas.frx` | Recurso binario del formulario Formularios\frmMesas.frm | Formularios\frmMesas.frm | USED |
| 226 | `Formularios\frmDetallePedido.frm` | FORM | `Formularios\frmDetallePedido.frm` | Gestión de pedidos POS | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 227 | `frmDetallePedido.frx` | FRX | `Formularios/frmDetallePedido.frx` | Recurso binario del formulario Formularios\frmDetallePedido.frm | Formularios\frmDetallePedido.frm | USED |
| 228 | `Formularios\frmJuntarMesas.frm` | FORM | `Formularios\frmJuntarMesas.frm` | Gestión de mesas/salón | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 229 | `frmJuntarMesas.frx` | FRX | `Formularios/frmJuntarMesas.frx` | Recurso binario del formulario Formularios\frmJuntarMesas.frm | Formularios\frmJuntarMesas.frm | USED |
| 230 | `Formularios\frmPagoPinPad.frm` | FORM | `Formularios\frmPagoPinPad.frm` | Cobro y medios de pago | modPuntoVenta.bas/mdiPuntoVenta.frm | USED |
| 231 | `modPuntoVenta` | MODULE | `Modulos\modPuntoVenta.bas` | Operación de venta POS | InfoRest.vbp Sub Main | USED |
| 232 | `modDeclaracion` | MODULE | `Modulos\modDeclaracion.bas` | Estado global y flags | InfoRest.vbp Sub Main | USED |
| 233 | `modProcedimiento` | MODULE | `Modulos\modProcedimiento.bas` | Utilitarios operativos centrales | InfoRest.vbp Sub Main | USED |
| 234 | `DLL3500` | MODULE | `Modulos\DLL3500.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 235 | `ModuloHardKey` | MODULE | `Modulos\ModuloHardKey.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 236 | `modConexionIp` | MODULE | `Modulos\modConexionIp.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 237 | `modKDS` | MODULE | `Modulos\modKDS.bas` | Producción/cocina | InfoRest.vbp Sub Main | USED |
| 238 | `modCrearInis` | MODULE | `Modulos\modCrearInis.bas` | Configuración del sistema | InfoRest.vbp Sub Main | USED |
| 239 | `CodigoControl` | MODULE | `Modulos\CodigoControl.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 240 | `modSeguridadInfhotel` | MODULE | `Modulos\modSeguridadInfhotel.bas` | Autenticación y seguridad | InfoRest.vbp Sub Main | USED |
| 241 | `modAuditoriaIntegral` | MODULE | `Modulos\modAuditoriaIntegral.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 242 | `FpLibX_Const` | MODULE | `Modulos\FpLibX_Const.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 243 | `ModuleBarcode` | MODULE | `Modulos\modBarcode.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 244 | `modAuditoriaEquipo` | MODULE | `Modulos\modAuditoriaEquipo.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 245 | `modImpresoraFiscal` | MODULE | `Modulos\modImpresoraFiscal.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 246 | `modTime` | MODULE | `Modulos\modTime.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 247 | `modProcedimientoNuevo` | MODULE | `Modulos\modProcedimientoNuevo.bas` | Utilitarios operativos centrales | InfoRest.vbp Sub Main | USED |
| 248 | `ModPictureBoxCustom` | MODULE | `Modulos\ModPictureBoxCustom.bas` | Componente declarado en InfoRest.vbp | InfoRest.vbp Sub Main | USED |
| 249 | `clsComando` | CLASS | `Clases\clsComando.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 250 | `License` | CLASS | `Clases\License.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 251 | `Mapping` | CLASS | `Clases\Mapping.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 252 | `clsxml` | CLASS | `Clases\clsxml.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 253 | `clsDiaContable` | CLASS | `Clases\clsDiaContable.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 254 | `clsAlmacen` | CLASS | `Clases\clsAlmacen.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 255 | `claCorreoElectronico` | CLASS | `Clases\claCorreoElectronico.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 256 | `clsTrama` | CLASS | `Clases\clsTrama.cls` | Componente declarado en InfoRest.vbp | InfoRest.vbp / módulos y formularios | USED |
| 257 | `ClsSeguridad` | CLASS | `Clases\ClsSeguridad.cls` | Autenticación y seguridad | InfoRest.vbp / módulos y formularios | USED |
| 258 | `ClsDocumento` | CLASS | `Clases\ClsDocumento.cls` | Emisión/documentación comercial | InfoRest.vbp / módulos y formularios | USED |
| 259 | `Reportes\dsrCuentaCobrar.Dsr` | REPORT_DESIGNER | `Reportes\dsrCuentaCobrar.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 260 | `Reportes\dsrGuiaCorrelativo.Dsr` | REPORT_DESIGNER | `Reportes\dsrGuiaCorrelativo.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 261 | `Reportes\dsrRecibo.Dsr` | REPORT_DESIGNER | `Reportes\dsrRecibo.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 262 | `Reportes\dsrCtaCteD.Dsr` | REPORT_DESIGNER | `Reportes\dsrCtaCteD.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 263 | `Reportes\dsrCtaCteR.Dsr` | REPORT_DESIGNER | `Reportes\dsrCtaCteR.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 264 | `Reportes\dsrCtaCteC.Dsr` | REPORT_DESIGNER | `Reportes\dsrCtaCteC.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 265 | `Reportes\dsrReserva.Dsr` | REPORT_DESIGNER | `Reportes\dsrReserva.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 266 | `Reportes\dsrReporteMensajeCocina.Dsr` | REPORT_DESIGNER | `Reportes\dsrReporteMensajeCocina.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 267 | `Reportes\dsrRepInsumos.Dsr` | REPORT_DESIGNER | `Reportes\dsrRepInsumos.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 268 | `Reportes\dsrRepInsumosPtoVenta.Dsr` | REPORT_DESIGNER | `Reportes\dsrRepInsumosPtoVenta.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 269 | `Reportes\dsrReciboEgreso.Dsr` | REPORT_DESIGNER | `Reportes\dsrReciboEgreso.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 270 | `Reportes\dsrReporteEntrega.Dsr` | REPORT_DESIGNER | `Reportes\dsrReporteEntrega.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 271 | `Reportes\dsrReporteEntregaResumidoProd.Dsr` | REPORT_DESIGNER | `Reportes\dsrReporteEntregaResumidoProd.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 272 | `Reportes\dsrBoleta.Dsr` | REPORT_DESIGNER | `Reportes\dsrBoleta.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 273 | `Reportes\dsrBoletaC.Dsr` | REPORT_DESIGNER | `Reportes\dsrBoletaC.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 274 | `Reportes\dsrNotaCredito.Dsr` | REPORT_DESIGNER | `Reportes\dsrNotaCredito.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 275 | `Reportes\dsrReporteEntregaFormato2.Dsr` | REPORT_DESIGNER | `Reportes\dsrReporteEntregaFormato2.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 276 | `Reportes\dsrNotaCreditoObservacion.Dsr` | REPORT_DESIGNER | `Reportes\dsrNotaCreditoObservacion.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 277 | `Reportes\dsrRepDescargo.Dsr` | REPORT_DESIGNER | `Reportes\dsrRepDescargo.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 278 | `Reportes\dsrNewFormatoGuia.Dsr` | REPORT_DESIGNER | `Reportes\dsrNewFormatoGuia.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 279 | `Reportes\dsrReciboEgresoCompra.Dsr` | REPORT_DESIGNER | `Reportes\dsrReciboEgresoCompra.Dsr` | Diseño de reporte VB6 (DataReport/Crystal) | Formularios de reportes POS | USED |
| 280 | `Idiomas.RES` | RES | `Idiomas.RES` | Recursos de idioma compilados | InfoRest.vbp ResFile32 | USED |
| 281 | `Inforest.ini` | CONFIG | `Inforest.ini` | Configuración principal POS | modPuntoVenta.bas | USED |
| 282 | `Facturacion.ini` | CONFIG | `Facturacion.ini` | Configuración FE | modPuntoVenta.bas | USED_CONDITIONAL |
| 283 | `Almacen.ini` | CONFIG | `Almacen.ini` | Configuración almacén | modPuntoVenta.bas | USED_CONDITIONAL |
| 284 | `Infhotel.ini` | CONFIG | `Infhotel.ini` | Integración hotelera | modPuntoVenta.bas | USED_CONDITIONAL |
| 285 | `RUTA.INI` | CONFIG | `RUTA.INI` | Rutas auxiliares | Runtime legacy | UNKNOWN |
| 286 | `USUARIO.INI` | CONFIG | `USUARIO.INI` | Configuración auxiliar de usuario | Runtime legacy | UNKNOWN |
| 287 | `DLL3500.ini` | CONFIG | `DLL3500.ini` | Parámetros PinPad | DLL3500.bas | USED_CONDITIONAL |
| 288 | `Tiempo.ini` | CONFIG | `Tiempo.ini` | Parámetros de sincronización/tiempo | modTime.bas | USED_CONDITIONAL |
| 289 | `MSCOMCTL.OCX` | OCX | `MSCOMCTL.OCX` | Controles UI VB6 | InfoRest.vbp Object | USED |
| 290 | `MCI32.OCX` | OCX | `MCI32.OCX` | Multimedia/control auxiliar | InfoRest.vbp Object | USED_CONDITIONAL |
| 291 | `IFEpson.ocx` | OCX | `IFEpson.ocx` | Impresión fiscal Epson (Argentina) | modImpresoraFiscal.bas | EXCLUDED_BY_SCOPE |
| 292 | `CAJA_PINPAD.dll` | DLL | `CAJA_PINPAD.dll` | Integración PinPad | DLL3500.bas | USED_CONDITIONAL |
| 293 | `MSBIND.DLL` | DLL | `MSBIND.DLL` | Data binding VB6 | InfoRest.vbp Reference | USED |
| 294 | `dzstactx.dll` | DLL | `dzstactx.dll` | Componente ActiveX legacy | InfoRest.vbp Object | USED_CONDITIONAL |
| 295 | `hkey-w32.dll` | DLL | `hkey-w32.dll` | Soporte de hardkey/licencia | ModuloHardKey.bas | USED_CONDITIONAL |
| 296 | `qrcodelib.dll` | DLL | `qrcodelib.dll` | Generación QR | Facturación electrónica | USED_CONDITIONAL |