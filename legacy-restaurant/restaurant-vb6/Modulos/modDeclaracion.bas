Attribute VB_Name = "modDeclaracion"
'Desclaraciones Temporales
Public Const sMensaje = "Mensaje Inforest"

'Declaración de Biblioteca del Windows
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function OpenFile% Lib "kernel32" (ByVal lpFileName$, lpReOpenBuff As OFSTRUCT, ByVal wStyle%)
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Declaración de Variables de Conneccion
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inforest\Data\inforest.mdb;Persist Security Info=False


Type OFSTRUCT '136 bytes in length
    cBytes As String * 1
    fFixedDisk As String * 1
    nErrCode As Integer
    Reserved As String * 4
    szPathName As String * 128
End Type

Type SHFILEOPSTRUCT
    hwnd                  As Long
    wFunc                 As Long
    pFrom                 As String
    pTo                   As String
    fFlags                As Long
    fAnyOperationsAborted As Boolean
    hNameMappings         As Long
    lpszProgressTitle     As String
End Type

Public Enum eFO
       FO_COPY = &H2&
       FOF_NOCONFIRMATION = &H10&
       FOF_NOCONFIRMMKDIR = &H200&
End Enum

Public tenlacebusqueda, tenlacebusquedaVenta As String '-- BUSQUEDA DE DATOS EN FORMULARIO VENTA , CAJA RAPIDA Y ADICION
Public Lib As Libreria16.Applications
Public Cn As Connection
Public Conex As Connection
Public CnAlmacen As Connection
Public CnDBF As Connection
Public CnInfhotel As Connection
Public sUserName As String
Public sUserPassword As String
Public sLocal As String
'consulta integrada
Public rsListaServidores As ADODB.Recordset

Public BDLink As String

'Declaracion de Variables Publicas
'Parametros Generales
Public sCliente As String
Public sRazonSocial As String
Public sRazonComercial As String
Public sDireccion As String
Public sDireccion2 As String
Public sMail As String
Public sWeb As String
Public sTelefono As String
Public sFax As String
Public sRUC As String
Public sMonN As String
Public sMonedaN As String
Public sMonE As String
Public sMonE2 As String
Public sMonE3 As String

Public sMonedaE As String
Public sMonedaE2 As String
Public sMonedaE3 As String
Public nDELIVERY As Double
Public nLlevar As Double
Public nCanal4 As Double
Public nCanal5 As Double
Public sImpuesto1 As String
Public sImpuesto2 As String
Public sImpuesto3 As String
Public nPorcentaje1 As Double
Public nPorcentaje2 As Double
Public nPorcentaje3 As Double
Public sPie As String
Public sInfacto As String
Public sExonerado As String

Public sPiePreCuenta As String
Public sPAdmin As String
Public sElimina As String
Public nTiempo As Integer
Public nFItem As Integer
Public nCabecera As Integer
Public nDetalle As Integer
Public nFItemGuia As Integer
Public nCabeceraGuia As Integer
Public nDetalleGuia As Integer
Public nChkTiempo As Integer
Public sComanda As String
Public sResumen As Boolean
Public sEmpresa As String
Public sUltimo As String
Public sBoton1 As String
Public sBoton2 As String
Public sBoton3 As String
Public sBoton4 As String
Public sBoton5 As String
Public nLongitud As Integer
Public lLongitud As Boolean
Public lPrinter As Boolean
Public lRapido As Boolean
Public lInfhotel As Boolean
Public sClub As String
Public nPunto As Double
Public nDecimal As Integer
Public nDias As Integer
Public nDiasDelivery As Integer
Public nPuerto As Integer
'Public nPuerto As Integer

Public nBalanzaComando As String
Public nBalanzaPuerto As Integer
Public nBalanzaBS As Long
Public nBalanzaBD As Long
Public nBalanzaParidad As String
Public nBalanzaBP As Long
Public nBalanzaCF As String
Public lBalTiempo As Integer
Public lBal As Boolean

Public tMensaje1 As String
Public tMensaje2 As String
Public lEquivalencia As Boolean
Public lSomelier As Boolean
Public lComboGeneral As Boolean
Public nTiempoDelivery As Integer
Public lObligaPax As Boolean
Public fImpresionDiaContable As Date
Public lImpresionDiaContable As Boolean
Public tContribuyenteEspecial As String
Public fContribuyenteEspecial As String
 
'Variables de Impresora
Public sSerieCaja As String
Public sFont As String
Public sFont1 As String
Public sFont2 As String
Public sPuertoIF As String
Public sBitSegIF As String

'Variables Caja
Public sPreCuenta As String
Public sDespachador As String
Public lOrdenesRappi As Boolean
Public wComanda As Boolean
Public vComanda As Boolean
Public lEliminaC As Boolean
Public lElimina As Boolean
Public lPasswordC As Boolean
Public lPassword As Boolean
Public sGrupoDefault As String
Public lConsumo1 As Boolean
Public lConsumo2 As Boolean
Public lConsumo3 As Boolean
Public lPrecuentaImpresora As Boolean
Public lAdicion As Boolean
Public lCierre As Boolean
Public lPrecuentaAgrupada As Boolean
Public lDocumentoAgrupado As Boolean
Public sTipoPedido As String
Public lCuentasXCobrar As Boolean
Public sTipoPedidoPD As String
Public lObligaCierre As Boolean
Public lFiltroTipoPedido As Boolean
Public lCancelacion As Boolean
Public lObligaPrinter As Boolean
Public lDirecto As Boolean
Public lObligaPrecuenta As Boolean
Public lComboPrecuenta As Boolean
Public lComboDocumento As Boolean
Public lCambioMesa As Boolean
Public lVisaNet As Boolean
Public lImpuestoPrecuenta As Boolean
Public lOrden As Boolean
Public lValorCortesia As Boolean
Public lObservacion As Boolean
Public lCajaRapida As Boolean
Public lPropiedadDocumento As Boolean
Public lPropiedadPrecuenta As Boolean
Public lPrecioNetoPrecuenta As Boolean
Public nLimitePrecuenta As Integer
Public nLimiteReimpresion As Integer
Public lPasswordTransferencia As Boolean
Public nLongitudBarra As Integer
Public lCapturaPeso As Boolean
Public lPasswordImportarPedido As Boolean
Public sUnidadNegocio As String
Public lMultiCajero As Boolean
Public lMCPV As Boolean
Public lFechaEntregaDelivery As Boolean
Public lCCVOX As Boolean
Public lObservacionPrecuenta As Boolean
Public lObservacionDocumento As Boolean
Public lObservacionCabDoc As Boolean
Public lPagoRapidoPV As Boolean
Public tTextoConsumo As String
Public lImpreEquivaPrecuenta As Boolean
Public lDocumEquivaPrecuenta As Boolean
Public lDisgrega As Boolean
Public lSiab As Boolean
Public CnSiab As Connection
Public sSYBASE As String
Public lBusquedaReservas As Boolean
Public lActivarValesConsumo As Boolean

' VALES DE CONSUMO
Public RsGrillaVConsumo As Recordset

Public nCorrelaVC As String
Public sTarjetaVC As String
Public sTarjetaVC_Final As String

Public lGuardarVC As Boolean




''''''''''''''''''''''
'Parametros de Ingreso
Public sRuta As String
Public sMDB As String
Public sAlmacenMDB As String
Public sAlmacenRuta As String
Public sDBF As String
Public lAlmacen As Boolean
Public lConsultaCierreInv As Boolean
Public sModulo As String
Public sRutaCD As String
Public sMDBCD As String

Public sRutaWAP As String
Public sMDBWAP As String

'--- SAP
Public lSAP As Boolean
Public sServidorSAp As String
Public sBdSAP As String
Public sCodSap As String
Public VSApForBusqueda As Boolean
'----visor----
Public lvisor  As Boolean
    
'Variables Infhotel
Public sHotel As String
Public sReserva As String
Public sHabitacion As String
Public sPasajero As String
Public sFichaPasajero As String
Public sCajaInfhotel As String
Public xUsuario As String
Public sPropina   As String
Public sMonPropina  As String
Public sTipoComanda As String
Public sPuntoVentaInfhotel As String

'Variables Genericas
Public sTurno As String
'CESAR TURNO
Public sTurnoModificacion As String


Public wInicioActivaApertura As Boolean

Public sCaja As String
Public sSalon As String
Public sUsuario As String
Public sPassword As String
Public sPedido As String
Public IsPedidoCajaRapida As Boolean
Public sDocumento As String
Public sMesa As String
Public wCambioMesa As Boolean
Public nTC As Double
Public nTC2 As Double
Public nTC3 As Double
Public sGrupoUsuario As String
Public sMozo As String
Public lImportacionReservas As Boolean
'origen de ventas
Public sOrigenVenta As String

Public lBotonTrans As Boolean
Public nCargo As Double
Public NFactura As String
Public sTienda As String
Public sTemporal As String
Public xCantidad As Integer

'Variables CD
Public sUserNameCD As String
Public sUserPasswordCD As String
Public CD As Boolean


Public WEBAP As Boolean
Public MESA247 As Boolean
Public EAN13 As Boolean

'Variables de Programacion
Public Isql As String
Public sCodigo As String
Public sMontoBermur As Double

Public sDescrip As String
Public nPos As Variant
Public nFactor As Double
Public lFactor As Boolean

Public sTemp As String
Public sTipo As String
Public xTipo As String
Public wEnter As Boolean
Public wMesa As Boolean
Public wInicio As Boolean
Public Sw As Boolean
Public sVar1 As String
Public sVar2 As String
Public sVar3 As String
Public nVar1 As Double
Public sFormulario As String
Public CorrelativoC As Integer

'GCAA 15092022 - VARIABLE PARA DIVISION PAGO DOCUMENTOS
Public EstoyDivisionDoc As String


'Declaracion de Nota de Credito
Public lactivaFechaNC As Boolean
Public lParcialNC As Boolean
Public lNCElimina As Boolean
Public lNCAnula As Boolean
'Declaracion de Variables de Color
Public Enum UbicaPuntero
   Primero = 1
   Ultimo = 2
   previo = 3
   siguiente = 4
   pgup = 5
   pgdn = 6
End Enum

Public Enum TAmbiente
    Prueba = 1
    Produccion = 2
End Enum

Public Enum TTipoEmision
    Normal = 1
    IndisponibilidadSys = 2
End Enum

Public xTotal As Double
Public xtipodoc As String




Public Const vbLista = &HFF8080
Public Const vbOcupada = &H80C0FF
Public Const vbReservada = &H8000&
Public Const vbSucia = &HC0C0&
Public Const vbBloqueada = &HC0&
Public Const vbFServicio = &HC0C0C0
Public Const vbPrecuentaImp = &H80FF&
Public Const vbOriginal = &H4080&
'Color mesa
Public Const vbHexLista = "#FFFFFF"
Public Const vbHexOcupada = "#FFC080"
Public Const vbHexPrecuenta = "#C0B9FA"
Public Const vbHexSucia = "#ECEC93"
Public Const vbHexMesasjuntas = "#76E4E6"
Public Const vbHexReservada = "#A1CCA1"
Public Const vbHexBloqueada = "#FF8484"
Public Const vbHexFServicio = "#E2E5E8"
'Public Const vbHexPrecuentaImp = &H80FF&
'version educativa
Public lVersionEducativa As Boolean
'HARDkey
Public lHARDkey As Boolean
Public clave1 As String
Public clave2 As String
'------------------------------

'ADMINISTRACION CENTRALIZADA
Public sServidorCentral As String
Public bdInforestCentral As String
Public lCentral As Boolean
Public lMenuTablas As Boolean
'======================
'CLIENTES CENTRALIZADOS
Public xlCentral As Boolean
Public xservidorcentral As String
Public xbasecenntral As String
Public xcnDefault As New ADODB.Connection
Public VerConexionInfoRestCentral As String
Public xrsParientes As Recordset
Public RsxCanalventa As Recordset
'====================== un exe multilocal

Public localConectado As String
Public multiLocal As Boolean
Public ultimoConectado As Boolean
Public cnDefault As New ADODB.Connection


'====================
Public cnAlmacenDefault As New ADODB.Connection
'====================

Public moduloUso As String


'ALMACEN REMOTO
Public sRutaAlmacenRemoto As String
Public sMDBAlmacenRemoto As String
Public CnAlmacenRemoto As ADODB.Connection
Public lAlmacenRemoto As Boolean
Public verificaAlmacenRemoto As String

'KDS
Public lKDS As Boolean
Public sOrderInfo As String
Public sBump As String

Public lKDSx As Boolean
Public sOrderInfox As String
Public sBumpx As String
Public lNumdiasAnulacion As Integer


Public lKDSInforest As Boolean
Public lKDSxArea As Boolean
Public lKDSTodos As Boolean

Public lActivaMotivoAnulacion As Boolean

Public lFEBizOffLine As Boolean
Public lActivarIncrementoPrecios As Boolean


'''
Public PrograPre_Titulo As String
Public PrograPre_Fecha As String
Public PrograPre_Hora As String
Public PrograPre_Saber As String
'''



'descripcion alternativa

Public lDescripcionAlternativa As Boolean

'para mesas
Public sTempMesa As String

'TVS
'Public lCompatibilidadTVS As Boolean


'EXTRANJERO- BOLIVIA
Public pais As String
Public textoComprobanteBolivia As String
'cgMiranda-------------------------
Public estadoReimpresion As Boolean
'fin cgmiranda---------------------
' controler
Public tcodigoUsuarioA As String


'pagoRapido
Public lPagoRapido As Boolean


'MODIFICACIONES NORKYS lg 03/2012
Public lPasswordPorCobrar As Boolean
Public lmodificatipoPedido As Boolean

'registroventasunat
Public nTCO As Double

'diacontable
Public lDiaContable  As Boolean 'true=automatico false=manual
Public tHoraCierreDiaContable As String
Public lIniciaPorDiaContable As Boolean
Public lDiaContableAperturado As Boolean

'club
Public lClub As Boolean

'control de licencias
Public tVersionEducativaLicencia As String


'tarifa
Public tTarifaActualMotorizado As String

'AUDITORIA
Public tModuloSeg As String
Public nCorrelativoAcceso As Long
Public CnSeg As Connection
Public sPasa As String
Public lAuditoria As Boolean

'impresion
Public lReimpresion As Boolean

'0084-2013 CESAR Para el Pago Rapido
Public lModuloPago As String


'Invitado
Public sCodigoInvitado As String


'PARIENTE
Public sCodigoParienteSeleccionado As String
Public sCodigoClienteFrecuente As String


'busquedaSocio
Public lCargaDesdePedido As Boolean
Public lTabBuscar As Boolean

'CESAR PAGO RAPIDO
Public nTotalPR As Double


'lg tipocanalreporte
Public sTipoCanalReporte As String
Public sTipoCanalNombreReporte As String

'consulta descargo al cierre
Public lActivaConsultaDescargo As Boolean

'bloqueaprecuenta
Public lBloqueaPrecuenta As Boolean

'CESAR ROTULADO
Public lRotulado As Boolean


'lg
Public lMultiAreaSubGrupo As Boolean
Public lMultiAreaCaja As Boolean


'HUELLA
Public wEnterHuella As Boolean
Public wenterHuellaSup As Boolean
Public lVieneHuella As Boolean
Public pTipo As String
Public lHuella As Boolean
Public lUsuarioHuella As Boolean


'FORMATO VARIABLE
Public nCabeceraV As Integer
Public nItemV As Integer
Public nPieV As Integer

'envio de correo
Public sUsuarioMail As String
Public sClaveMail As String
Public Smtp_Prorroga As String
Public UsuarioEnvio_Prorroga As String
Public PasswordEnvio_Prorroga As String
Public Asunto_Prorroga As String
Public Cuerpo_Prorroga As String

'Validacion RUC
Public xlTipoDocumento As Boolean

'Requerimientos
Public sRequerimiento As String

'ImpresionImagenPrecuenta
Public lImprimeImagCabPrecuenta As Boolean
Public lImprimeImagPiePrecuenta As Boolean

'MostrarVencimiento
Public sVencimientoLicencia As String
Public sCantidadLicenciaModulo As String

'ImpresionImagenDocumento
Public lImprimeImagCabDocumento As Boolean
Public lImprimeImagPieDocumento As Boolean
Dim lTamanio As Integer
Dim rsQr As New ADODB.Recordset
Dim comilla As String

'despachopedido
Public lAccesoDespachoPedido As Boolean
Public nLongitudAlmacen As Long



'FACTURACION ELECTRONICA
Public lFacturacionE As Boolean
Public tCodigoFE As String
Public lAmbienteProduccion As Boolean
Public tPieDocumento1 As String

Public sRutaFE As String
Public sMDBFE As String
Public CnFE As Connection
Public IsqlFact As String
Public xi As Integer

' erick DLC
Public lQRFE As Boolean
Public RutaImgFE As String



Public ClaveAcceso As String


'huellas
Public lHuellaDigitalPersona As Boolean
Public lHuellaSecugen As Boolean
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long




'AGENTE RETENCION
Public tTextoAgenteRetencion As String

'imprime dni
Public xTipoIdentidad As String

'guarda LOG CAJA RAPIDA
Public lLogCajaRapida As Boolean
'por pastipan
Public lBuscaPedidoNumero As Boolean

'codigobarrarecibo
Public lCodigoReciboIngreso As Boolean

'autorizacion
Public tUsuarioAutoriza As String

Public lConsumo4 As Boolean
Public lPrecuentaNoValorizada As Boolean

'PAGO MASIVO
Public tTipoPagoMasivo As String
Public tTipoTarjetaMasivo As String
Public tOtroTipoCancelacionMasivo  As String
Public tDocumentoPagoMasivo As String
Public tBancoPagoMasivo As String

Public lImpresionCodigoBarras As Boolean


Public lEnvioAutomatico As Boolean
Public lMinutoEnvioAntes As Integer
Public lEnvioProduccionUsuario As Boolean
Public lEnvioProduccionCaja As Boolean

'ENLACE ELECTRONICO OFISIS
Public lFEOfisis As Boolean
Public lFEOventas As Boolean

Public sTelefonoReserva As String


'BUSCAR
Public lBuscarPedidoVisualizarGrilla As Boolean
Public lBuscarPedidoFiltrarMesa As Boolean

' glosa de impresion transferencia gratuita
Public lGlosaTrans As String

Public lPagoAntesImpresion As Boolean

Public lPagocortesiaAI As Boolean
' Cover - ecuador
Public lcover As Boolean
Public sMontoMinCover As Double
Public sCodItemCover As String


'-- ofisis notas de credito
Public lNcOfisis As Boolean

Public AvisoLicencia As Integer

Public lFEpape As Boolean
Public IPpape As String
Public PUERTOpape As String
Public PUERTOLOCALpape As String

Public lDesPagoCheque As Boolean
Public lDesPagoOtra As Boolean

Public lKardexFechaIngreso As Boolean

' FACTURACION PAPERLEES
Public clsTramaFE As clsTrama
Public PapeMatricial As String
Public PapeTermico As String

'ENLACE ELECTRONICO SPRING
Public lFESpring As Boolean

'ENLACE ELECTRONICO CARBAJAL
Public lFECarbajal As Boolean

Public lDesactivaNCFP As Boolean

Public Estado As String

Public lFEBiz As Boolean

Public lImprimeMotivoDescuentoFB As Boolean

Public lActivaAnticipo As Boolean

Public lFeGoodHope As Boolean

Public lImpPropina As Boolean

Public lImpComandaf2 As Boolean

Public lFEubl21 As Boolean

Public lPassOtrosPagos As Boolean

Public lBloqInafecto As Boolean

Public lFEEcuador As Boolean

Public lFEGesa As Boolean

Public DiasContingenciaDocumentos As Double

Public lImpClienteCabDoc As Boolean

Public lBuscarPedidoBarraMesero As Boolean

Public sLey1 As String
Public sValorLey1 As String

Public RutaConsultaRuc As String

Public RutaConsultaRucRetenciones As String

Public VisibleTC As Integer

Public lIntVisaNetQR As Boolean

Public IntFPAY As Boolean

Public lActivaCreditoCoorporativo As Boolean

Public lIntMercadoPagoQR As Boolean

Public lRutaImagenIntegraciones As String

Public lImpComandaBarra As Boolean

Public lbloqueadocconsumo As Boolean

Public lCashDro As Boolean

Public MensajeVisaQR As String

Public TipoFacturacion As String

Public lBloqueoAfectoInafecto As Boolean

Public lIntPatio As Boolean

Public lIntMercadoPagoQREstatico As Boolean
Public lIntPagoEfectivo As Boolean
Public lPagoDirectoCD As Boolean
Public lCajaContingencia As Boolean
Public TextoPagoEfectivo As String


' ELDC: 11/01/2021
'****** Se agrego las variables publicas generales nPorcentajePropina, lImprimePropinaSugerida *************
Public lImprimePropinaSugerida As Boolean
Public lActivarAvisoRetenciones As Boolean
Public nPorcentajePropina As Integer

Public lActivaTresDecimales As Boolean
Public lIntegracionCASSWII30CD As Boolean
Public lNoVisualizarItemCombos As Boolean
Public lMostrarPrecioItemsCombo As Boolean
Public lActivarMensajeCtsCorrientes As Boolean
Public lIntMobileKDS As Boolean
Public lImprimirGuiaFormatoCorp As Boolean
Public lPermitirAnularMaxDias As Boolean
Public lBloquearNCsinEstado As Boolean
Public lImprimirFormaPago As Boolean
Public lFETusFacturasAPP As Boolean
Public lVistaPreviaItemsCombo As Boolean
Public lActivaRestriClienteDelivery As Boolean



'GCAA 15/03/2021
Public tMontoRetencion As String
Public lOCR As Boolean

'GCAA 02/07/2021
Public lReservaPedido As Boolean
Public lMontoMaximo As Double

' ELDCQ 01/07/2021
Public sDefaultOrigenVenta As String
Public SYPROV As String
Public ActivoActualizador As Boolean
Public lImprimePropinaDoc As Boolean

'JAPN 11/08/2023
Public Mesero As String
Public Guia As String

Public lPasswordMarcacion As Boolean

Public sCodigoClienteGeneral As String

Public lBloqueoAnulaPedido As Boolean
Public lBloqueoAnulaItemsPedido As Boolean
Public lOcultaRazonSocial As Boolean

Public SizeTamanoPiePrecuenta As Integer

Public lBloqueoCanal1 As Boolean
Public lBloqueoCanal2 As Boolean
Public lBloqueoCanal3 As Boolean
Public lBloqueoCanal4 As Boolean
Public lBloqueoCanal5 As Boolean
Public lbloqueoEdicionPedidoConsolaAgr As Boolean

Public lCambiarContrasena As Boolean
Public lDatosSeparados As Boolean


Public lImprimePrecuentaCabunica As Boolean
Public lBloquearPedidoSinMesa As Boolean
Public lAutorizacionSinMeza As Boolean
Public lSolicitaClaveEgreso As Boolean
Public lActivaTotemMesa As Boolean

Public locultaimpuesto2 As Boolean
Public locultaimpuesto3 As Boolean

Public lActivaPasswordMulticajero As Boolean
Public lActivaReciboSaldoQr As Boolean
Public lImprimePrepagoPrecuenta As Boolean
Public teclaFunction As String
Public mensajeImpresora As String

'Voucher y InforestAppp
Public lTurnoCompartido As Boolean
Public lActivaVoucher As Boolean

'Alerta de Consola
Public lAlertaConsola As Boolean

'Valida Formulario Anticipo
Public bCargaFormulario As Boolean
