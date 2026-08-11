VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cmdDocumento            As ADODB.Command
Dim WithEvents RsDocumento  As ADODB.Recordset
Attribute RsDocumento.VB_VarHelpID = -1

Dim cmdDocumentoObservado            As ADODB.Command
Dim WithEvents RsDocumentoObservado  As ADODB.Recordset
Attribute RsDocumentoObservado.VB_VarHelpID = -1

Dim cmdPago           As ADODB.Command
Dim WithEvents RsPago As ADODB.Recordset
Attribute RsPago.VB_VarHelpID = -1

Public Sub HistorialCompra(ByVal hcCorrelativo As String, ByVal hcAccion As String)
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "spHistorialCompra"
         .Parameters.Refresh
         .Parameters("@vch_tCorrelativo") = hcCorrelativo
         .Parameters("@Accion") = hcAccion
    End With
    cmdDocumento.Execute
End Sub

Public Function VerificaModificacion(ByVal tCorrelativo As String) As Integer
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "sp_VerificaModifDoc"
         .Parameters.Refresh
         .Parameters("@vch_tCorrelativo") = tCorrelativo
         .Parameters("@int_Dif") = 0
    End With
    cmdDocumento.Execute
    VerificaModificacion = cmdDocumento.Parameters("@int_Dif").value
End Function

Public Sub InsmDocumento(ByVal hctDocumento As String, ByVal hcfEmision As Date, ByVal hcfIngreso As Date, ByVal hcnTotal As Double, ByVal hcnDescuento As Double, ByVal hcnNeto As Double, ByVal hcnImpuesto1 As Double, ByVal hcnImpuesto2 As Double, ByVal hcnImpuesto3 As Double, ByVal hctCodigoProveedor As String, ByVal hctRequerimiento As String, ByVal hctUsuario As String, ByVal hctEstadoDocumento As String, ByVal hcfProgramacion As String, ByVal hctCorrelativo As String, ByVal hcnCambio As Double, _
                         ByVal hctGlosa As String, ByVal hctTipoIngreso As String, ByVal hctMoneda As String, ByVal hctCodigoArea As String, ByVal hcfRegistro As String, ByVal hctProcesa As String, ByVal hcfProcesa As String, ByVal hcfCancelado As String, ByVal hctDocumentoReferencia As String, ByVal hcnSaldo As Double, ByVal hctTipoDocumento As String, ByVal hctTipoDocumentoRef As String, ByVal hclCanje As Boolean, ByVal hclDescuento As Boolean, ByVal hcnRedondeo As Double, ByVal hctCodigoOperacion As String)
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpInsmDocumentoC"
         .Parameters.Refresh
         .Parameters("@vch_tDocumento") = hctDocumento
         .Parameters("@dtm_fEmision") = hcfEmision
         .Parameters("@dtm_fIngreso") = hcfIngreso
         .Parameters("@num_nTotal") = hcnTotal
         .Parameters("@num_nDescuento") = hcnDescuento
         .Parameters("@num_nNeto") = hcnNeto
         .Parameters("@num_nImpuesto1") = hcnImpuesto1
         .Parameters("@num_nImpuesto2") = hcnImpuesto2
         .Parameters("@num_nImpuesto3") = hcnImpuesto3
         .Parameters("@vch_tCodigoProveedor") = hctCodigoProveedor
         .Parameters("@vch_tRequerimiento") = hctRequerimiento
         .Parameters("@vch_tUsuario") = hctUsuario
         .Parameters("@vch_tEstadoDocumento") = hctEstadoDocumento
         .Parameters("@dtm_fProgramacion") = IIf(hcfProgramacion = "", Null, CDate(hcfProgramacion))
         .Parameters("@vch_tCorrelativo") = hctCorrelativo
         .Parameters("@num_nCambio") = hcnCambio
         .Parameters("@vch_tGlosa") = hctGlosa
         .Parameters("@vch_tTipoIngreso") = hctTipoIngreso
         .Parameters("@vch_tMoneda") = hctMoneda
         .Parameters("@vch_tCodigoArea") = hctCodigoArea
         .Parameters("@dtm_fRegistro") = hcfRegistro
         .Parameters("@vch_tProcesa") = hctProcesa
         .Parameters("@dtm_fProcesa") = IIf(hcfProcesa = "", Null, CDate(hcfProcesa))
         .Parameters("@dtm_fCancelado") = IIf(hcfCancelado = "", Null, CDate(hcfCancelado))
         .Parameters("@vch_tDocumentoReferencia") = hctDocumentoReferencia
         .Parameters("@num_nSaldo") = hcnSaldo
         .Parameters("@vch_tTipoDocumento") = hctTipoDocumento
         .Parameters("@vch_tTipoDocumentoRef") = hctTipoDocumentoRef
         .Parameters("@bln_lCanje") = hclCanje
         .Parameters("@bln_lDescuento") = hclDescuento
         .Parameters("@num_nRedondeo") = hcnRedondeo
         .Parameters("@vch_tCodigoOperacion") = hctCodigoOperacion
         .Parameters("@bln_lReplica") = True
    End With
    cmdDocumento.Execute
End Sub

Public Function InsUpdDocumentoC(ByVal hctDocumento As String, ByVal hcfEmision As Date, ByVal hcfIngreso As Date, ByVal hcnTotal As Double, ByVal hcnDescuento As Double, ByVal hcnNeto As Double, ByVal hcnImpuesto1 As Double, ByVal hcnImpuesto2 As Double, ByVal hcnImpuesto3 As Double, ByVal hctCodigoProveedor As String, ByVal hctRequerimiento As String, ByVal hctUsuario As String, ByVal hctEstadoDocumento As String, ByVal hcfProgramacion As String, ByVal hctCorrelativo As String, ByVal hcnCambio As Double, ByVal hctGlosa As String, _
                                 ByVal hctTipoIngreso As String, ByVal hctMoneda As String, ByVal hctCodigoArea As String, ByVal hcfRegistro As String, ByVal hctProcesa As String, ByVal hcfProcesa As String, ByVal hcfCancelado As String, ByVal hctDocumentoReferencia As String, ByVal hcnSaldo As Double, ByVal hctTipoDocumento As String, ByVal hctTipoDocumentoRef As String, ByVal hclCanje As Boolean, ByVal hclDescuento As Boolean, ByVal hcnRedondeo As Double, ByVal hctCodigoOperacion As String, ByVal hctPeriodo As String, ByVal fdiacontable As String, ByVal tSede As String, ByVal tCodigoControl As String, ByVal hctAutorizacion As String, ByVal hctCodProgramaPago As String, ByVal hcnTransporte As Double, ByVal tLibroCompra As String, ByVal fFechaCaducidad As String, _
                                 ByVal hctAnioC As String, ByVal hctMesC As String)
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "InsUpdDocumentoC"
         .Parameters.Refresh
         .Parameters("@vch_tDocumento") = hctDocumento
         .Parameters("@dtm_fEmision") = hcfEmision
         .Parameters("@dtm_fIngreso") = hcfIngreso
         .Parameters("@num_nTotal") = hcnTotal
         .Parameters("@num_nDescuento") = hcnDescuento
         .Parameters("@num_nNeto") = hcnNeto
         .Parameters("@num_nImpuesto1") = hcnImpuesto1
         .Parameters("@num_nImpuesto2") = hcnImpuesto2
         .Parameters("@num_nImpuesto3") = hcnImpuesto3
         .Parameters("@vch_tCodigoProveedor") = hctCodigoProveedor
         .Parameters("@vch_tRequerimiento") = hctRequerimiento
         .Parameters("@vch_tUsuario") = hctUsuario
         .Parameters("@vch_tEstadoDocumento") = hctEstadoDocumento
         .Parameters("@dtm_fProgramacion") = IIf(hcfProgramacion = "", Null, hcfProgramacion)
         .Parameters("@vch_tCorrelativo") = hctCorrelativo
         .Parameters("@num_nCambio") = hcnCambio
         .Parameters("@vch_tGlosa") = hctGlosa
         .Parameters("@vch_tTipoIngreso") = hctTipoIngreso
         .Parameters("@vch_tMoneda") = hctMoneda
         .Parameters("@vch_tCodigoArea") = hctCodigoArea
         .Parameters("@dtm_fRegistro") = hcfRegistro
         .Parameters("@vch_tProcesa") = IIf(hctProcesa = "NULL", Null, hctProcesa)
         .Parameters("@dtm_fProcesa") = IIf(hcfProcesa = "NULL", Null, hcfProcesa)
         .Parameters("@dtm_fCancelado") = IIf(hcfCancelado = "NULL", Null, hcfCancelado)
         .Parameters("@vch_tDocumentoReferencia") = hctDocumentoReferencia
         .Parameters("@num_nSaldo") = hcnSaldo
         .Parameters("@vch_tTipoDocumento") = hctTipoDocumento
         .Parameters("@vch_tTipoDocumentoRef") = hctTipoDocumentoRef
         .Parameters("@bln_lCanje") = hclCanje
         .Parameters("@bln_lDescuento") = hclDescuento
         .Parameters("@num_nRedondeo") = hcnRedondeo
         .Parameters("@vch_tCodigoOperacion") = hctCodigoOperacion
         .Parameters("@bln_lDetraccion") = 0
         .Parameters("@num_nPercepcion") = 0
         .Parameters("@vch_Periodo") = hctPeriodo
         .Parameters("@bln_lCanjeDoc") = 0
         .Parameters("@dtm_fDiaContable") = IIf(fdiacontable = "NULL", Null, fdiacontable)
         .Parameters("@vch_tSede") = tSede
         .Parameters("@vch_tCodigoControl") = tCodigoControl
         .Parameters("@vch_tAutorizacion") = hctAutorizacion
         .Parameters("@vch_tCodProgramaPago") = hctCodProgramaPago
         .Parameters("@vch_nTransporte") = hcnTransporte
         .Parameters("@vch_tLibroCompra") = tLibroCompra
         .Parameters("@vch_tAnioC") = hctAnioC
         .Parameters("@vch_tMesC") = hctMesC
         .Parameters("@vch_fFechaCaducidad") = CDate(fFechaCaducidad)
         .Parameters("@Accion") = "I"
    End With
    cmdDocumento.Execute
    InsUpdDocumentoC = cmdDocumento.Parameters("@vch_tCorrelativo").value
End Function

Public Function InsUpdmDocumento(ByVal hctDocumento As String, ByVal hcfEmision As Date, ByVal hcfIngreso As Date, ByVal hcnTotal As Double, ByVal hcnDescuento As Double, ByVal hcnNeto As Double, ByVal hcnImpuesto1 As Double, ByVal hcnImpuesto2 As Double, ByVal hcnImpuesto3 As Double, ByVal hctCodigoProveedor As String, ByVal hctRequerimiento As String, ByVal hctUsuario As String, ByVal hctEstadoDocumento As String, ByVal hcfProgramacion As String, ByVal hctCorrelativo As String, ByVal hcnCambio As Double, ByVal hctGlosa As String, _
                                  ByVal hctTipoIngreso As String, ByVal hctMoneda As String, ByVal hctCodigoArea As String, ByVal hcfRegistro As String, ByVal hctProcesa As String, ByVal hcfProcesa As String, ByVal hcfCancelado As String, ByVal hctDocumentoReferencia As String, ByVal hcnSaldo As Double, ByVal hctTipoDocumento As String, ByVal hctTipoDocumentoRef As String, ByVal hclCanje As Boolean, ByVal hclDescuento As String, ByVal hcnRedondeo As Double, ByVal hctCodigoOperacion As String, ByVal hclDetraccion As Boolean, ByVal hclPercepcion As Double, _
                                  ByVal hctCodigoControl As String, ByVal hctCodigoAutorizacion As String, ByVal ICE As String, ByVal fdiacontable As String, ByVal hctUsuAutoriza As String, ByVal hctSede As String, ByVal hcfFechaCaducidad As String, ByVal hctCodProgramaPago As String, ByVal hcnTransporte As Double, ByVal hctLibroCompra As String, ByVal hctAnioC As String, ByVal hctMesC As String, ByVal hctAccion As String)
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpInsUpdmDocumentoC"
         .Parameters.Refresh
         .Parameters("@vch_tDocumento") = hctDocumento
         .Parameters("@dtm_fEmision") = hcfEmision
         .Parameters("@dtm_fIngreso") = hcfIngreso
         .Parameters("@num_nTotal") = hcnTotal
         .Parameters("@num_nDescuento") = hcnDescuento
         .Parameters("@num_nNeto") = hcnNeto
         .Parameters("@num_nImpuesto1") = hcnImpuesto1
         .Parameters("@num_nImpuesto2") = hcnImpuesto2
         .Parameters("@num_nImpuesto3") = hcnImpuesto3
         .Parameters("@vch_tCodigoProveedor") = hctCodigoProveedor
         
         .Parameters("@vch_tRequerimiento") = hctRequerimiento
         .Parameters("@vch_tUsuario") = hctUsuario
         .Parameters("@vch_tEstadoDocumento") = hctEstadoDocumento
         .Parameters("@dtm_fProgramacion") = IIf(hcfProgramacion = "", Null, hcfProgramacion)
         .Parameters("@vch_tCorrelativo") = hctCorrelativo
         .Parameters("@num_nCambio") = hcnCambio
         .Parameters("@vch_tGlosa") = hctGlosa
         .Parameters("@vch_tTipoIngreso") = hctTipoIngreso
         .Parameters("@vch_tMoneda") = hctMoneda
         .Parameters("@vch_tCodigoArea") = hctCodigoArea
         
         .Parameters("@dtm_fRegistro") = IIf(hcfRegistro = "", Null, hcfRegistro)
         .Parameters("@vch_tProcesa") = IIf(hctProcesa = "", Null, hctProcesa)
         .Parameters("@dtm_fProcesa") = IIf(hcfProcesa = "", Null, hcfProcesa)
         .Parameters("@dtm_fCancelado") = IIf(hcfCancelado = "", Null, hcfCancelado)
         .Parameters("@vch_tDocumentoReferencia") = hctDocumentoReferencia
         .Parameters("@num_nSaldo") = hcnSaldo
         .Parameters("@vch_tTipoDocumento") = hctTipoDocumento
         .Parameters("@vch_tTipoDocumentoRef") = hctTipoDocumentoRef
         .Parameters("@bln_lCanje") = hclCanje
         .Parameters("@bln_lDescuento") = IIf(hclDescuento = "Null", Null, IIf(hclDescuento = "1", 1, 0))
         
         .Parameters("@num_nRedondeo") = hcnRedondeo
         .Parameters("@vch_tCodigoOperacion") = hctCodigoOperacion
         .Parameters("@bln_lDetraccion") = hclDetraccion
         .Parameters("@num_nPercepcion") = hclPercepcion
         .Parameters("@vch_Periodo") = ""
         .Parameters("@bln_lCanjeDoc") = 0
         .Parameters("@vch_CodigoControl") = hctCodigoControl
         .Parameters("@vch_CodigoAutorizacion") = hctCodigoAutorizacion
         .Parameters("@vch_ICE") = ICE
         .Parameters("@dtm_fDiaContable") = IIf(fdiacontable = "", Null, fdiacontable)
         .Parameters("@vch_UsuarioAutorizaMod") = hctUsuAutoriza
         .Parameters("@vch_tSede") = hctSede
         .Parameters("@dtm_fFechaCaducidad") = "1999-01-01"
         .Parameters("@vch_tCodProgramaPago") = hctCodProgramaPago
         .Parameters("@num_nTransporte") = hcnTransporte
         .Parameters("@num_tLibroCompra") = hctLibroCompra
         .Parameters("@vch_tAnioC") = hctAnioC
         .Parameters("@vch_tMesC") = hctMesC
         .Parameters("@Accion") = hctAccion
         .Parameters("@vchMsgError") = ""
    End With
    cmdDocumento.Execute
    'InsUpdmDocumento = cmdDocumento.Parameters("@vch_tCorrelativo").value
    InsUpdmDocumento = cmdDocumento.Parameters("@vchMsgError").value
End Function

Public Sub CambiaEstadoDocumento(ByVal cetCorrelativo As String, ByVal cetUsuario As String, ByVal cetEstado As String, ByVal cefProcesa As String, ceEmision As String, ceCual As Integer, ceDiaContable As String)
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpCambiaEstadoDocumento"
         .Parameters.Refresh
         .Parameters("@vch_tCorrelativo") = IIf(cetCorrelativo = "", Null, cetCorrelativo)
         .Parameters("@vch_tUsuario") = IIf(cetUsuario = "", Null, cetUsuario)
         .Parameters("@vch_tEstado") = IIf(cetEstado = "", Null, cetEstado)
         If cefProcesa = "" Then
            .Parameters("@dat_fProcesa") = Null
         Else
            .Parameters("@dat_fProcesa") = CDate(cefProcesa)
         End If
         If ceEmision = "" Then
            .Parameters("@dat_Emision") = Null
         Else
            .Parameters("@dat_Emision") = CDate(ceEmision)
         End If
         If ceDiaContable = "" Then
            .Parameters("@dat_fDiaContable") = Null
         Else
            .Parameters("@dat_fDiaContable") = CDate(ceDiaContable)
         End If
         .Parameters("@int_Cual") = ceCual
    End With
    cmdDocumento.Execute
End Sub

Public Sub InsdDocumento(ByVal idtDocumento As String, ByVal idtCodigoProveedor As String, ByVal idnItem As Integer, ByVal idtCodigoProducto As String, ByVal idtUnidadEntrada As String, _
                         ByVal idnCantidad As Double, ByVal idnPrecio As Double, ByVal idnPVenta As Double, ByVal idnNeto As Double, ByVal idnDescuento As Double, ByVal idnRecargo As Double, _
                         ByVal idnImpuesto1 As Double, ByVal idnImpuesto2 As Double, ByVal idnImpuesto3 As Double, ByVal idnPorcentaje1 As Double, ByVal idnPorcentaje2 As Double, _
                         ByVal idnPorcentaje3 As Double, ByVal idnTotal As Double, ByVal idtEstadoDocumento As String, ByVal idtTipoIngreso As String, ByVal idlIncluidoImpuesto1 As Boolean, _
                         ByVal idlIncluidoImpuesto2 As Boolean, ByVal idlIncluidoImpuesto3 As Boolean, ByVal idtCorrelativo As String, ByVal idtCuentacontable As String, ByVal idtCentroCosto As String, ByVal tAreaOrigen As String, ByVal idAccion As String, _
                         ByVal lAsumirIUE_IT As Boolean, ByVal nRetencion1 As Double, ByVal nRetencion2 As Double, ByVal nRetencion3 As Double, ByVal nPRetencion1 As Double, ByVal nPRetencion2 As Double, ByVal nPRetencion3 As Double, ByVal lIncluidoRetencion1 As Boolean, _
                         ByVal lIncluidoRetencion2 As Boolean, ByVal lIncluidoRetencion3 As Boolean, ByVal nICE As Double, ByVal nOtrosCargosInafecto As Double)
                         
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpInsUpdDDOCUMENTO"
         .Parameters.Refresh
         .Parameters("@vch_tDocumento") = idtDocumento
         .Parameters("@vch_tCodigoProveedor") = idtCodigoProveedor
         .Parameters("@num_nItem") = idnItem
         .Parameters("@vch_tCodigoProducto") = idtCodigoProducto
         .Parameters("@vch_tUnidadEntrada") = idtUnidadEntrada
         .Parameters("@num_nCantidad") = idnCantidad
         .Parameters("@num_nPrecio") = idnPrecio
         .Parameters("@num_nPVenta") = idnPVenta
         .Parameters("@num_nNeto") = idnNeto
         .Parameters("@num_nDescuento") = idnDescuento
         .Parameters("@num_nRecargo") = idnRecargo
         .Parameters("@num_nImpuesto1") = idnImpuesto1
         .Parameters("@num_nImpuesto2") = idnImpuesto2
         .Parameters("@num_nImpuesto3") = idnImpuesto3
         .Parameters("@num_nPorcentaje1") = idnPorcentaje1
         .Parameters("@num_nPorcentaje2") = idnPorcentaje2
         .Parameters("@num_nPorcentaje3") = idnPorcentaje3
         .Parameters("@num_nTotal") = idnTotal
         .Parameters("@vch_tEstadoDocumento") = idtEstadoDocumento
         .Parameters("@vch_tTipoIngreso") = idtTipoIngreso
         .Parameters("@bln_lIncluidoImpuesto1") = idlIncluidoImpuesto1
         .Parameters("@bln_lIncluidoImpuesto2") = idlIncluidoImpuesto2
         .Parameters("@bln_lIncluidoImpuesto3") = idlIncluidoImpuesto3
         .Parameters("@vch_tCorrelativo") = idtCorrelativo
         .Parameters("@vch_tCuentacontable") = idtCuentacontable
         .Parameters("@vch_tCentroCosto") = idtCentroCosto
         .Parameters("@vch_tAreaOrigen") = tAreaOrigen
         .Parameters("@bln_lAsumirIUE_IT") = lAsumirIUE_IT
         .Parameters("@num_nRetencion1") = nRetencion1
         .Parameters("@num_nRetencion2") = nRetencion2
         .Parameters("@num_nRetencion3") = nRetencion3
         .Parameters("@num_nPRetencion1") = nPRetencion1
         .Parameters("@num_nPRetencion2") = nPRetencion2
         .Parameters("@num_nPRetencion3") = nPRetencion3
         .Parameters("@bln_lIncluidoRetencion1") = lIncluidoRetencion1
         .Parameters("@bln_lIncluidoRetencion2") = lIncluidoRetencion2
         .Parameters("@bln_lIncluidoRetencion3") = lIncluidoRetencion3
         .Parameters("@num_nICE") = nICE
         .Parameters("@num_nOtrosCargosInafecto") = nOtrosCargosInafecto
'         .Parameters("@vch_tUsuario") = IIf(IsNull(sUsuario), "", sUsuario)
'         .Parameters("@vch_UsuarioAutorizaMod") = IIf(IsNull(sUsuarioAutoriza), "", sUsuarioAutoriza)
         .Parameters("@Accion") = idAccion
    End With
    cmdDocumento.Execute
End Sub

Public Function DocumentoDescanje(ByVal fCorrelativo As String)
    Set cmdDocumento = New ADODB.Command
    With cmdDocumento
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpDescanjear"
         .Parameters.Refresh
         .Parameters("@vch_tCorrelativo") = fCorrelativo
    End With
    cmdDocumento.Execute
    DocumentoDescanje = cmdDocumento.Parameters("@vch_tCorrelativo").value
End Function


Public Sub DocumentoObservacion(ByVal dotDocumento As String, ByVal dotProveedor As String, ByVal dofRegistro As String, ByVal dotCorrelativo As String, ByVal dolActivo As Boolean, ByVal doAccion As String)
    Set cmdDocumentoObservado = New ADODB.Command
    With cmdDocumentoObservado
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpInsUpdtDocumentoObservado"
         .Parameters.Refresh
         .Parameters("@vch_tDocumento") = dotDocumento
         .Parameters("@vch_tProveedor") = dotProveedor
         .Parameters("@dtm_fRegistro") = dofRegistro
         .Parameters("@vch_tCorrelativo") = dotCorrelativo
         .Parameters("@bln_lActivo") = dolActivo
         .Parameters("@Accion") = doAccion
    End With
    cmdDocumentoObservado.Execute
End Sub

Public Function ListaFecha(ByVal flfFechaIni As String, ByVal flfFechaFin As String, ByVal flfTabla As String, ByVal flfTipo As Boolean) As ADODB.Recordset
    Set cmdDocumentoObservado = New ADODB.Command
    Set RsDocumentoObservado = New ADODB.Recordset
    RsDocumentoObservado.CursorLocation = adUseClient
    With cmdDocumentoObservado
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpLisFechas"
         .Parameters.Refresh
         .Parameters("@dtm_fIni") = IIf(flfFechaIni = "Null", Null, Format(flfFechaIni, "YYYY/mm/dd 00:00:00"))
         .Parameters("@dtm_fFin") = IIf(flfFechaFin = "Null", Null, Format(flfFechaFin, "YYYY/mm/dd 23:59:59"))
         .Parameters("@Tabla") = flfTabla
         .Parameters("@Tipo") = flfTipo
    End With
    RsDocumentoObservado.Open cmdDocumentoObservado, , adOpenStatic, adLockOptimistic
    Set ListaFecha = RsDocumentoObservado
End Function

Sub InsmPago(ByVal itCodigoProveedor As String, ByVal itVoucher As String, ByVal itDocumento As String, ByVal itCodigoTransaccion As String, ByVal inImporte As Double, ByVal itMoneda As String, ByVal inCambio As Double, ByVal itTipoCancelacion As String, ByVal itUsuario As String, ByVal itCodigoCuenta As String, ByVal inNumeroCheque As String, ByVal itDocumentoReferencia As String, ByVal itCuentaContable As String, ByVal itCodigoBanco As String, ByVal ifDiferido As String, ByVal inCorrela As Double, ByVal itCorrelativo As String, ByVal itTipo As String)
    Set cmdPago = New ADODB.Command
    Set RsPago = New Recordset
    With cmdPago
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpInsmPago"
         .Parameters.Refresh
         .Parameters("@vch_tCodigoProveedor") = itCodigoProveedor
         .Parameters("@vch_tVoucher") = itVoucher
         .Parameters("@vch_tDocumento") = itDocumento
         .Parameters("@vch_tCodigoTransaccion") = itCodigoTransaccion
         .Parameters("@num_nImporte") = inImporte
         .Parameters("@vch_tMoneda") = itMoneda
         .Parameters("@num_nCambio") = inCambio
         .Parameters("@vch_tTipoCancelacion") = itTipoCancelacion
         .Parameters("@vch_tUsuario") = itUsuario
         .Parameters("@vch_tCodigoCuenta") = itCodigoCuenta
         .Parameters("@vch_nNumeroCheque") = inNumeroCheque
         .Parameters("@vch_tDocumentoReferencia") = itDocumentoReferencia
         .Parameters("@vch_tCuentaContable") = itCuentaContable
         .Parameters("@vch_tCodigoBanco") = itCodigoBanco
         .Parameters("@dtm_fDiferido") = IIf(ifDiferido = "NULL", Null, ifDiferido)
         .Parameters("@num_nCorrela") = inCorrela
         .Parameters("@vch_tCorrelativo") = itCorrelativo
         .Parameters("@tTipo") = itTipo
    End With
    cmdPago.Execute
End Sub

Sub InsmPagoParcial(ByVal itCodigoProveedor As String, ByVal itVoucher As String, ByVal itDocumento As String, ByVal itCodigoTransaccion As String, ByVal inImporte As Double, ByVal itMoneda As String, ByVal inCambio As Double, ByVal itTipoCancelacion As String, ByVal itUsuario As String, ByVal itCodigoCuenta As String, ByVal inNumeroCheque As String, ByVal itDocumentoReferencia As String, ByVal itCuentaContable As String, ByVal itCodigoBanco As String, ByVal ifDiferido As String, ByVal inCorrela As Double, ByVal itCorrelativo As String, ByVal itTipo As String)
    Set cmdPago = New ADODB.Command
    Set RsPago = New Recordset
    With cmdPago
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "SpInsmPago_Parcial"
         .Parameters.Refresh
         .Parameters("@vch_tCodigoProveedor") = itCodigoProveedor
         .Parameters("@vch_tVoucher") = itVoucher
         .Parameters("@vch_tDocumento") = itDocumento
         .Parameters("@vch_tCodigoTransaccion") = itCodigoTransaccion
         .Parameters("@num_nImporte") = inImporte
         .Parameters("@vch_tMoneda") = itMoneda
         .Parameters("@num_nCambio") = inCambio
         .Parameters("@vch_tTipoCancelacion") = itTipoCancelacion
         .Parameters("@vch_tUsuario") = itUsuario
         .Parameters("@vch_tCodigoCuenta") = itCodigoCuenta
         .Parameters("@vch_nNumeroCheque") = inNumeroCheque
         .Parameters("@vch_tDocumentoReferencia") = itDocumentoReferencia
         .Parameters("@vch_tCuentaContable") = itCuentaContable
         .Parameters("@vch_tCodigoBanco") = itCodigoBanco
         .Parameters("@dtm_fDiferido") = IIf(ifDiferido = "NULL", Null, ifDiferido)
         .Parameters("@num_nCorrela") = inCorrela
         .Parameters("@vch_tCorrelativo") = itCorrelativo
         .Parameters("@tTipo") = itTipo
    End With
    cmdPago.Execute
End Sub

Public Function LisCuentaCorrente(ByVal slccFecha As String, ByVal slccTipo As String) As ADODB.Recordset
    Set cmdPago = New ADODB.Command
    Set RsPago = New Recordset
    RsPago.CursorLocation = adUseClient
    With cmdPago
         .ActiveConnection = CnAlmacen
         .CommandType = adCmdStoredProc
         .CommandText = "spLisProveedorCtaCte"
         .Parameters.Refresh
         .Parameters("@dat_Fech") = IIf(slccFecha = "Null", Null, slccFecha)
         .Parameters("@vch_Tipo") = slccTipo
    End With
   RsPago.Open cmdPago, , adOpenStatic, adLockOptimistic
   Set LisCuentaCorrente = RsPago
End Function



