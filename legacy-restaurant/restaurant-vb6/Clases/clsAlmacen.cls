VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cmdAlmacen           As ADODB.Command
Dim cmdInforest           As ADODB.Command
Dim WithEvents RsAlmacen As ADODB.Recordset
Attribute RsAlmacen.VB_VarHelpID = -1

Function FunInsertamSubKardex(ByVal iktCodigoSubArea As String, ByVal iktCodigoProducto As String, ByVal iknCorrelativo As Integer, ByVal iktCodigoArea As String, _
                              ByVal ikfRegistro As String, ByVal iktUsuario As String, ByVal iktTipoDocumento As String, ByVal iklInterno As Boolean, ByVal iktDocumento As String, _
                              ByVal iknIngreso As Double, ByVal iknSalida As Double, ByVal iknValor As Double, ByVal iknStockActual As Double, ByVal iknStockUltimo As Double, ByVal iktEstadoKardex As String, ByVal iktEtiquetaCodigoUnico As String, _
                              ByVal iktDocumentoVenta As String, ByVal iktTipoDocumentoVenta As String, ByVal ikfDiaContable As String, ByVal ikfFechaPedido As String)
    Set cmdAlmacen = New ADODB.Command
    On Error GoTo ExecuteSp
    With cmdAlmacen
        If verificaAlmacenRemoto = "ON" Then
            .ActiveConnection = CnAlmacenRemoto
        Else
            .ActiveConnection = CnAlmacen
        End If
        
        .CommandType = adCmdStoredProc
        .CommandText = "SpInsertamSubKardex"
        .Parameters.Refresh
        .Parameters("@vch_tCodigoSubArea") = iktCodigoSubArea
        .Parameters("@vch_tCodigoProducto") = iktCodigoProducto
        .Parameters("@num_nCorrelativo") = iknCorrelativo
        .Parameters("@vch_tCodigoArea") = iktCodigoArea
        .Parameters("@dtm_fRegistro") = ikfRegistro
        .Parameters("@vch_tUsuario") = iktUsuario
        .Parameters("@vch_tTipoDocumento") = iktTipoDocumento
        .Parameters("@bln_lInterno") = iklInterno
        .Parameters("@vch_tDocumento") = iktDocumento
        .Parameters("@num_nIngreso") = iknIngreso
        .Parameters("@num_nSalida") = iknSalida
        .Parameters("@num_nValor") = iknValor
        .Parameters("@num_nStockActual") = iknStockActual
        .Parameters("@num_nStockUltimo") = iknStockUltimo
        .Parameters("@vch_tEstadoKardex") = iktEstadoKardex
        .Parameters("@vch_tEtiquetaCodigoUnico") = iktEtiquetaCodigoUnico
        .Parameters("@vch_tDocumentoVenta") = iktDocumentoVenta
        .Parameters("@vch_tTipoDocumentoVenta") = iktTipoDocumentoVenta
        .Parameters("@dtm_fDiaContableVenta") = ikfDiaContable
        .Parameters("@dtm_fFechaPedido") = ikfFechaPedido
    End With
    cmdAlmacen.Execute
    FunInsertamSubKardex = cmdAlmacen.Parameters("@vch_tCodigoProducto").Value
    Exit Function
    
ExecuteSp:
    MsgBox "Error: " & Error, vbInformation, "Sistema"
    'frmDescargo.msgError = Err.Description
End Function

Function FunInsertamKardex(ByVal ikttCodigoProducto As String, ByVal iktnCorrelativo As Integer, ByVal ikttTipoDocumento As String, ByVal iktlInterno As Boolean, ByVal ikttCodigoArea As String, ByVal ikttDocumento As String, _
                           ByVal iktnIngreso As Double, ByVal iktnValor As Double, ByVal iktnSalida As Double, ByVal iktnStockActual As Double, ByVal iktnStockUltimo As Double, ByVal ikttEstadoKardex As String, _
                           ByVal iktnPrecioPromedio As Double, ByVal iktnPrecioCosto As Double, ByVal iktfRegistro As String, ByVal ikttUsuario As String, ByVal ikttCorrelativo As String)

    Set cmdAlmacen = New ADODB.Command
    With cmdAlmacen
        If verificaAlmacenRemoto = "ON" Then
            .ActiveConnection = CnAlmacenRemoto
        Else
            .ActiveConnection = CnAlmacen
        End If
    
        .CommandType = adCmdStoredProc
        .CommandText = "SpInsertamKardex"
        .Parameters.Refresh
        .Parameters("@vch_tCodigoProducto") = ikttCodigoProducto
        .Parameters("@num_nCorrelativo") = iktnCorrelativo
        .Parameters("@vch_tTipoDocumento") = ikttTipoDocumento
        .Parameters("@bln_lInterno") = iktlInterno
        .Parameters("@vch_tCodigoArea") = ikttCodigoArea
        .Parameters("@vch_tDocumento") = ikttDocumento
        .Parameters("@num_nIngreso") = iktnIngreso
        .Parameters("@num_nValor") = iktnValor
        .Parameters("@num_nSalida") = iktnSalida
        .Parameters("@num_nStockActual") = iktnStockActual
        .Parameters("@num_nStockUltimo") = iktnStockUltimo
        .Parameters("@vch_tEstadoKardex") = ikttEstadoKardex
        .Parameters("@num_nPrecioPromedio") = iktnPrecioPromedio
        .Parameters("@num_nPrecioCosto") = iktnPrecioCosto
        .Parameters("@dtm_fRegistro") = iktfRegistro
        .Parameters("@vch_tUsuario") = ikttUsuario
        .Parameters("@vch_tCorrelativo") = ikttCorrelativo
    End With
    cmdAlmacen.Execute
    FunInsertamKardex = cmdAlmacen.Parameters("@vch_tCodigoProducto").Value
End Function


Function FunInsertaLogDescargo(ByVal iktnCorrelativo As Integer, ByVal iktfRegistro As String, ByVal ikttUsuario As String, ByVal iktFecInicio As String, ByVal iktfFecFin As String, ByVal iktnCantPedidos As Integer, _
                               ByVal iktnCantPedOk As Integer, ByVal iktnCantPedError As Integer, ByVal ikttEquIp As String, ByVal ikttEquName As String, ByVal ikttEquUser As String)

    Set cmdInforest = New ADODB.Command
    With cmdInforest
        .ActiveConnection = Cn
    
        .CommandType = adCmdStoredProc
        .CommandText = "usp_Inforest_InsertaLogDescargo"
        .Parameters.Refresh
        .Parameters("@num_nCorrelativo") = iktnCorrelativo
        .Parameters("@dtm_fRegInicio") = iktfRegistro
        .Parameters("@vch_tUsuario") = ikttUsuario
        .Parameters("@dtm_fFecInicio") = iktFecInicio
        .Parameters("@dtm_fFecFin") = iktfFecFin
        .Parameters("@num_nCantPedidos") = iktnCantPedidos
        .Parameters("@num_nCantPedOk") = iktnCantPedOk
        .Parameters("@num_nCantPedError") = iktnCantPedError
        .Parameters("@vch_tEquIp") = ikttEquIp
        .Parameters("@vch_tEquName") = ikttEquName
        .Parameters("@vch_tEquUser") = ikttEquUser
    End With
    cmdInforest.Execute
    FunInsertaLogDescargo = cmdInforest.Parameters("@num_nCorrelativo").Value
End Function

Function EjecutarDescargoAutomatico(ByVal NombreServicio As String, ByVal Estado As Boolean) As Boolean
    Set cmdInforest = New ADODB.Command
    With cmdInforest
        .ActiveConnection = Cn
    
        .CommandType = adCmdStoredProc
        .CommandText = "usp_ControlServicioWindows"
        .Parameters.Refresh
        .Parameters("@NombreServicio") = NombreServicio
        .Parameters("@Accion") = IIf(Estado, "START", "STOP")
    End With
    cmdInforest.Execute
    EjecutarDescargoAutomatico = cmdInforest.Parameters("@Resultado").Value
End Function

