VERSION 5.00
Begin VB.Form frmCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Documento"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "frmCambio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   " Tipo Documento "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   30
      TabIndex        =   9
      Top             =   1110
      Width           =   3945
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   525
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   525
         Index           =   2
         Left            =   1380
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   525
         Index           =   3
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCliente 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   750
      Width           =   2415
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2220
      Width           =   1245
   End
   Begin VB.TextBox txtNuevo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   390
      Width           =   1845
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   0
      Left            =   2730
      Picture         =   "frmCambio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2220
      Width           =   1245
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   1410
      Picture         =   "frmCambio.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2220
      Width           =   1245
   End
   Begin VB.TextBox txtActual 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "F-00001-000000001"
      Top             =   30
      Width           =   1845
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageCab 
      Height          =   135
      Left            =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      Height          =   195
      Index           =   2
      Left            =   930
      TabIndex        =   8
      Top             =   810
      Width           =   570
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Documento Nuevo :"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Documento Actual :"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1410
   End
End
Attribute VB_Name = "frmCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   'validacionMontoMaximo
Dim nMontoPedidoFacturar As Double
Dim nMontoDocumento As Double
   'validacionMontoMaximo
Dim RsTipoDocumento As Recordset
Dim sNuevo As String
Dim sSerie As String
Dim sCorrela As String
Dim sPrefijo As String
Dim sTD As String
Dim stipodocumento As String
Dim sImp As String
Dim sClienteNuevo As String
Dim lCliente As Boolean
Dim rstFuente As ADODB.Recordset
Dim numeroSerieImpresora As String 'SUNAT
Dim codigoImpresora As String 'SUNAT

'============================================= extranjero bolivia
Dim tAutorizacion As String
Dim tcodigoControl As String
Dim tDosificacion As String
Dim tIdentidadNIT As String

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 'Aceptar
               If lCliente And (sClienteNuevo = "" Or Len(Trim(sClienteNuevo)) = 0) Then
                  MsgBox "Error : Documento sin Cliente", vbExclamation, sMensaje
                  Exit Sub
               End If
               If txtNuevo.Text = "" Then
                  Exit Sub
               End If
                           'validacionMontoMaximo
            nMontoPedidoFacturar = 0
            nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
            nMontoDocumento = Calcular("select nventa as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
            If nMontoPedidoFacturar > 0 Then
            If nMontoPedidoFacturar < nMontoDocumento Then
                MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
                wEnter = False
                Exit Sub
            End If
            End If
               'validacionMontoMaximo
               frmKeyBoard.Caption = "Motivo del Cambio"
               frmKeyBoard.Show vbModal
               If Not wEnter And Len(Trim(sDescrip)) = 0 Then
                  Exit Sub
               End If
               
               
               


               Select Case pais
                     Case "001" 'Bolivia
                         tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                         tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                         If tAutorizacion <> "" And tDosificacion <> "" Then
                            Else
                            MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                            Exit Sub
                         End If
                         
                     Case Else 'Peru, Ecuador
                         tAutorizacion = ""
                         tcodigoControl = ""
                         tDosificacion = ""
               End Select
               
               
               
               
               
               
               
               
               
               
               
               
               
               Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTD & "' and tCaja ='" & sCaja & "'"
               Cn.Execute "Update DPAGODOCUMENTO set tdocumento ='" & sNuevo & "' where tDocumento = '" & sDocumento & "'"
               Cn.Execute "Update DPEDIDO set tDocumento = '" & sNuevo & "' where tDocumento = '" & sDocumento & "'"
               
               Isql = "Insert into DDOCUMENTO " & _
                      "       ( tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta ) " & _
                      "select '" & sNuevo & "' as tDocumento , tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial, nImpuesto1, nImpuesto2, nImpuesto3, nVenta From DDOCUMENTO where tDocumento ='" & sDocumento & "'"
               Cn.Execute Isql
               
               Dim nTotal As Double
                nTotal = Calcular("select isnull(nventa,0) as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                                Select Case pais
                                    Case "001"
                                        tcodigoControl = devuelveCodigoControl(sCaja, Mid(sNuevo, 7, 9), tAutorizacion, tDosificacion, sClienteNuevo, nTotal)
                                End Select
               
               Isql = "Insert into MDOCUMENTO " & _
                      "       ( tDocumento, tTipoDocumento, tCodigoCliente, tCortesia, nNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nVenta, nRecargo, nDescuento, nPrecioOficial, nPropina, nTotal, nAbono, nVuelto, tEstadoDocumento, tClientePago, tTurno, fPago, tCaja, tSalon, " & _
                      "tUsuario, tUsuarioAutoriza, fRegistro, tObservacion ,fDiaContable, tautorizacion,tcodigocontrol,lreplica) " & _
                      "select '" & sNuevo & "' as tDocumento , '" & sTD & "' as tTipoDocumento, '" & sClienteNuevo & "' as tClienteDelivery, tCortesia, nNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, " & _
                      "nVenta, nRecargo, nDescuento, nPrecioOficial, nPropina, nTotal, nAbono, nVuelto, tEstadoDocumento, tClientePago, '" & sTurno & "' as tTurno, fPago, '" & sCaja & "' as tCaja, '" & sSalon & "' as tSalon, " & _
                      "'" & sUsuario & "' as tUsuario, '" & sUsuario & "' as tUsuarioAutoriza, getdate() as fRegistro, tObservacion , fdiacontable, '" & tAutorizacion & "' ,'" & tcodigoControl & "',1 From MDOCUMENTO where tDocumento ='" & sDocumento & "'"
               Cn.Execute Isql
                                             
                
                           

                                
               Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tObservacion ='" & sDescrip & "', " & _
                          "tUsuarioAnulado='" & sUsuario & "', fRegistroAnulado= getdate(),lreplica=1  where tDocumento = '" & sDocumento & "'"
         
                
         
                
         
'         '--------------------SUNAT----------------------------------------------
'              If RsTipoDocumento.EOF Then
'                'SUNAT
'                 numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, "001")
'                 codigoImpresora = "001"
'
'              Else
'                'SUNAT
'                 numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, RsTipoDocumento!timpresora)
'                 codigoImpresora = RsTipoDocumento!timpresora
'
'              End If
'        Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & RsSuma!tDocumento & "' "
'        '--------------------------------------------------------------------
                    
                'impresion imagen
                Set rstFuente = New ADODB.Recordset
                imageCab.Picture = Nothing
                imagepIE.Picture = Nothing
                Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                imageCab.DataField = "foto"
                Set imageCab.DataSource = rstFuente
                imagepIE.DataField = "fotoPie"
                Set imagepIE.DataSource = rstFuente
                    

               'Impresion
               Dim wConsumo As Boolean
               Dim nRespuesta As Integer
               Dim RsEmision As Recordset
               Dim RsImpresion As Recordset
               wConsumo = False
               Set RsEmision = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' order by tTipoEmision", Cn)
               nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNo + vbDefaultButton1, sMensaje)
               If Mid(sNuevo, 1, 1) <> "0" Then
                  If nRespuesta = vbYes Then
                     frmKeyBoard.txtResultado = "POR CONSUMO"
                     frmKeyBoard.Show vbModal
                     If sDescrip = "" Or Not wEnter Then
                        MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
                        Exit Sub
                     End If
                     wConsumo = True
                  ElseIf nRespuesta = vbCancel Then
                     Exit Sub
                  End If
               End If
               
               Screen.MousePointer = vbHourglass
               Cn.Execute "Update MDOCUMENTO set tEmision = '" & IIf(wConsumo, "C", "D") & "', tConsumo ='" & IIf(wConsumo, sDescrip, "") & "' " & _
                          ",lreplica=1  where tDocumento = '" & sNuevo & "'"
               
               'Imprime Documentos
               Isql = "select * from vDocumentoImpresora " & _
                      "where tDocumento ='" & sNuevo & "'"
               Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                
               If RsImpresion.RecordCount <> 0 Then
                  'Configura la Impresora
                  RsEmision.MoveFirst
                  sPedido = RsImpresion!tCodigoPedido
                  RsEmision.Find ("Prefijo='" & Mid(sNuevo, 1, 1) & "'")
                  If RsEmision.EOF Then
                    'SUNAT
                    numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, "001")
                    codigoImpresora = "001"
                  
                     Imprimir ("001")
                     Printer.FontName = sFont
                     Printer.FontBold = False
                  Else
                    'SUNAT
                    numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, RsTipoDocumento!timpresora)
                    codigoImpresora = RsTipoDocumento!timpresora

                     Imprimir (RsEmision!timpresora)
                     Printer.FontName = sFont
                     Printer.FontBold = False
                  End If
                    'SUNAT
                    Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sNuevo & "' "

                  If RsEmision!tFormulario = "01" Then
                     Select Case Mid(RsImpresion!tDocumento, 1, 1)
                            Case Is = "F"
                                 If wConsumo Then
                                    ImprimeFacturaConsumoT RsImpresion, sDescrip, sTD, imageCab, imagepIE
                                 Else
                                    ImprimeFacturaT RsImpresion, sTD, imageCab, imagepIE
                                 End If
                             Case Is = "0"
                                 ImprimeCortesia RsImpresion, sTD, imageCab, imagepIE
                            Case Else
                                 If wConsumo Then
                                    ImprimeBoletaConsumoT RsImpresion, sDescrip, sTD, imageCab, imagepIE
                                 Else
                                    ImprimeBoletaT RsImpresion, sTD, imageCab, imagepIE
                                 End If
                     End Select
                     
                  ElseIf RsEmision!tFormulario = "02" Then
                     Select Case Mid(RsImpresion!tDocumento, 1, 1)
                            Case Is = "F"
                                 If wConsumo Then
                                    ImprimeFacturaConsumoN RsImpresion, sDescrip, sEmpresa
                                 Else
                                    ImprimeFacturaN RsImpresion, sEmpresa, stipodocumento
                                 End If
                             Case Is = "0"
                                 ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                             Case Else
                                 If wConsumo Then
                                    ImprimeBoletaConsumoN RsImpresion, sDescrip, sEmpresa
                                 Else
                                    ImprimeBoletaN RsImpresion, sEmpresa, stipodocumento
                                 End If
                     End Select
                     
                  Else
                     Select Case Mid(RsImpresion!tDocumento, 1, 1)
                             Case Is = "F"
                                 If wConsumo Then
                                    ImprimeFacturaVariableConsumo RsImpresion, sDescrip, sEmpresa
                                 Else
                                    ImprimeFacturaVariable RsImpresion, sEmpresa
                                 End If
                             Case Is = "0"
                                 ImprimeCortesia RsImpresion, sTD, imageCab, imagepIE
                             Case Else
                                 If wConsumo Then
                                    ImprimeBoletaConsumoT RsImpresion, sDescrip, sTD, imageCab, imagepIE
                                 Else
                                    ImprimeBoletaT RsImpresion, sTD, imageCab, imagepIE
                                 End If
                     End Select
                                    
                  End If
                  
               End If
               Set RsEmision = Nothing
               Set RsImpresion = Nothing
               Screen.MousePointer = vbDefault
               
               MsgBox "Documento Cambiado satisfactoriamente", vbInformation, sMensaje
               sDocumento = sNuevo
               wEnter = True
               Unload Me
               
          Case Is = 1 'Cancelar
               wEnter = False
               Unload Me
               
          Case Is = 2 'Cliente
               sTemp = ""
               Isql = "SELECT * from vCliente where lActivo = 1 Order by Descripcion"
               frmBusquedaRapida.cmdOpcion(1).Enabled = True
               frmBusquedaRapida.cmdOpcion(2).Enabled = True
               frmBusquedaRapida.cmdOpcion(3).Enabled = True
               Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                               "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                                               "Cliente", 2, "Descripcion", 5500, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True Then
                  sCliente = sCodigo
                  sClienteNuevo = sCodigo
                  txtCliente.Text = sDescrip
               Else
                  sCliente = ""
                  sClienteNuevo = ""
                  txtCliente.Text = ""
               End If
   End Select
End Sub

Private Sub cmdTipoDocumento_Click(Index As Integer)
    'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    End If
            
    'Consistencia Factura
    If Not RsTipoDocumento!Cliente Then
       lCliente = False
       sCliente = ""
       sClienteNuevo = ""
       txtCliente.Text = "SIN CLIENTE"
    Else
       lCliente = True
    End If
    
    If RsTipoDocumento!Cliente And (sCliente = "" Or sCliente = "") Then
       cmdOpcion_Click (2)
       If wEnter = False Then
          MsgBox "Error: No tiene Cliente", vbCritical, sMensaje
          Exit Sub
       End If
    'imprimedni
    Else
        'imprimedni
        Dim RsTc As ADODB.Recordset
        Set RsTc = New ADODB.Recordset
        Set RsTc = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & RsTipoDocumento!TTipoEmision & "','" & sCliente & "'", Cn)
        If Not (RsTc.EOF Or RsTc.BOF) Then
            RsTc.MoveFirst
            If RsTc.Fields(0) <> "ok" Then
                MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                Exit Sub
            End If
        End If
    End If
                                                        
    'Genera los Numero de Documento
    sSerie = RsTipoDocumento!tSerie
    sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
    sPrefijo = RsTipoDocumento!prefijo
    stipodocumento = RsTipoDocumento!TTipoEmision
    sTD = RsTipoDocumento!TTipoEmision
    sImp = RsTipoDocumento!timpresora
    sNuevo = sPrefijo & sSerie & sCorrela
    sResumen = RsTipoDocumento!lResumen
    txtNuevo.Text = Format(sNuevo, "@-@@@@@-@@@@@@@@@")
    Screen.MousePointer = vbDefault
 
End Sub

Private Sub Form_Load()
   Centrar Me
   'Tipo de Documentos
   wEnter = False
   
   If pais = "002" Then 'Ecuador
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 0 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
   Else
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 order by tTipoEmision", Cn)
   End If
   
   AsignaComando 3, RsTipoDocumento, cmdTipoDocumento()
   txtActual.Text = Format(sDocumento, "@-@@@@@-@@@@@@@@@")
   txtCliente.Text = sDescrip
   sCliente = sCodigo
   sClienteNuevo = sCodigo
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsTipoDocumento = Nothing
   Set frmCambio = Nothing
End Sub
