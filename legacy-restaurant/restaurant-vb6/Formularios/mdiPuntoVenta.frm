VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.MDIForm mdiPuntoVenta 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Punto de Venta"
   ClientHeight    =   9240
   ClientLeft      =   -2700
   ClientTop       =   -2055
   ClientWidth     =   11760
   Icon            =   "mdiPuntoVenta.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStretch 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8235
      Left            =   0
      Picture         =   "mdiPuntoVenta.frx":57E2
      ScaleHeight     =   8235
      ScaleWidth      =   11760
      TabIndex        =   21
      Top             =   705
      Visible         =   0   'False
      Width           =   11760
      Begin VB.Timer TimerBizlink 
         Interval        =   60000
         Left            =   7920
         Top             =   240
      End
      Begin MCI.MMControl MMControl1 
         Height          =   615
         Left            =   10680
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1085
         _Version        =   393216
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         PauseVisible    =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         RecordVisible   =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Timer Timerwebapp 
         Interval        =   10000
         Left            =   6705
         Top             =   240
      End
   End
   Begin VB.Timer Timer 
      Left            =   0
      Top             =   1440
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox xPicture 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11700
      TabIndex        =   1
      Top             =   0
      Width           =   11760
      Begin VB.CommandButton cmdRecargarTarjeta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recargar Tarjeta"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9940
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdConsultaSaldo 
         Caption         =   "Consultar Saldos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9940
         Picture         =   "mdiPuntoVenta.frx":11E403
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcionMensaje 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mensajes"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8680
         Picture         =   "mdiPuntoVenta.frx":11E78D
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdAmpliar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ampliar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8680
         Picture         =   "mdiPuntoVenta.frx":11ED17
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ctas Corrient"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2490
         Picture         =   "mdiPuntoVenta.frx":11EE69
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cierre"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2490
         Picture         =   "mdiPuntoVenta.frx":11F3F3
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Carta de Productos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4930
         Picture         =   "mdiPuntoVenta.frx":11F4F5
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delivery en Transito"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6180
         Picture         =   "mdiPuntoVenta.frx":11FA7F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documentos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1250
         Picture         =   "mdiPuntoVenta.frx":120009
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion15 
         Caption         =   "Activar Pin Pad"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9940
         Picture         =   "mdiPuntoVenta.frx":120593
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   675
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedidos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         Picture         =   "mdiPuntoVenta.frx":12091D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delivery Entregados"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7420
         Picture         =   "mdiPuntoVenta.frx":120A67
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recibos de Ingresos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6180
         Picture         =   "mdiPuntoVenta.frx":120FF1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ctas x Cobrar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3720
         Picture         =   "mdiPuntoVenta.frx":1210F3
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion16 
         BackColor       =   &H00EBECFF&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9940
         Picture         =   "mdiPuntoVenta.frx":12167D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Mesas"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3720
         Picture         =   "mdiPuntoVenta.frx":12176F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Punto Venta"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1250
         Picture         =   "mdiPuntoVenta.frx":121869
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Apertura"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         Picture         =   "mdiPuntoVenta.frx":12196B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Recibos de Egresos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4930
         Picture         =   "mdiPuntoVenta.frx":121A6D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reservas"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7420
         Picture         =   "mdiPuntoVenta.frx":121B6F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
      End
      Begin VB.Image ImagePais 
         Height          =   480
         Left            =   12600
         Stretch         =   -1  'True
         Top             =   90
         Width           =   885
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8940
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "11/02/2026"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "03:26 p.m."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuMovimiento 
      Caption         =   "M&ovimientos"
      Begin VB.Menu mnuInicio 
         Caption         =   "&Apertura de Turno"
      End
      Begin VB.Menu mnuVenta 
         Caption         =   "&Punto de Venta"
      End
      Begin VB.Menu mnuPinPad 
         Caption         =   "Activar PinPad (No Financiera)"
      End
      Begin VB.Menu mnuCierre 
         Caption         =   "Ci&erre de Turno"
      End
      Begin VB.Menu linea8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorrelativo 
         Caption         =   "Correlativo de Documentos"
      End
      Begin VB.Menu linea7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMesa 
         Caption         =   "&Mesas"
      End
      Begin VB.Menu mnuInsumoCritico 
         Caption         =   "&Insumos/Platos de Stock Crítico"
      End
      Begin VB.Menu mnuImportacion 
         Caption         =   "&Importación de Requerimientos"
      End
   End
   Begin VB.Menu mnuCuentas 
      Caption         =   "&Correlativos"
      Begin VB.Menu mnuCorrelativoPedido 
         Caption         =   "Correlativo de Pedidos"
      End
      Begin VB.Menu mnuCorrelativoDocumento 
         Caption         =   "Correlativo de Documentos"
      End
      Begin VB.Menu mnuGuiaTransporte 
         Caption         =   "Correlativo de Guias de Transporte"
      End
      Begin VB.Menu mnuCtaCte 
         Caption         =   "Cuentas Corrientes"
      End
      Begin VB.Menu mnuRecibo 
         Caption         =   "Recibos de Egreso"
      End
      Begin VB.Menu mnuReciboIngreso 
         Caption         =   "Recibos de Ingreso"
      End
      Begin VB.Menu mnuNotaCredito 
         Caption         =   "Notas de Crédito"
      End
      Begin VB.Menu mnuReserva 
         Caption         =   "Reservas"
      End
      Begin VB.Menu mnuDocumentoElectronico 
         Caption         =   "Documentos Electrónicos"
      End
      Begin VB.Menu mnuCorrelativoCentralPedidos 
         Caption         =   "Pedidos Central Producción"
      End
      Begin VB.Menu mnuRecargarTarjetaProximidad 
         Caption         =   "Recargar Tarjeta de Proximidad"
      End
      Begin VB.Menu mnuTarjetaProximidad 
         Caption         =   "Tarjeta Proximidad"
      End
      Begin VB.Menu linea6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCuentaCobrar 
         Caption         =   "Cuenta por Cobrar"
      End
   End
   Begin VB.Menu mnuConexion 
      Caption         =   "C&onexión"
      Begin VB.Menu mnuCambiaLocal 
         Caption         =   "Cambiar de Local"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
      Begin VB.Menu mnuLiberar 
         Caption         =   "Liberar Descargo Venta"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
   Begin VB.Menu mnuestilos 
      Caption         =   "&Estilos"
      Visible         =   0   'False
      Begin VB.Menu mnuClasico 
         Caption         =   "Estilo Menu Clasico"
      End
      Begin VB.Menu mnuVista 
         Caption         =   "Estilo Menu Vista"
      End
      Begin VB.Menu mnuRoyal 
         Caption         =   "Estilo Menu Royal"
      End
      Begin VB.Menu mnuLuna 
         Caption         =   "Estilo Menu Luna"
      End
      Begin VB.Menu mnuEstiloOF 
         Caption         =   "Estilo Menu Office 2007"
      End
   End
End
Attribute VB_Name = "mdiPuntoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const IMAGESIZE = 0.566893424036281
Dim RsPedidosCarta As Recordset
Dim RsArea As Recordset
Dim RsImpresion As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsComboPropiedad As Recordset
Dim RsCabecerax As Recordset
Dim Contador As Long


Private Sub cmdAmpliar_Click()
    If cmdAmpliar.Caption = "Esconder" Then
       xPicture.Height = 700
       cmdAmpliar.Caption = "Ampliar"
    Else
       xPicture.Height = 1380
       cmdAmpliar.Caption = "Esconder"
    End If
End Sub

Private Sub cmdConsultaSaldo_Click()
   frmConsultaSaldo.Show vbModal
End Sub



Private Sub cmdOpcion1_Click()
    If lDiaContable = False Then 'manual
        frmDiaContable.obtieneModoIngreso "Apertura"
        frmDiaContable.Show vbModal
        StatusBar.Panels.Item(1).Text = "Día Contable : " & obtieneDiaContable
    End If


    If lDiaContable = False Then ' manual
                    If lDiaContableAperturado = True Then
                            If lMultiCajero = True Then
                                If validaInicioCajaRapida = False Then
                                        frmInicio.Show vbModal
                                End If
                            Else
                                frmInicio.Show vbModal
                                
                            End If
                    End If
            Else ' automatico
            
            If lMultiCajero = True Then
                    If validaInicioCajaRapida = False Then
                                        frmInicio.Show vbModal
                                End If
                Else
                frmInicio.Show vbModal
            End If
    End If
    
    
    
    
    StatusBar.Panels.Item(4).Text = IIf(wInicio, "Turno : " & sTurno, "Turno : No Iniciado")
    
    
    cmdOpcion1.Enabled = IIf(wInicio, False, True)
    mnuInicio.Enabled = IIf(wInicio, False, True)
End Sub

Public Function validaInicioCajaRapida() As Boolean
    validaInicioCajaRapida = True ' ingresadirecto
    Dim rsTipoCambio As New ADODB.Recordset
    
    'valida tipo cambio
    Set rsTipoCambio = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
    If rsTipoCambio.EOF Or rsTipoCambio.BOF Then
    
        validaInicioCajaRapida = False ' pide tipocambio
        Exit Function
    Else
        nTC = IIf(IsNull(rsTipoCambio!nVenta), 0, IIf(IsNull(rsTipoCambio!nVenta), 0, rsTipoCambio!nVenta))
         If nTC = 0 Then
                validaInicioCajaRapida = False ' pide tipocambio
                Exit Function
        End If
         
         
    End If
    'valida existencia de turno
    Dim rstTurno As New ADODB.Recordset
    Dim nAbonoN As Double
    Dim nAbonoE As Double
    If lMCPV Then
      Isql = "select * from MTURNO where tUsuario ='" & sUsuario & "' order by tTurno"
   Else
      Isql = "select * from MTURNO where tCaja ='" & sCaja & "' order by tTurno"
   End If
   Set rstTurno = Lib.OpenRecordset(Isql, Cn)
   If rstTurno.RecordCount = 0 Then
        validaInicioCajaRapida = False ' pide turno
        Exit Function
    Else
      rstTurno.MoveLast
                  
      If Not rstTurno!lCierre = True Then
                sTurno = rstTurno!tTurno
                nAbonoN = IIf(IsNull(rstTurno!nMontoIN), 0, rstTurno!nMontoIN)
                nAbonoE = IIf(IsNull(rstTurno!nMontoIE), 0, rstTurno!nMontoIE)
                
         
                  Isql = "update MTURNO set " & _
                         "tUsuario ='" & sUsuario & "', " & _
                         "nMontoIN = " & nAbonoN & ", " & _
                         "nMontoIE = " & nAbonoE & " " & _
                         "where tTurno ='" & sTurno & "'"
                   Cn.Execute Isql
                   ActivaInicio (True)
                    wInicio = True
                    
                    validaInicioCajaRapida = True ' ingresadirecto
       Else
            validaInicioCajaRapida = False ' pide turno
                Exit Function
       End If
    End If
End Function



Private Sub cmdOpcion10_Click()
    If nTC = 0 Then MsgBox "Es necesario aperturar el Turno!": Exit Sub
    frmCtaCte.Show
End Sub

Private Sub cmdOpcion11_Click()
    frmCuentaCobrar.Show
End Sub

Private Sub cmdOpcion12_Click()
    frmPrecios.Show
End Sub

Private Sub cmdOpcion13_Click()
    frmPedidoDelivery.Show
End Sub

Private Sub cmdOpcion14_Click()
    frmPedidoDeliveryNo.Show
End Sub

Private Sub cmdOpcion15_Click()
    Call mnuPinPad_Click
End Sub

Private Sub cmdOpcion16_Click()
    'OO---------------------------------------
    If lMCPV Then
        'Metodo para cerrar todo e Inicializar()
        frmMozoUsuario.Show vbModal
        InicializaMCPV
    Else
        'Cambiar el SQL
        Isql = " update  LOG_SESIONES SET lActivo = 1 WHERE lActivo = 0 and TCAJA ='" & sCaja & "' and PC='" & Environ("computername") & "'"
        Cn.Execute Isql
        Salir
    End If
    '---------------------------------------
End Sub



Private Sub cmdOpcion2_Click()

'Oscar Ortega
    sTipo = "TC"
    If nTC = 0 And pais <> "003" Then
        frmNumPad.Show vbModal
        If wEnter Then
            nTC = val(sDescrip)

            Dim oComando As clsComando
            Set oComando = New clsComando
            If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
                Set oComando = Nothing
                Exit Sub
            End If

            oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, nTC
            oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, sUsuario
            oComando.CreateParameter "@nTCO", adDouble, adParamInput, 0, nTCO
            oComando.CreateParameter "@nTC2", adDouble, adParamInput, 0, nTC3
            oComando.CreateParameter "@nTC3", adDouble, adParamInput, 0, nTC4

            If Not oComando.GetParamOK Then
                Set oComando = Nothing
                Exit Sub
            End If
            If Not oComando.ExecSP Then
                Set oComando = Nothing
                Exit Sub
            End If

        End If
    ElseIf (nTC = 0 Or nTC2 = 0 Or nTC3 = 0) And pais = "003" Then
        frmNumPad.Caption = "Tipo Cambio 1 " & sMonedaE
        frmNumPad.Show vbModal
        If wEnter Then
            nTC = val(sDescrip)
        Else
            MsgBox "Tipo de Cambio 1 Obligatorio"
            Exit Sub
        End If
        
        frmNumPad.Caption = "Tipo Cambio 2 " & sMonedaE2
        frmNumPad.Show vbModal
        If wEnter Then
            nTC2 = val(sDescrip)
        Else
            MsgBox "Tipo de Cambio 2 Obligatorio"
            Exit Sub
        End If
        
        frmNumPad.Caption = "Tipo Cambio 3 " & sMonedaE3
        frmNumPad.Show vbModal
        If wEnter Then
            nTC3 = val(sDescrip)
        Else
            MsgBox "Tipo de Cambio 3 Obligatorio"
            Exit Sub
        End If
        If wEnter Then
            nTC = val(sDescrip)

            Dim oComando2 As clsComando
            Set oComando2 = New clsComando
            If Not oComando2.CreateCmdSp("spIns_TipoCambio", Cn) Then
                Set oComando2 = Nothing
                Exit Sub
            End If

            oComando2.CreateParameter "@nTc", adDouble, adParamInput, 0, nTC
            oComando2.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, sUsuario
            oComando2.CreateParameter "@nTCO", adDouble, adParamInput, 0, nTCO
            oComando2.CreateParameter "@nTC2", adDouble, adParamInput, 0, nTC3
            oComando2.CreateParameter "@nTC3", adDouble, adParamInput, 0, nTC4

            If Not oComando2.GetParamOK Then
                Set oComando2 = Nothing
                Exit Sub
            End If
            If Not oComando2.ExecSP Then
                Set oComando2 = Nothing
                Exit Sub
            End If
        End If
    End If

   Screen.MousePointer = vbHourglass
   If lCajaRapida Or lMultiCajero Then
      frmCajaRapida.Show vbModal
      If Not wEnter Then
         frmVenta.Show vbModal
      End If
   Else
      frmVenta.Show vbModal

      If lMCPV Then
         frmMozoUsuario.Show vbModal
         InicializaMCPV
      End If
  End If
End Sub

Private Sub cmdOpcion3_Click()
    frmLiquidacionDetalle.Show vbModal

    If wInicio = False Then
    ActivaInicio (False)
    End If
End Sub

Private Sub cmdOpcion4_Click()
    Screen.MousePointer = vbHourglass
    sTipo = "V"
    frmMesaConsulta.Show
End Sub

Private Sub cmdOpcion5_Click()
    frmReciboEgreso.Show
End Sub

Private Sub cmdOpcion6_Click()
    frmReciboIngreso.Show
End Sub

Private Sub cmdOpcion7_Click()
    frmReserva.Show
End Sub

Private Sub cmdOpcion8_Click()
    frmPedidoCorrelativo.Show
End Sub

Private Sub cmdOpcion9_Click()
    frmDocumentoCorrelativo.Show
End Sub

'CGMiranda-------------------------------------------------
Private Sub cmdOpcionMensaje_Click()
    frmMensajeCocina.Show
End Sub
'Fin CGMiranda---------------------------------------------


Private Sub cmdRecargarTarjeta_Click()
    FrmRecargarTarjeta.Show
    
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    Dim ImageWidth As Single
    Dim ImageHeight As Single
    
    picStretch.Visible = False
    picStretch.AutoRedraw = True
    picStretch.Height = Me.Height
    picStretch.Width = Me.Width
    
    ImageWidth = picStretch.Picture.Width * IMAGESIZE
    ImageHeight = (picStretch.Picture.Height * IMAGESIZE) '+ 3000
    
    picStretch.PaintPicture picStretch.Picture, 0, 0, Me.Width, Me.Height, 0, 0, ImageWidth, ImageHeight
    Set Me.Picture = picStretch.Image
    
    
    
End Sub



Sub ParseXmlDocument()
   Dim Doc As New MSXML2.DOMDocument
   Dim success As Boolean

   success = Doc.Load(App.path & "\2021_08_24_17_37_42_1022_2_1_192.168.3.78.xml")
   If success = False Then
      MsgBox Doc.parseError.reason
   Else
      Dim nodeList As MSXML2.IXMLDOMNodeList

      Set nodeList = Doc.selectNodes("/Transaction/Order")

      If Not nodeList Is Nothing Then
         Dim node As MSXML2.IXMLDOMNode
         Dim name As String
         Dim value As String

         For Each node In nodeList
            ' Could also do node.attributes.getNamedItem("name").text
'            name = node.selectSingleNode("@Received_Time").Text
'            value = node.selectSingleNode("@value").Text
            name = node.Attributes.Item(3).nodeValue
         Next node
      End If
   End If
End Sub
Private Function getSemana() As String

Dim oParser As MSXML2.DOMDocument

Set oParser = New MSXML2.DOMDocument

If oParser.Load("C:\out\2021_08_24_17_37_42_1022_2_1_192.168.3.78.xml") Then
Dim Nodos As MSXML2.IXMLDOMNodeList
Dim oNodo As MSXML2.IXMLDOMNode
Set Nodos = oParser.childNodes
Set oNodo = Nodos.Item(0).selectSingleNode("/Transaction/Order/Received_Time")
'MsgBox ("Valor" + oNodo.attributes["Valor1"])
getSemana = oNodo.Attributes.Item(1).nodeValue

End If
 
End Function
Private Sub MDIForm_Load()
''''getSemana
'ParseXmlDocument

'Skn$ = CStr(App.Path & "\Skyn\Clasico.cjstyles")
'    SkinFramework1.LoadSkin Skn$, vbNullString
 '   SkinFramework1.ApplyWindow Me.hwnd
    
    If lOCR Then
        mnuVenta.Caption = "&Caja Rapida"
        mnuVenta.Visible = False
        cmdOpcion2.Caption = "&Caja Rapida"
    End If
    
    
    Centrar Me
    Me.Caption = "Punto de Venta : Local " & localConectado
    'OO--------------------------------
    If lMCPV Then
        Me.Visible = False
        frmMozoUsuario.Show vbModal
        Me.Visible = True
    End If
    '----------------------------------

    ActivaInicio (False)
    If nPuerto > 0 Then
       Visor String(Int((20 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((20 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
    End If
    
    If Not lAlmacen Then
       mnuImportacion.Enabled = False
      ' mnuCorrelativoCentralPedidos.Enabled = False
    End If
    
    
    
    
    If Not lFacturacionE Then
            mnuDocumentoElectronico.Enabled = False
    Else
       If lFEOfisis Then
            mnuDocumentoElectronico.Enabled = False
       End If
    End If
    
    If pais = "002" Then
        mnuDocumentoElectronico.Enabled = False
    End If
    
    If pais = "003" And lFETusFacturasAPP Then
          mnuDocumentoElectronico.Enabled = True
    End If
    
    
    If lMCPV Then
        InicializaMCPV
    End If
    StatusBar.Panels.Item(1).Text = "Día Contable : No Iniciado"
    StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(3).Text = "Usuario : " & IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
    StatusBar.Panels.Item(4).Text = IIf(wInicio, "Turno :" & sTurno, "Turno : No Iniciado")
    
    If lSiab Then
       cmdConsultaSaldo.Visible = True
    End If
       
    If lVisaNet Then
       cmdOpcion15.Visible = True
       Open App.path & "\DLL3500.INI" For Input As #1
       Do While Not EOF(1)
          Dim xStr As String
          Input #1, sLinea
          
          Select Case Mid(sLinea, 1, 4)
                 Case "HOST"
                      xStr = Trim(Mid(sLinea, InStr(1, sLinea, "=") + 1))
                      IpPinPad = Mid(xStr, 2, Len(xStr) - 2)
                 
                 Case "PORT"
                       xStr = Trim(Mid(sLinea, InStr(1, sLinea, "=") + 1))
                       xStr = Mid(xStr, 2, Len(xStr) - 2)
                       IpPort = val(xStr)
          
                 Case "APPL"
                       xStr = Trim(Mid(sLinea, InStr(1, sLinea, "=") + 1))
                       xStr = Mid(xStr, 2, Len(xStr) - 2)
                       nTimeOut = val(xStr)
                       nTimeOut = nTimeOut * 2
          End Select
       Loop
       Close #1

       If Not ValidaIP(IpPinPad, IpPort) Then
'          MsgBox "Error de conexión", vbCritical, "VisaNet"
          lVisaNet = False
'          mnuPinPad.Visible = False
'          Exit Sub
       End If
    
       nRet = fiOpenPort(App.path & "\DLL3500.INI")
       If nRet = RET_NOK Then
          MsgBox "Error de Puerto", vbCritical, "VisaNet"
          lVisaNet = False
          mnuPinPad.Visible = False
          Exit Sub
       End If
    Else
       mnuPinPad.Visible = False
    End If
    If multiLocal = False Then
    mnuConexion.Visible = False
    End If
   
   
   'GCAA 27012022 GAAAAAAAAAAAAAAAAAAAAAAAAAAA
   
   Dim jajaj As Integer
   jajaj = Calcular("SELECT COUNT(*) as Codigo FROM LOG_SESIONES WHERE isnull(lActivo,0) = 0 and  tCaja = '" & sCaja & "' and ISNULL(PC,'')<> '" & Environ("computername") & "'", Cn)
    
   
   If jajaj > 0 Then
        MsgBox "La Caja " & sCaja & " , Ya existe activo otro Punto de venta con el mismo Codigo", vbInformation, sMensaje
        Salir
    Else
      'Cambiar el SQL
        Isql = " INSERT INTO LOG_SESIONES (tCaja, fRegistro,lActivo,tUsuario, PC) " & _
               " values ('" & sCaja & "',getdate(),0, " & _
               " '" & sUsuario & "', '" & Environ("computername") & "' ) "
        Cn.Execute Isql
   End If
   
   Call Accesos(Me, "02", sUsuario)
   On Error Resume Next
   Me.Picture = LoadPicture(App.path & "\bmps\Inforest.EMF")

   'IMAGEN PAIS
   Dim fso As Object
   Set fso = CreateObject("Scripting.FileSystemObject")
       If fso.FileExists(App.path & "\bmps\Paises\" & pais & ".jpg") Then
              ImagePais.Picture = LoadPicture(App.path & "\bmps\Paises\" & pais & ".jpg")
       End If

   Set fso = Nothing

   'frmNoticias.Show vbModal

   Call IniciarNoticia
   
   '' REQ REST-049-C-2022 - Osaka Argentina
   If lActivarIncrementoPrecios Then
        programacionPrecios ("A")
   End If

    
    
End Sub
Public Sub InicializaMCPV()
    Dim RsTurno As Recordset
    Isql = "select * from MTURNO where tUsuario ='" & sUsuario & "' and lcierre=0 order by tTurno"
    Set RsTurno = Lib.OpenRecordset(Isql, Cn)
    If RsTurno.RecordCount > 0 Then
        nTC = Calcular("SELECT nVenta as Codigo From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
        If nTC = 0 Then
            wInicio = True
            ActivaInicio (False)
            cmdOpcion1.Enabled = True
            sTurno = RsTurno!tTurno
        Else
            wInicio = True
            ActivaInicio (True)
            cmdOpcion1.Enabled = False
            sTurno = RsTurno!tTurno
        End If
    Else
         wInicio = False
         ActivaInicio (False)
         cmdOpcion1.Enabled = True
         sTurno = ""
    End If
    StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(3).Text = "Usuario : " & IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
    StatusBar.Panels.Item(4).Text = IIf(wInicio, "Turno :" & sTurno, "Turno : No Iniciado")
        
    Call Accesos(Me, "02", sUsuario)
End Sub

Public Sub Salir()
   If lMCPV Then
      If nPuerto > 0 Then
         Visor "", "", nPuerto, "N"
      End If
      
      If lVisaNet Then
         fiClosePort
      End If
      Unload Me
   Else
'      sino = MsgBox("Deseas Salir del Sistema", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje)
'      If sino = vbYes Then
         If nPuerto > 0 Then
            Visor "", "", nPuerto, "N"
         End If
           
         If lVisaNet Then
            fiClosePort
         End If
         If lHARDkey Then
            '--------Libera Licencia de la Llave------------------------
            Dim Verifica As Boolean
            Verifica = hk.FinalizarConexion(Aplicacion.PuntoVenta)
            '-----------------------------------------------------------
         End If
         
                
         'auditoria
        
         registroAccesoAuditoria "S", sUsuario
         'auditoria
         End
      'End If
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If lHARDkey Then
        '--------Libera Licencia de la Llave------------------------
        Dim Verifica As Boolean
        Verifica = hk.FinalizarConexion(Aplicacion.PuntoVenta)
        '-----------------------------------------------------------
    End If
           
           
 'Cambiar el SQL
Isql = " update  LOG_SESIONES SET lActivo = 1 WHERE lActivo = 0 and TCAJA ='" & sCaja & "' and PC = '" & Environ("computername") & "'"
Cn.Execute Isql
           
 'auditoria

 registroAccesoAuditoria "S", sUsuario
 'auditoria
 
  '  End
End Sub

Private Sub mnuAcerca_Click()
   frmAbout.Show vbModal
End Sub

Private Sub mnuAnulado_Click()
   frmAnulado.Show
End Sub

Private Sub mnuCancelado_Click()
   frmCancelado.Show
End Sub

Private Sub mnuCambiaLocal_Click()
        Dim verificaForms As Boolean
        verificaForms = VerificaFormAbiertos
        If verificaForms <> True Then
            frmServidorEnlace.Show
        Else
            MsgBox "Debe cerrar todos los formularios", vbInformation, sMensaje
        End If
End Sub

Private Sub mnuCierre_Click()
   cmdOpcion3_Click
End Sub

Private Sub mnuClienteDeuda_Click()
   frmRepClienteDeuda.Show
End Sub


Private Sub mnuClasico_Click()
Skn$ = CStr(App.path & "\Skyn\Clasico.cjstyles")
    SkinFramework1.LoadSkin Skn$, vbNullString
    SkinFramework1.ApplyWindow Me.hwnd
End Sub
Private Sub mnuEstiloOF_Click()
    Skn$ = CStr(App.path & "\Skyn\Office2007.cjstyles")
    SkinFramework1.LoadSkin Skn$, vbNullString
    SkinFramework1.ApplyWindow Me.hwnd
End Sub

Private Sub mnuLiberar_Click()
 Cn.Execute "update tparametro set tsolouno=0, lNomSoloUno='' "
 MsgBox "!Descargo de Venta Liberado", vbInformation, "Descargo de ventas"
End Sub

Private Sub mnuLuna_Click()
    Skn$ = CStr(App.path & "\Skyn\WinXP.Luna.cjstyles")
    SkinFramework1.LoadSkin Skn$, vbNullString
    SkinFramework1.ApplyWindow Me.hwnd
End Sub

Private Sub mnuRecargarTarjetaProximidad_Click()
    FrmRecargarTarjeta.Show
   
End Sub

Private Sub mnuRoyal_Click()
    Skn$ = CStr(App.path & "\Skyn\WinXP.Royale.cjstyles")
    SkinFramework1.LoadSkin Skn$, vbNullString
    SkinFramework1.ApplyWindow Me.hwnd
End Sub

Private Sub mnuTarjetaProximidad_Click()
FrmTarjetaAproximidad.Show
End Sub

Private Sub mnuVista_Click()
    Skn$ = CStr(App.path & "\Skyn\Vista.cjstyles")
    SkinFramework1.LoadSkin Skn$, vbNullString
    SkinFramework1.ApplyWindow Me.hwnd
End Sub
Private Sub mnuCorrelativo_Click()
   frmFactura.Show vbModal
End Sub

Private Sub mnuCorrelativoCentralPedidos_Click()
  frmCentralPedidos.Show
End Sub

Private Sub mnuCorrelativoDocumento_Click()
   frmDocumentoCorrelativo.Show
End Sub

Private Sub mnuCorrelativoPedido_Click()
   frmPedidoCorrelativo.Show
End Sub

Private Sub mnuCtaCte_Click()
    If nTC = 0 Then MsgBox "Es necesario aperturar el Turno!": Exit Sub
   frmCtaCte.Show
End Sub

Private Sub mnuCuentaCorriente_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuCuentaCobrar_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuDocumentoElectronico_Click()
   frmDocumentoElectronicoCorrelativo.Show
End Sub

 

Private Sub mnuGuiaTransporte_Click()
   frmGuiaTransporte.Show
End Sub

Private Sub mnuImportacion_Click()
   frmImportacionRequerimientos.Show
End Sub

Private Sub mnuInicio_Click()
   cmdOpcion1_Click
End Sub

Private Sub mnuInsumoCritico_Click()
    frmInsumo.Show
End Sub

Private Sub mnuMesa_Click()
   Screen.MousePointer = vbHourglass
   sTipo = "V"
   frmMesaConsulta.Show vbModal
End Sub

'Private Sub mnuPuntoVenta_Click()
'   cmdOpcion2_Click
'End Sub

Private Sub mnuRanking_Click()
   frmRepRanking.Show
End Sub

Private Sub mnuNotaCredito_Click()
   frmNotaCredito.Show
End Sub

Private Sub mnuPinPad_Click()
   Dim nRet As Integer
   Dim sOperacion As String
   Dim sRetorno As String * 512
   Dim lLoop As Boolean
   Dim nContador As Integer
      
   sOperacion = OP_NO_FINANCIERA & "A" & "0000000000.00" & Chr$(FS) & _
                                   "B" & "000000000000" & Chr$(FS) & _
                                   "C" & "0" & Chr$(FS) & _
                                   "D" & sEmpresa & Chr$(FS) & _
                                   "E" & sCaja

   nRet = fiStartOperation(sOperacion, 2, sRetorno)
      
   If nRet = RET_OK Or nRet = RET_RUNNING Then
      
      If Not Imprimir(sPreCuenta) Then
         Exit Sub
      End If
      Printer.FontName = sFont
      Printer.FontBold = False
      lLoop = True
      nContador = 0

      Do
        sRetorno = ""
        nRet = fiGetStatus(sRetorno, 512)
        MensajePinPad sRetorno
        
        Mensaje "PinPad Listo. Esperando...", "PinPad", 500
        nContador = nContador + 1
        If nContador >= nTimeOut Then
           If MsgBox("Tiempo de espera agotado, deseas mas tiempo?", vbExclamation + vbOKCancel, "VisaNet") = vbOK Then
              lLoop = True
              nContador = nTimeOut / 2
           Else
              lLoop = False
           End If
        End If
        
        If nRet <> "0" Then
           nContador = 0
        End If
      Loop While (Mid$(sRetorno, 5, 2) <> "C1") And lLoop
      Printer.EndDoc
   Else
      MsgBox "Error de conectividad", vbCritical, sMensaje
   End If
              
End Sub

Private Sub mnuRecibo_Click()
   cmdOpcion5_Click
End Sub

Private Sub mnuRegistroVenta_Click()
   frmRepRegistroVenta.Show
End Sub

Private Sub mnuReporteVenta_Click()
   frmRepPaloteo.Show
End Sub

Private Sub mnuReciboIngreso_Click()
   cmdOpcion6_Click
End Sub

Private Sub mnuReserva_Click()
   frmReserva.Show
End Sub

Private Sub mnuSalir_Click()
   Salir
End Sub

Private Sub mnuVenta_Click()
    Screen.MousePointer = vbHourglass
    frmVenta.Show vbModal
End Sub

Private Sub mnuVentaMensual_Click()
   frmRepVentaMensual.Show
End Sub

'UN EXE VARIAS BD
Public Sub reinicia()
    Unload Me
    mdiPuntoVenta.Show
End Sub

Private Sub Timer_Timer()
    If multiLocal = True Then
        If ultimoConectado = False Then
             frmServidorEnlace.Show
            Timer.Interval = 0
        End If
    End If
    

    
End Sub
 
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
End Function
Public Function obtieneNoticia() As String
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("SP_NOTICIAS_INFOREST", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, 1
                  oComando.CreateParameter "@usuario", adVarChar, adParamInput, 50, sUsuario
                  oComando.CreateParameter "@resultado", adVarChar, adParamOutput, 200, ""
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneNoticia = oComando.GetParameterValue("@resultado")
End Function

Private Sub TimerActualizador_Timer()
On Error GoTo fin:
'    Dim CnSeg As New Connection
'
'    Set CnSeg = New Connection
'    CnSeg.Provider = "SQLOLEDB"
'    CnSeg.CursorLocation = adUseServer
'    CnSeg.ConnectionString = "User ID=" & sUserName & _
'                          ";password=" & sPassword & _
'                          ";Data Source=" & sRuta & _
'                          ";DataTypeCompatibility=80" & _
'                          ";Initial Catalog=INFSEGURIDAD"
'    CnSeg.CommandTimeout = 250
'    CnSeg.Open
If ActivoActualizador Then
    If CnSeg.State = 0 Then
        CnSeg.Open
    End If
    
    Dim Valor, ValorMotor, ValorMotorActual As String
    'Valor = Calcular("select top 1 VersionInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)
    Open App.path & "\version.txt" For Input As #1   ' Abre el archivo para recibir los datos.
    Line Input #1, Valor
    Close #1
        
    ValorMotor = Calcular("select top 1 VersionFeInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)
    ValorMotorActual = Calcular("select top 1 VersionActualFEInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)

    Dim sVersionExe As String
    sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
    
    If (ValorMotor <> "") Then
        If (ValorMotor <> ValorMotorActual) Then
            If (RsParametro!lFEBiz) Then
                FrmActualizacion.tipo = "FE"
                FrmActualizacion.lblMensaje.Caption = "Nueva Actualización disponible de Facturacion Electronica" & vbNewLine & " FE Versión: " & ValorMotor
                FrmActualizacion.Show vbModal
                Exit Sub
            End If
        End If
    End If
    
    If (Valor <> "") Then
        If (Valor <> sVersionExe) Then
            FrmActualizacion.tipo = "INFOREST"
            FrmActualizacion.lblMensaje.Caption = "Nueva Actualización disponible" & vbNewLine & " InfoRest Versión: " & Valor & vbNewLine & "Se recomienda actualizar el sistema, ya que puede presentar problemas en algunos procesos.!!!"
            FrmActualizacion.Show vbModal
            Exit Sub
        End If
    End If
End If
Exit Sub
fin:
MsgBox "Error:" & Error
End Sub

Private Sub TimerBizlink_Timer()
    If Contador > 5 Then
        If (lFEBiz) And (lFEBizOffLine) Then
            If INSERTA_FE_INFOREST("", 6, DateTime.Now) = False Then
            End If
        End If
        If pais = "000" Then
            If lFacturacionE = True And lFEBiz = False And lFEOfisis = False And lFEpape = False And lFECarbajal = False And lFESpring = False And lFEGesa = False Then
                INSERTAFE "UTIL", "", 4, ""
            End If
        End If
        Contador = 0
    Else
        Contador = Contador + 1
    End If
End Sub

Private Sub Timerwebapp_Timer()

    Dim statusHora As Boolean
    Dim RsAlertaVenta As Recordset
    Dim Isql As String
'    validaciòn de rango horario
    statusHora = Calcular("SELECT CASE WHEN CONVERT(time, GETDATE()) BETWEEN '23:50:00' AND '23:59:59' THEN 'True' Else 'False'END AS Codigo", Cn)
'    consulta configuraciòn de alerta venta
    Isql = "select lAlertaVenta from TPARAMETRO"
    Set RsAlertaVenta = Lib.OpenRecordset(Isql, Cn)
    If (RsAlertaVenta!lAlertaVenta = True) Then
        If (statusHora) Then
            If (wInicioActivaApertura = True) Then
                If MsgBox("Es necesario realizar descarga de Ventas") Then
                    wInicioActivaApertura = False
                End If
             End If
        End If
    End If

'     If frmVenta.ActivaOpcion = True Then
'        MsgBox ("abierto")
'    Else
'        MsgBox ("cerrado")
'    End If
 
'    If Screen.ActiveForm.Name = frmVenta.Name Then
'        frmVenta.SetFocus
'        Exit Sub
'    Else
      
'    If WEBAP Then
'        ' GCAA 29102020
'
'        Dim vcanal1 As String
'        Dim vcanal2 As String
'        Dim vcanal3 As String
'        Dim vcanal4 As String
'        Dim vcanal5 As String
'        Dim val As Integer
'
'        Isql = "select tCanalventa1, tCanalventa2, tCanalventa3, tCanalventa4, tCanalventa5 from TCAJACANALVENTA Where  tCaja = '" + sCaja + "' "
'        Set RsListaAsignada = Lib.OpenRecordset(Isql, Cn)
'
'        val = 0
'         If Not (RsListaAsignada.EOF Or RsListaAsignada.BOF) Then
'            If RsListaAsignada!tCanalventa1 = True Then
'                vcanal1 = "01"
'                val = val + 1
'            Else
'                vcanal1 = ""
'                val = val + 0
'            End If
'
'            If RsListaAsignada!tCanalventa2 = True Then
'                vcanal2 = "02"
'                val = val + 1
'            Else
'                vcanal2 = ""
'                val = val + 0
'            End If
'
'            If RsListaAsignada!tCanalventa3 = True Then
'                vcanal3 = "03"
'                val = val + 1
'            Else
'                vcanal3 = ""
'                val = val + 0
'            End If
'
'            If RsListaAsignada!tCanalventa4 = True Then
'                vcanal4 = "04"
'                val = val + 1
'            Else
'                vcanal4 = ""
'                val = val + 0
'            End If
'
'            If RsListaAsignada!tCanalventa5 = True Then
'                vcanal5 = "05"
'                val = val + 1
'            Else
'                vcanal5 = ""
'                val = val + 0
'            End If
'        End If
'
''        If val = 0 Then
''            Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion = 2 and testadopedido<>'03'"
''        Else
''            Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion = 2 and testadopedido<>'03' and  (tTipoPedido = '" + vcanal1 + "' or tTipoPedido = '" + vcanal2 + "' or tTipoPedido = '" + vcanal3 + "'  or tTipoPedido = '" + vcanal4 + "'  or tTipoPedido = '" + vcanal5 + "')"
''        End If
''
''
''        Set RsPedido = Lib.OpenRecordset(Isql, Cn)
''        If RsPedido.RecordCount > 0 Then
''          '  Me.bocina.url = App.Path & "\bmps\rappi.mp3"
''        Else
''             'Me.bocina.Controls.Stop
''        End If
'    End If
    
    'integracion KDS GCAA - N018-2021 - INFOMATICA - 17082021
    If (lIntMobileKDS) Then
            Isql = "SELECT  top 1 *, MPEDIDO.tTiporecepcion, MMOBILE_KDS.accion as xaccion ,MMOBILE_KDS.item as xItemEli from vPedidoCabecera inner join MPEDIDO on vPedidoCabecera.Codigo  = MPEDIDO.tCodigoPedido inner join MMOBILE_KDS ON MPEDIDO.tCodigoPedido = MMOBILE_KDS.tCodigoPedido " & _
             "where vPedidoCabecera.tEstadoPedido IN ('01','03') order by vPedidoCabecera.fFECHA"
        
        Set RsCabecerax = Lib.OpenRecordset(Isql, Cn)
        
        '1 - MOBILE
        '0 - INFOREST
        If Not (RsCabecerax.EOF Or RsCabecerax.BOF) Then
            If RsCabecerax!xaccion = "ADD" And RsCabecerax!tEstadoPedido = "01" Then 'AGREGAR PEDIDO
                Call KDS_AnadirNuevaOrden(RsCabecerax, 1)
            ElseIf RsCabecerax!xaccion = "DEL" And RsCabecerax!tEstadoPedido = "01" Then 'ELIMINAR ITEM
                Call KDS_EliminarProducto(RsCabecerax, RsCabecerax!xItemEli, 1)
            ElseIf RsCabecerax!xaccion = "PED" And RsCabecerax!tEstadoPedido = "03" Then 'ELIMINAR PEDIDO
                Call KDS_EliminarOrden(RsCabecerax) 'KDS
                Cn.Execute "delete MMOBILE_KDS  where tCodigoPedido='" & RsCabecerax!codigo & "' and accion = '" & RsCabecerax!xaccion & "'"
            End If
        End If
    End If
    'FINNNNN
    
    
    
    MMControl1.Command = "stop"
'    MMControl1.FileName = App.Path & "\bmps\patio.wav"
'    MMControl1.Command = "Open"
'    MMControl1.Command = "Play"
    If lIntPatio Then
        ' Me.bocina.url = App.Path & "\bmps\patio.mp3"
        Isql = "select  tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=3 and testadopedido<>'03' and isnull(nreimpresion,0)=0 "
        Set RsPedido = Lib.OpenRecordset(Isql, Cn)
        If RsPedido.RecordCount > 0 Then
'            Me.bocina.url = App.Path & "\bmps\patio.mp3"
            MMControl1.FileName = App.path & "\bmps\patio.wav"
            MMControl1.Command = "open"
            MMControl1.Command = "play"
            Call comandaAuto
        Else
'            Me.bocina.Controls.Stop
            MMControl1.Command = "stop"
        End If
    End If
'  End If
End Sub
Sub comandaAuto()
 'Obteniendo pedidos por imprimir automatico
    ' 20012021
    Dim xSql As String
    Dim i As Integer
    Dim xPedido As String
    xPedido = ""
    Dim xUsuarioAutorizaenvio, xPerfilUsusario As String
    xUsuarioAutorizaenvio = "MASTER"
    xPerfilUsusario = "01"
     xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                         "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                         "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                         "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
                         
    Isql = "select top 1 tCodigoPedido , tSalon from MPEDIDO WHERE tTiporecepcion=3 and testadopedido<>'03'  and isnull(nReimpresion,0) = 0 order by fregistro desc "
    Set RsPedidosCarta = Lib.OpenRecordset(Isql, Cn)
     

     
    If Not RsPedidosCarta.EOF Then
    
        xPedido = RsPedidosCarta!tCodigoPedido
        Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
         
        'Producto Propiedad
        Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador " & _
               "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo and dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto and dbo.TPRODUCTOPROPIEDAD.tenlace= t1.tenlace " & _
               "where tCodigoPedido = '" & xPedido & "'"
        Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
           
        'Combo Propiedad
        Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador " & _
               "FROM dbo.TCOMBOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace " & _
               "where tCodigoPedido = '" & xPedido & "'"
         Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
        
        Isql = "select * from vPedido " & _
           "Where Codigo = '" & xPedido & "' And lImprimeArea = 1 and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 )) " & _
           "ORDER BY nOrden, tItem, nOrdenCombo, tetiqueta,combo"
           '"ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                           
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                
        If Not RsImpresion.EOF Then
           RsArea.MoveFirst
           For i = 1 To RsArea.RecordCount
               RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
               If RsArea!tIcono = "" Or RsPedidosCarta!tSalon = RsArea!tIcono Or ((IsNull(RsPedidosCarta!tSalon) Or RsPedidosCarta!tSalon = "") And RsArea!nValor = 1) Then
                  If RsImpresion.RecordCount <> 0 Then
                     RsImpresion.MoveFirst
                     ImprimePedidoAuto RsImpresion, "N", RsArea!timpresora, RsArea!Area, True, RsProductoPropiedad, RsComboPropiedad, "Inforest", xPedido
                  End If
               End If
               RsArea.MoveNext
           Next i
           
        Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1  where tCodigoPedido='" & xPedido & "'"
        Cn.Execute "Update DPEDIDO Set lImprime = 1, tUsuarioAutorizaEnvio = '" & xUsuarioAutorizaenvio & "', tPerfilUsuarioAutoriza = '" & xPerfilUsusario & "', fEnvio = getdate(), nEnvio = 1 where tCodigoPedido = '" & xPedido & "' and lImprime<>1"
        Cn.Execute "Update CPEDIDO Set lImprime = 1 where tCodigoPedido = '" & xPedido & "'"
        
        End If
        
     
    
    End If
End Sub

Public Sub IniciarNoticia()
On Error GoTo fin:
    Dim resul As String
    resul = obtieneNoticia
    If resul = "1" Then
        Exit Sub
    End If


    Dim envio As String
    'envio = sUsuario & "," & "http://www.licencias.infomatica.pe/assets/image/WhatsApp Image 2021-10-21 at 15.24.13.jpeg"
    envio = sUsuario & "," & resul & "," & sMDB & "," & sRuta

    'INICIO DE ACTUALIZADOR
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(App.path & "\Noticia\InfoNoticias.exe") Then
        'Kill App.Path & "\images\fact.bmp"
        ShellExecute 0, "Open", App.path & "\Noticia\InfoNoticias.exe", envio, "", 1
'    Else
'        MsgBox "La ruta de la Noticia no se ha encontrado, favor de informar al administrador de sistemas!!!: "
    End If
    
   ''''  Shell App.path & "\Noticia\InfoNoticias.exe", vbNormalFocus
Exit Sub
fin:
MsgBox "Noticia: " & Error
End Sub
