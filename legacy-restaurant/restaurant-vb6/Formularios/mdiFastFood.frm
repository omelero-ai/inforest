VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiFastFood 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Integral para Restaurantes"
   ClientHeight    =   7065
   ClientLeft      =   -2700
   ClientTop       =   1155
   ClientWidth     =   11880
   Icon            =   "mdiFastFood.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Recibos de Ingresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   8
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Ctas x Cobrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   3870
         Picture         =   "mdiFastFood.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   7800
         Picture         =   "mdiFastFood.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Recibos de Egresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   5190
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Cierre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   2580
         Picture         =   "mdiFastFood.frx":057E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Fast Food"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   1290
         Picture         =   "mdiFastFood.frx":0680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1290
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Apertura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   0
         Picture         =   "mdiFastFood.frx":0782
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6765
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   4586
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "20/05/2005"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "12:32 p.m."
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
      Begin VB.Menu mnuCierre 
         Caption         =   "Ci&erre de Turno"
      End
      Begin VB.Menu linea8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorrelativo 
         Caption         =   "Correlativo de Documentos"
      End
   End
   Begin VB.Menu mnuCuentas 
      Caption         =   "&Documentos"
      Begin VB.Menu mnuCorrelativoPedido 
         Caption         =   "Correlativo de Pedidos"
      End
      Begin VB.Menu mnuCorrelativoDocumento 
         Caption         =   "Correlativo de Documentos"
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
      Begin VB.Menu linea6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCuentaCobrar 
         Caption         =   "Cuenta por Cobrar"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mdiFastFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Inicio
               frmInicio.Show vbModal
               cmdOpcion(0).Enabled = IIf(wInicio, False, True)
               mnuInicio.Enabled = IIf(wInicio, False, True)
               
          Case Is = 1 ' Pto. Venta
               Screen.MousePointer = vbHourglass
               frmFastFood.Show vbModal
               
          Case Is = 2 ' Cierre
               If sGrupoUsuario = "01" Then
                  frmCierre.Show vbModal
               Else
                  frmCiego.Show vbModal
               End If
               
          Case Is = 4 ' CtasxCtes
               frmCuentaCobrar.Show
               
          Case Is = 5 ' Recibos
               frmReciboEgreso.Show
               
          Case Is = 6 ' Salir
               Salir
                              
          Case Is = 8 ' Recibos Ingresos
               frmReciboIngreso.Show
   End Select
End Sub

Private Sub MDIForm_Load()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\bmps\Inforest.EMF")
    StatusBar.Panels.Item(1).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(2).Text = "Usuario : " & sUsuario
    StatusBar.Panels.Item(3).Text = IIf(wInicio, "Turno :" & sTurno, "Turno : No Iniciado")
    mnuReporte.Enabled = IIf(sGrupoUsuario = "01", True, False)
    ActivaInicio (False)
End Sub

Public Sub Salir()
   sino = MsgBox("Deseas Salir del Sistema", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje)
   If sino = vbYes Then
      End
   End If
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

Private Sub mnuCierre_Click()
   cmdOpcion_Click (2)
End Sub

Private Sub mnuClienteDeuda_Click()
   frmRepClienteDeuda.Show
End Sub

Private Sub mnuCorrelativo_Click()
   frmFactura.Show vbModal
End Sub

Private Sub mnuCorrelativoDocumento_Click()
   frmDocumentoCorrelativo.Show
End Sub

Private Sub mnuCorrelativoPedido_Click()
   frmPedidoCorrelativo.Show
End Sub

Private Sub mnuCtaCte_Click()
   frmCtaCte.Show
End Sub

Private Sub mnuCuenta_Click()
   cmdOpcion_Click (4)
End Sub

Private Sub mnuCuentaCorriente_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuCuentaCobrar_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuInicio_Click()
   cmdOpcion_Click (0)
End Sub

Private Sub mnuLiquidacion_Click()
    frmRepLiquidacion.Show
End Sub

Private Sub mnuPuntoVenta_Click()
   cmdOpcion_Click (1)
End Sub

Private Sub mnuPropina_Click()
   frmRepPropina.Show
End Sub

Private Sub mnuRanking_Click()
   frmRepRanking.Show
End Sub

Private Sub mnuNotaCredito_Click()
   frmNotaCredito.Show
End Sub

Private Sub mnuRecibo_Click()
   cmdOpcion_Click (5)
End Sub

Private Sub mnuRegistroVenta_Click()
   frmRepRegistroVenta.Show
End Sub

Private Sub mnuReporteVenta_Click()
   frmRepPaloteo.Show
End Sub

Private Sub mnuReciboIngreso_Click()
   cmdOpcion_Click (8)
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
