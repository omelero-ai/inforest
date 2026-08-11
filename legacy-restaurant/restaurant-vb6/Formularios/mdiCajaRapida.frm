VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiCajaRapida 
   BackColor       =   &H8000000C&
   Caption         =   "Caja Rápida"
   ClientHeight    =   7065
   ClientLeft      =   -2700
   ClientTop       =   1155
   ClientWidth     =   11880
   Icon            =   "mdiCajaRapida.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11385
      Top             =   6030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   2
         Left            =   2490
         Picture         =   "mdiCajaRapida.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   8
         Left            =   4980
         Picture         =   "mdiCajaRapida.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   13
         Left            =   7470
         Picture         =   "mdiCajaRapida.frx":0996
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   7
         Left            =   3735
         Picture         =   "mdiCajaRapida.frx":0D20
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   14
         Left            =   8865
         Picture         =   "mdiCajaRapida.frx":0E6A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Caja Rapida"
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
         Index           =   1
         Left            =   1245
         Picture         =   "mdiCajaRapida.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   0
         Left            =   0
         Picture         =   "mdiCajaRapida.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   10
         Left            =   6225
         Picture         =   "mdiCajaRapida.frx":1160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1245
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
            Object.Width           =   4560
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
            TextSave        =   "09/03/2021"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "08:21 a.m."
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPinPad 
         Caption         =   "Activar PinPad (No Financiera)"
      End
      Begin VB.Menu mnuCierre 
         Caption         =   "Ci&erre de Turno"
      End
      Begin VB.Menu linea8 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCorrelativo 
         Caption         =   "Correlativo de Documentos"
         Visible         =   0   'False
      End
      Begin VB.Menu linea7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMesa 
         Caption         =   "&Mesas"
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
Attribute VB_Name = "mdiCajaRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Inicio
               frmInicio.Show vbModal
               StatusBar.Panels.Item(3).Text = IIf(wInicio, "Turno : " & sTurno, "Turno : No Iniciado")
               cmdOpcion(0).Enabled = IIf(wInicio, False, True)
               mnuInicio.Enabled = IIf(wInicio, False, True)
          
          Case Is = 1 ' Pto. Venta
               Screen.MousePointer = vbHourglass
               frmCajaRapida.Show vbModal
               
          Case Is = 2 ' Cierre
               frmLiquidacionDetalle.Show vbModal
               If wInicio = False Then
                  ActivaInicio (False)
               End If
               
          Case Is = 7 ' Correlativo de pedidos
               frmPedidoCorrelativo.Show
          
          Case Is = 8 ' Correlativo de documentos
               frmDocumentoCorrelativo.Show
               
          Case Is = 10 ' Recibos Egreso
               frmReciboEgreso.Show
               
          Case Is = 13 ' PinPad
                Call mnuPinPad_Click
                                   
          Case Is = 14 ' Salir
               Salir
                              
   End Select
End Sub

Private Sub MDIForm_Load()

    If nPuerto > 0 Then
       Visor String(Int((20 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((20 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
    End If
    
    StatusBar.Panels.Item(1).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(2).Text = "Usuario : " & sUsuario
    StatusBar.Panels.Item(3).Text = IIf(wInicio, "Turno :" & sTurno, "Turno : No Iniciado")
    
    If lVisaNet Then
       cmdOpcion(13).Visible = True
       Open App.Path & "\DLL3500.INI" For Input As #1
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
          MsgBox "Error de conexión", vbCritical, "VisaNet"
          lVisaNet = False
          mnuPinPad.Visible = False
          Exit Sub
       End If
    
       nRet = fiOpenPort(App.Path & "\DLL3500.INI")
       If nRet = RET_NOK Then
          MsgBox "Error de Puerto", vbCritical, "VisaNet"
          lVisaNet = False
          mnuPinPad.Visible = False
          Exit Sub
       End If
    Else
       mnuPinPad.Visible = False
    End If
    
    On Error Resume Next
    Me.Picture = LoadPicture("bmps\Inforest.EMF")
    ActivaInicio (False)
End Sub

Public Sub Salir()
   sino = MsgBox("Deseas Salir del Sistema", vbDefaultButton1 + vbYesNo + vbQuestion, sMensaje)
   If sino = vbYes Then
      If nPuerto > 0 Then
         Visor "", "", nPuerto, "N"
      End If
      
      If lVisaNet Then
         fiClosePort
      End If
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

Private Sub mnuMesa_Click()
   Screen.MousePointer = vbHourglass
   sTipo = "V"
   frmMesaConsulta.Show vbModal
End Sub

Private Sub mnuPuntoVenta_Click()
   cmdOpcion_Click (1)
End Sub

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

   sOperacion = OP_NO_FINANCIERA & "A" & sEmpresa & FS & _
                                   "B" & sCaja
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

