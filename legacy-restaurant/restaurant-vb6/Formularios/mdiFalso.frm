VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mdiFalso 
   Caption         =   "Punto de Venta"
   ClientHeight    =   7965
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11310
   Icon            =   "mdiFalso.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox xPicture 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11250
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      Begin VB.CommandButton cmdOpcion7 
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
         Left            =   7740
         Picture         =   "mdiFalso.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion5 
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
         Left            =   5160
         Picture         =   "mdiFalso.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion1 
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
         Picture         =   "mdiFalso.frx":0E0E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion2 
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
         Left            =   1290
         Picture         =   "mdiFalso.frx":0F10
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion4 
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
         Left            =   3870
         Picture         =   "mdiFalso.frx":1012
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion16 
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
         Left            =   9030
         Picture         =   "mdiFalso.frx":110C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion11 
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
         Left            =   3870
         Picture         =   "mdiFalso.frx":11FE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion6 
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
         Left            =   6450
         Picture         =   "mdiFalso.frx":1788
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion14 
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
         Left            =   7740
         Picture         =   "mdiFalso.frx":188A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion8 
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
         Picture         =   "mdiFalso.frx":1E14
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   9030
         Picture         =   "mdiFalso.frx":1F5E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   675
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion9 
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
         Left            =   1290
         Picture         =   "mdiFalso.frx":22E8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion13 
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
         Left            =   6450
         Picture         =   "mdiFalso.frx":2872
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion12 
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
         Left            =   5160
         Picture         =   "mdiFalso.frx":2DFC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion3 
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
         Left            =   2580
         Picture         =   "mdiFalso.frx":3386
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion10 
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
         Left            =   2580
         Picture         =   "mdiFalso.frx":3488
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   675
         Width           =   1245
      End
      Begin VB.CommandButton cmdAmpliar 
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
         Left            =   10305
         Picture         =   "mdiFalso.frx":3A12
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   795
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   18
      Top             =   7665
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   3581
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
            TextSave        =   "18/05/2010"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "04:14 p.m."
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
         Caption         =   "&Apertura de Turnos"
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
      Begin VB.Menu Linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorrelativo 
         Caption         =   "Correlativo de Documentos"
      End
      Begin VB.Menu Linea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMesa 
         Caption         =   "&Mesas"
      End
   End
   Begin VB.Menu mnuCuentas 
      Caption         =   "&Correlativo"
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
         Caption         =   "Recibo de Egreso"
      End
      Begin VB.Menu mnuReciboIngreso 
         Caption         =   "Recibo de Ingreso"
      End
      Begin VB.Menu mnuNotaCredito 
         Caption         =   "Notas de Credito"
      End
      Begin VB.Menu mnuReserva 
         Caption         =   "Reservas"
      End
      Begin VB.Menu Linea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCuentaCobrar 
         Caption         =   "Cuentas por Cobrar"
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
Attribute VB_Name = "mdiFalso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAmpliar_Click()
    If cmdAmpliar.Caption = "Esconder" Then
       xPicture.Height = 700
       cmdAmpliar.Caption = "Ampliar"
    Else
       xPicture.Height = 1380
       cmdAmpliar.Caption = "Esconder"
    End If
End Sub

Private Sub cmdOpcion1_Click()
    frmInicio.Show vbModal
    StatusBar.Panels.Item(3).Text = IIf(wInicio, "Turno : " & sTurno, "Turno : No Iniciado")
    cmdOpcion1.Enabled = IIf(wInicio, False, True)
    mnuInicio.Enabled = IIf(wInicio, False, True)
    If wInicio Then
       ActivaInicioFalso (True)
    End If
End Sub

Private Sub cmdOpcion10_Click()
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
    Salir
End Sub

Private Sub cmdOpcion2_Click()
   Screen.MousePointer = vbHourglass
   If lCajaRapida Or lMultiCajero Then
      frmCajaRapida.Show vbModal
      If Not wEnter Then
         frmVenta.Show vbModal
      End If
   Else
      frmVenta.Show vbModal
   End If
End Sub

Private Sub cmdOpcion3_Click()
    frmLiquidacionDetalle.Show vbModal
    If wInicio = False Then
    ActivaInicioFalso (False)
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

Private Sub Form_Load()
    ActivaInicioFalso (False)
    If nPuerto > 0 Then
       Visor String(Int((20 - Len(tMensaje1)) / 2), " ") & tMensaje1, String(Int((20 - Len(tMensaje2)) / 2), " ") & tMensaje2, nPuerto, "N"
    End If
    
    If lMCPV Then
       Dim RsTurno As Recordset
       Isql = "select * from MTURNO where tUsuario ='" & sUsuario & "' and lcierre=0 order by tTurno"
       Set RsTurno = Lib.OpenRecordset(Isql, Cn)
       If RsTurno.RecordCount > 0 Then
          wInicio = True
          ActivaInicioFalso (True)
          cmdOpcion1.Enabled = False
          sTurno = RsTurno!tTurno
          nTC = Calcular("SELECT nVenta as Codigo From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
       End If
    End If
    
    StatusBar.Panels.Item(1).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(2).Text = "Usuario : " & IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
    StatusBar.Panels.Item(3).Text = IIf(wInicio, "Turno :" & sTurno, "Turno : No Iniciado")
       
    If lVisaNet Then
       cmdOpcion15.Visible = True
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
                       IpPort = Val(xStr)
          
                 Case "APPL"
                       xStr = Trim(Mid(sLinea, InStr(1, sLinea, "=") + 1))
                       xStr = Mid(xStr, 2, Len(xStr) - 2)
                       nTimeOut = Val(xStr)
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
    
    Call Accesos(Me, "02", sUsuario)
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\bmps\Inforest.EMF")
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
   cmdOpcion3_Click
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

Private Sub mnuCuentaCorriente_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuCuentaCobrar_Click()
   frmCuentaCobrar.Show
End Sub

Private Sub mnuInicio_Click()
   cmdOpcion1_Click
End Sub

Private Sub mnuMesa_Click()
   Screen.MousePointer = vbHourglass
   sTipo = "V"
   frmMesaConsulta.Show vbModal
End Sub

Private Sub mnuPuntoVenta_Click()
   cmdOpcion2_Click
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

Public Sub ActivaInicioFalso(Activa As Boolean)
    mdiFalso.cmdOpcion2.Enabled = Activa
    mdiFalso.cmdOpcion3.Enabled = Activa
    mdiFalso.mnuVenta.Enabled = Activa
    mdiFalso.mnuCierre.Enabled = Activa
End Sub
