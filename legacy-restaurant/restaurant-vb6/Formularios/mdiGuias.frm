VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiGuias 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo de Guías de Remisión"
   ClientHeight    =   7065
   ClientLeft      =   -2700
   ClientTop       =   1155
   ClientWidth     =   11880
   Icon            =   "mdiGuias.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList 
      Left            =   90
      Top             =   6150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGuias.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGuias.frx":068E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGuias.frx":1E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGuias.frx":26FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGuias.frx":2A18
            Key             =   ""
         EndProperty
      EndProperty
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
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   8502
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            Text            =   "Now"
            TextSave        =   "13/02/2006"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "05:16 p.m."
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Configuración de Impresión"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guías de Remisión"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reporte de Guías"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   720
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuTransferencia 
      Caption         =   "&Administración"
      Begin VB.Menu mnuCliente 
         Caption         =   "Destinatario - Transportistas"
      End
      Begin VB.Menu mnuMotivo 
         Caption         =   "Motivo de Traslado"
      End
      Begin VB.Menu mnuChofer 
         Caption         =   "Choferes"
      End
      Begin VB.Menu mnuVehiculo 
         Caption         =   "Vehículos"
      End
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "&Procesos"
      Begin VB.Menu mnuGuias 
         Caption         =   "Correlativo de Guías de Remisión"
      End
      Begin VB.Menu mnuMarcar 
         Caption         =   "Marcar Factruras"
      End
   End
   Begin VB.Menu mnuReporte 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuGuia 
         Caption         =   "Guias de Remisión"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu mnuRegistroVenta 
         Caption         =   "Registro de Ventas Consolidado"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "mdiGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\bmps\Inforest.EMF")
    StatusBar.Panels.Item(1).Text = "Caja : " & sCaja
    StatusBar.Panels.Item(2).Text = "Usuario : " & sUsuario
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

Private Sub mnuChofer_Click()
   frmChofer.Show
End Sub

Private Sub mnuCliente_Click()
   frmClienteFactura.Show
End Sub
'erwin Lunes 10-10-05
Private Sub mnuGuia_Click()
frmRepGuias.Show vbModal
End Sub
'end erwin

Private Sub mnuGuias_Click()
   frmGuias.Show
End Sub

Private Sub mnuMarcar_Click()
   frmMarcar.Show
End Sub

Private Sub mnuMotivo_Click()
   frmMotivo.Show
End Sub

Private Sub mnuSalir_Click()
   Salir
End Sub

Private Sub mnuVehiculo_Click()
   frmVehiculo.Show
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   'erwin Mar 11-10-05
   Select Case Button.Index
          Case Is = 1
               On Error Resume Next
               dlgPrinter.ShowPrinter
          Case Is = 3
               frmGuias.Show
               'no se puede llamar directamente al detalle de guias porq alli se necestia q el objeto rsCabecera de frmGuias este creado
          Case Is = 4
               frmRepGuias.Show
          Case Is = 6
               Salir
    'end erwin
   End Select
End Sub
