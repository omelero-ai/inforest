VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiDespachador 
   BackColor       =   &H8000000C&
   Caption         =   "Módulo de Despachador"
   ClientHeight    =   7065
   ClientLeft      =   -2700
   ClientTop       =   1155
   ClientWidth     =   11880
   Icon            =   "mdiDespachador.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStretch 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   6915
      Left            =   0
      Picture         =   "mdiDespachador.frx":030A
      ScaleHeight     =   6915
      ScaleWidth      =   11880
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
   End
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiDespachador.frx":A47F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiDespachador.frx":A48FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiDespachador.frx":A4C16
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
            Object.Width           =   8476
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
            TextSave        =   "15/10/2025"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   2205
            TextSave        =   "01:01 p.m."
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
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   720
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "&Procesos"
      Begin VB.Menu mnuDespachador 
         Caption         =   "Despachador"
      End
      Begin VB.Menu mnuTiempoDelivery 
         Caption         =   "Tiempo Delivery"
      End
      Begin VB.Menu PRappi 
         Caption         =   "Pedido Rappi"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTarifaMotorizado 
         Caption         =   "Tarifa para Planilla de Motorizados"
      End
      Begin VB.Menu mnuMaximoMotorizado 
         Caption         =   "Valor Máximo Asignación a Motorizado"
      End
   End
   Begin VB.Menu mnuAsignacionMotorizado 
      Caption         =   "&Reasignacion de Motorizado"
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mdiDespachador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const IMAGESIZE = 0.566893424036281

Private Sub MDIForm_Load()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\bmps\Inforest.EMF")
    
    '0083-2013 CESAR
    If tTarifaActualMotorizado = "Tarifa Especial" Then  'Tarifa Especial

    Else
           Dim dia As String
           dia = WeekdayName(Weekday(Date))
           If dia = "domingo" Or dia = "Domingo" Or dia = "Sunday" Then  'Tarifa Domingo
              tTarifaActualMotorizado = "Tarifa Dom"
              Cn.Execute "Update tparametro set tTarifaActualMotorizado ='" & tTarifaActualMotorizado & "'"
           Else                     'Tarifa Lunes - Sabado
              tTarifaActualMotorizado = "Tarifa Lun - Sab"
              Cn.Execute "Update tparametro set tTarifaActualMotorizado ='" & tTarifaActualMotorizado & "'"
           End If

    End If
    
    StatusBar.Panels.Item(1).Text = "Caja : " & sCaja & "  /  Tarifa Actual Motorizado : " & tTarifaActualMotorizado
    StatusBar.Panels.Item(2).Text = "Usuario : " & sUsuario
    
    If lOrdenesRappi = False Then
        PRappi.Visible = False
    Else
        PRappi.Visible = True
    End If

    frmDespachador.Show
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
    ImageHeight = (picStretch.Picture.Height * IMAGESIZE) + 3000
    picStretch.PaintPicture picStretch.Picture, 0, 0, Me.Width, Me.Height, 0, 0, ImageWidth, ImageHeight
    Set Me.Picture = picStretch.Image
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

Private Sub mnuAsignacionMotorizado_Click()
   frmReasignacionMotorizado.Show
End Sub

Private Sub mnuDespachador_Click()
   frmDespachador.Show
End Sub

Private Sub mnuMaximoMotorizado_Click()
   frmAsignacionMotorizado.Show
End Sub

Private Sub mnuSalir_Click()
   Salir
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
          Case Is = 1
               frmDespachador.Show
          Case Is = 2
          Case Is = 4
               Salir
   End Select
End Sub

Private Sub mnuTarifaMotorizado_Click()
frmTarifaMotorizado.Show vbModal
End Sub

Private Sub mnuTiempoDelivery_Click()
    sTipo = "Prepintado"
    sCodigo = Calcular("select nTiempoMinutoCD as Codigo FROM TPARAMETRO", Cn)
    frmNumPad.Show vbModal
    If wEnter And Val(sDescrip) > 0 Then
       Cn.Execute "update TPARAMETRO set nTiempoMinutoCD = " & Val(sDescrip)
       MsgBox "Parámetro Actualizado", vbExclamation, sMensaje
    End If
End Sub

Private Sub PRappi_Click()
    frmOrdenesConsola.Show
End Sub
