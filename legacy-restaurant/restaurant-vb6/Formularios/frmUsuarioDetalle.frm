VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{675D2E3C-5FBF-4F10-8213-EE25D71049AC}#2.3#0"; "sgfplibx.ocx"
Begin VB.Form frmUsuarioDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10320
   Icon            =   "frmUsuarioDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   10320
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10260
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4050
      Width           =   10320
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   5
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmUsuarioDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmUsuarioDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmUsuarioDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmUsuarioDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmUsuarioDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmUsuarioDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro"
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
            Left            =   1440
            TabIndex        =   12
            Top             =   120
            Width           =   1665
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   5880
         Picture         =   "frmUsuarioDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   4710
         Picture         =   "frmUsuarioDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   7050
         Picture         =   "frmUsuarioDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   1170
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
         Height          =   615
         Index           =   3
         Left            =   8220
         Picture         =   "frmUsuarioDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   1170
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   2640
      TabIndex        =   13
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmUsuarioDetalle.frx":3026
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboGrupoUsuario"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtBanda"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDetallado"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtResumido"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCodigo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkActivo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPassword"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Huella Digital"
      TabPicture(1)   =   "frmUsuarioDetalle.frx":3042
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "HiddenPict"
      Tab(1).Control(1)=   "fraDigitalPersona"
      Tab(1).Control(2)=   "fraSecugen"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   20
         Text            =   " "
         Top             =   2280
         Width           =   2595
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   5640
         TabIndex        =   19
         Top             =   2760
         Width           =   840
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txtResumido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   17
         Text            =   " "
         Top             =   915
         Width           =   5070
      End
      Begin VB.TextBox txtDetallado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   1305
         Width           =   5070
      End
      Begin VB.TextBox txtBanda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   100
         PasswordChar    =   "*"
         TabIndex        =   15
         Text            =   " "
         Top             =   2745
         Width           =   2595
      End
      Begin MSDataListLib.DataCombo cboGrupoUsuario 
         Height          =   315
         Left            =   1575
         TabIndex        =   21
         Top             =   1770
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
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
      Begin VB.PictureBox HiddenPict 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   -69000
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame fraDigitalPersona 
         Caption         =   "Digital Persona"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   7335
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   2600
            Left            =   0
            ScaleHeight     =   2535
            ScaleWidth      =   2535
            TabIndex        =   30
            Top             =   240
            Width           =   2600
            Begin VB.Label lblHuella 
               Alignment       =   2  'Center
               Caption         =   "Huella Dactilar Registrada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   855
               Left            =   0
               TabIndex        =   31
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.ListBox Status 
            Height          =   2400
            Left            =   2640
            TabIndex        =   29
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Huellas Dactilares Necesarias :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Samples 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   34
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Prompt 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Toque el lector de Huella Dactilar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   33
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label1 
            Caption         =   "Indicaciones :"
            Height          =   255
            Left            =   2640
            TabIndex        =   32
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.Frame fraSecugen 
         Caption         =   "SecuGen Hamster Plus (Model  HSDU03P)"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   7335
         Begin SGFPLIBXLib.FpLibXCapture FpLibXCapture1 
            Height          =   2295
            Index           =   1
            Left            =   3840
            TabIndex        =   37
            Top             =   600
            Width           =   2715
            _Version        =   131075
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   64
            Brightness      =   -30584
            Contrast        =   -30584
            Gain            =   -30584
         End
         Begin SGFPLIBXLib.FpLibXCapture FpLibXCapture1 
            Height          =   2295
            Index           =   0
            Left            =   960
            TabIndex        =   38
            Top             =   600
            Width           =   2715
            _Version        =   131075
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   64
            Brightness      =   -30584
            Contrast        =   -30584
            Gain            =   -30584
         End
         Begin SGFPLIBXLib.FpLibXVerify FpLibXVerify1 
            Left            =   6720
            Top             =   2280
            _Version        =   131075
            _ExtentX        =   1058
            _ExtentY        =   1058
            _StockProps     =   0
         End
         Begin VB.Label LabelMessage 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   2925
            Width           =   7140
         End
         Begin VB.Label LabelTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1st Huella Dactilar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   40
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label LabelTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2nd Huella Dactilar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   39
            Top             =   240
            Width           =   2715
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Grupo :"
         Height          =   195
         Index           =   3
         Left            =   945
         TabIndex        =   27
         Top             =   1815
         Width           =   525
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   195
         Index           =   4
         Left            =   690
         TabIndex        =   26
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   885
         TabIndex        =   25
         Top             =   525
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Index           =   1
         Left            =   915
         TabIndex        =   24
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   23
         Top             =   1350
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Banda Magnética :"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   2790
         Width           =   1350
      End
   End
   Begin VB.Image Image 
      Height          =   3930
      Left            =   15
      Picture         =   "frmUsuarioDetalle.frx":305E
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2595
   End
End
Attribute VB_Name = "frmUsuarioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim RsGrupoUsuario As Recordset
Dim sBanda As String

'HUELLA
Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim CreateFtrs As DPFPFeatureExtraction
Dim CreateTempl As DPFPEnrollment
Dim ConvertSample As DPFPSampleConversion
Dim Templ As DPFPTemplate
Dim blob() As Byte
Dim rHuella As Object
Dim lGrabaHuella As Boolean
Dim X As String

Dim xNuevo As String

' General error
Dim i As Integer
Const ERROR_NONE = 0
Const ERROR_CREATION_FAILED = 1
Const ERROR_FUNCTION_FAILED = 2
Const ERROR_INVALID_PARAM = 3
Const ERROR_NOT_USED = 4
Const ERROR_DLLLOAD_FAILED = 5
Const ERROR_DLLLOAD_FAILED_DRV = 6
Const ERROR_DLLLOAD_FAILED_ALGO = 7
      'Dim Cn As Connection
                   
' Device error
Const ERROR_SYSLOAD_FAILED = 51     'system file load fail
Const ERROR_INITIALIZE_FAILED = 52  'chip initialize fail
Const ERROR_LINE_DROPPED = 53       'image data drop
Const ERROR_TIME_OUT = 54           'getliveimage timeout error
Const ERROR_DEVICE_NOT_FOUND = 55   'device not found
Const ERROR_DRVLOAD_FAILED = 56     'dll file load fail
Const ERROR_WRONG_IMAGE = 57         'wrong image
Const ERROR_LACK_OF_BANDWIDTH = 58   'USB Bandwith Lack Error
Const ERROR_DEV_ALREADY_OPEN = 59     'Device Exclusive access Error
Const ERROR_GETSN_FAILED = 60         'Fail to get Device Serial Number
Const ERROR_UNSUPPORTED_DEV = 61      ' Unsupported device
   
   
' Extract&verification error
Const ERROR_FEAT_NUMBER = 101               'too small number of minutiae
Const ERROR_INVALID_TEMPLATE_TYPE = 102     'wrong template type
Const ERROR_INVALID_TEMPLATE1 = 103         'error in decoding template 1
Const ERROR_INVALID_TEMPLATE2 = 104         'error in decoding template 2
Const ERROR_EXTRACT_FAIL = 105
Const ERROR_MATCH_FAIL = 106

' Template Format
Const TEMPLATE_FORMAT_ANSI378 = 256
Const TEMPLATE_FORMAT_SG400 = 512
Const TEMPLATE_FORMAT_ISO19794 = 768


Dim g_FirstMinData() As Byte
Dim g_SecondMinData() As Byte
Dim g_MatchingMinData() As Byte
Dim g_firstStep As Boolean
Dim g_secondStep As Boolean
Dim g_vefiryResult As Boolean


Sub LlenaCombos()
    With cboGrupoUsuario
         Isql = "Select tGrupoUsuario as codigo, tDetallado as Descripcion from TGRUPOUSUARIO where tGrupoUsuario <> '00' and lActivo=1 order by Descripcion"
         Set RsGrupoUsuario = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupoUsuario
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub

Sub Asignar()
    With frmUsuario.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtPassword = Desencapsula(IIf(IsNull(!tpassword), "", !tpassword))
        txtBanda = Desencapsula(IIf(IsNull(!tBandaMagnetica), "", !tBandaMagnetica))
        cboGrupoUsuario.BoundText = IIf(IsNull(!tGrupoUsuario), "", !tGrupoUsuario)
            
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        
        'HUELLA
        If !thuella <> "" Then
            LabelMessage.Caption = "Huella Dactilar Registrada"
            lblHuella.Visible = True
        Else
            LabelMessage.Caption = ""
            lblHuella.Visible = False
        End If
        
    End With
End Sub

Private Sub chkMozo_Click()

End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmUsuario.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmUsuario.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmUsuario.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmUsuario.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmUsuario.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmUsuario.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmUsuario.RsCabecera.RecordCount = 0, 0, frmUsuario.RsCabecera.AbsolutePosition) & " de " & frmUsuario.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Dim xGrupo As String
   Dim xGrupoCombo As String
   Dim xlControlNivel As Boolean
                  
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.Value = 1
               'Cambia el Nombre del Primer Text
               txtResumido.SetFocus
               
               'HUELLA
               lblHuella.Visible = False
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               If txtPassword.Text = "" Then MsgBox "Ingrese la Clave Secreta", vbExclamation, sMensaje: txtPassword.SetFocus: Exit Sub
               If cboGrupoUsuario.Text = "" Then MsgBox "Ingrese el Grupos de Usuario", vbExclamation, sMensaje: cboGrupoUsuario.SetFocus: Exit Sub
                                                                                                                        
                X = ""
                        
               If Sw Then
               
                  'CESAR VALIDACION MODIFICACION DE GRUPO
                  xGrupo = Calcular("Select tGrupoUsuario As Codigo From  TUSUARIO Where tResumido = '" & sUsuario & "' ", Cn)
                  xlControlNivel = Calcular("Select ISNULL(lControlUsuario,0) As Codigo From  TPARAMETRO", Cn)
                  xGrupoCombo = cboGrupoUsuario.BoundText
                  
                  
                  If xGrupo <> "00" Then
                        If xlControlNivel Then
                            If xGrupo > xGrupoCombo Then
                               MsgBox "No puede Registrar Usuarios con un Nivel Superior", vbExclamation, sMensaje
                               Exit Sub
                            End If
                        End If
                  End If
                  '------------------
               
                  If Calcular("select count(tResumido) as codigo from tusuario where tResumido='" & txtResumido.Text & "'", Cn) > 0 Then
                     MsgBox "Usuario Repetido, intente de nuevo", vbExclamation, sMensaje
                     txtResumido.SetFocus
                     Exit Sub
                  End If

                  If Calcular("select count(tPassword) as codigo from tusuario where tpassword='" & Encapsula(txtPassword.Text) & "'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                   If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                             
                                              
                  
                  sBanda = Extrae(txtBanda.Text)
                  If txtBanda.Text <> "" And Calcular("select count(tBandaMagnetica) as Codigo from tUsuario where tBandaMagnetica='" & Encapsula(sBanda) & "'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
               
                  'Asignar El Campo de Codificación
                   nCorrela = Calcular("select max(tCodigoUsuario) as Codigo from TUSUARIO", Cn)
                   If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "00001"
                   Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
                   End If
                   Sw = False
                   sBanda = Extrae(txtBanda.Text)
                   
                                    
                    sPasa = txtCodigo.Text
                                       
                 'Inserta Movimiento auditoria
                   lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TUSUARIO", "USUARIO", "01", sUsuario, sPasa, "", _
                       "TCODIGOUSUARIO", "CODIGO USUARIO", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                       "tPassword", "Contraseña", Encapsula(txtPassword.Text), "tBandaMagnetica", "Banda Magnetica", Encapsula(sBanda), _
                       "tGrupoUsuario", "Grupo Usuario", cboGrupoUsuario.BoundText, "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), "tUsuarioModifica", "Usuario Modifica", sUsuario)
                       
                       
                    If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                            Exit Sub
                    End If
                   'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                      
        
                   'Cambiar el SQL
                   Isql = "insert into TUSUARIO( " & _
                          "tCodigoUsuario, tDetallado, tResumido, tPassword, tBandaMagnetica, tGrupoUsuario, lActivo, fRegistro, tHuella, tUsuarioModifica) " & _
                          "values ('" & txtCodigo.Text & "', " & _
                                 " '" & txtDetallado.Text & "', " & _
                                 " '" & txtResumido.Text & "', " & _
                                 " '" & Encapsula(txtPassword.Text) & "', " & _
                                 " '" & Encapsula(sBanda) & "', " & _
                                 " '" & cboGrupoUsuario.BoundText & "', " & _
                                        chkActivo.Value & ", " & _
                                        "getdate(),'" & X & "','" & sUsuario & "')"
            
                       Cn.Execute Isql
                       
                         'HUELLA
                   If lGrabaHuella = True Then
                        If lHuellaDigitalPersona = True Then
                            X = ""
                            Dim blob11() As Byte
                            If Templ Is Nothing Then
                              X = ""
                            Else
                              blob11 = Templ.Serialize
                              X = arraytohex(blob11)
                            End If
                         Else
                                 X = arraytohex(g_FirstMinData)
                              
                         End If
                         
                         Cn.Execute "UPDATE TUSUARIO SET THUELLA='" & X & "' WHERE TCODIGOUSUARIO='" & txtCodigo.Text & "'"
                    End If
                   
   
                       If lMCPV Then
                        GrabaMozo
                       End If
                       frmUsuario.RsCabecera.Sort = "Codigo ASC"
                       frmUsuario.RsCabecera.Requery
                       frmUsuario.RsCabecera.MoveLast
                       MsgBox "Registro Guardado", vbInformation, sMensaje
                       ActivarBotones (True)
                       cmdTexto.Caption = "Registro " & IIf(frmUsuario.RsCabecera.RecordCount = 0, 0, frmUsuario.RsCabecera.AbsolutePosition) & " de " & frmUsuario.RsCabecera.RecordCount
               Else
               
                  'CESAR VALIDACION MODIFICACION DE GRUPO
                  xGrupo = Calcular("Select tGrupoUsuario As Codigo From  TUSUARIO Where tResumido = '" & sUsuario & "' ", Cn)
                  xlControlNivel = Calcular("Select ISNULL(lControlUsuario,0) As Codigo From  TPARAMETRO", Cn)
                  xGrupoCombo = cboGrupoUsuario.BoundText
                  
                  If xGrupo <> "00" Then
                        If xlControlNivel Then
                            If xGrupo > xGrupoCombo Then
                               MsgBox "No puede Registrar Usuarios con un Nivel Superior", vbExclamation, sMensaje
                               Exit Sub
                            End If
                        End If
                  End If
                  '------------------
                  
                  If Calcular("select count(tResumido) as codigo from tUsuario where tCodigoUsuario <>'" & txtCodigo & "' and tResumido='" & txtResumido & "'", Cn) > 0 Then
                     MsgBox "Usuario no permitido, intente de nuevo", vbExclamation, sMensaje
                     txtResumido.SetFocus
                     Exit Sub
                  End If
                  
                  If Calcular("select count(tPassword) as codigo from tUsuario where tCodigoUsuario <>'" & txtCodigo & "' and tPassword='" & Encapsula(txtPassword.Text) & "'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  
                  If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  
                  If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  
                  sBanda = Extrae(txtBanda.Text)
                  If txtBanda.Text <> "" And Calcular("select count(tBandaMagnetica) as Codigo from tUsuario where tCodigoUsuario<>'" & txtCodigo & "' and tBandaMagnetica='" & Encapsula(sBanda) & "'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  
                  sPasa = txtCodigo.Text
                      
                   'Inserta Movimiento auditoria
                   lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TUSUARIO", "USUARIO", "02", sUsuario, sPasa, "", _
                       "TCODIGOUSUARIO", "CODIGO USUARIO", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                       "tPassword", "Contraseña", Encapsula(txtPassword.Text), "tBandaMagnetica", "Banda Magnetica", Encapsula(sBanda), _
                       "tGrupoUsuario", "Grupo Usuario", cboGrupoUsuario.BoundText, "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), "tUsuarioModifica", "Usuario Modifica", sUsuario)
                       
                    If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                            Exit Sub
                    End If
                       
                   'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                  
                  'Cambiar el SQL
                  Isql = "update TUSUARIO set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "tPassword = '" & Encapsula(txtPassword.Text) & "', " & _
                         "tBandaMagnetica = '" & Encapsula(sBanda) & "', " & _
                         "tGrupoUsuario ='" & cboGrupoUsuario.BoundText & "', " & _
                         "tUsuarioModifica = '" & sUsuario & "', " & _
                         "lActivo =" & chkActivo.Value & _
                         " where tCodigoUsuario = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   
                   
                   'HUELLA
                   If lGrabaHuella = True Then
                        If lHuellaDigitalPersona = True Then
                            X = ""
                            Dim blob1() As Byte
                            If Templ Is Nothing Then
                              X = ""
                            Else
                              blob1 = Templ.Serialize
                              X = arraytohex(blob1)
                            End If
                         Else
                                 X = arraytohex(g_FirstMinData)
                              
                         End If
                         
                         Cn.Execute "UPDATE TUSUARIO SET THUELLA='" & X & "' WHERE TCODIGOUSUARIO='" & txtCodigo.Text & "'"
                    End If
                   
                   
                   If lMCPV Then
                    GrabaMozo
                   End If
                   nPos = frmUsuario.RsCabecera.Bookmark
                   frmUsuario.RsCabecera.Requery
                   If frmUsuario.RsCabecera.RecordCount = 0 Then
                      frmUsuario.RsCabecera.Filter = adFilterNone
                   End If
                   frmUsuario.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                Cn.Execute " update tusuario set lnuevoModificado=1 WHERE TCODIGOUSUARIO='" & txtCodigo.Text & "'"
                
          Case Is = 2 ' Eliminar
               If frmUsuario.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Usuario " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               
               sPasa = txtCodigo.Text
               'Inserta Movimiento auditoria
               lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TUSUARIO", "USUARIO", "03", sUsuario, sPasa, "", _
                          "TCODIGOUSUARIO", "CODIGO USUARIO", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                          
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                          
                   'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
               'Cambia el Delete
               Cn.Execute "delete from TUSUARIO where tCodigoUsuario = '" & txtCodigo & "'"
               frmUsuario.RsCabecera.Requery
               If frmUsuario.RsCabecera.RecordCount <> 0 Then
                  frmUsuario.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmUsuario.RsCabecera.RecordCount = 0, 0, frmUsuario.RsCabecera.AbsolutePosition) & " de " & frmUsuario.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
          
   End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    If lHuellaDigitalPersona = True Then
        'HUELLA
        Set Capture = New DPFPCapture
        Capture.StartCapture
        Set CreateFtrs = New DPFPFeatureExtraction
        Set CreateTempl = New DPFPEnrollment
        Samples.Caption = CreateTempl.FeaturesNeeded
        Set ConvertSample = New DPFPSampleConversion
        '----------------------------------
            fraDigitalPersona.Visible = True
            fraSecugen.Visible = False
        
    End If
    If lHuellaSecugen = True Then
            fraDigitalPersona.Visible = False
            fraSecugen.Visible = True
            g_firstStep = False
            g_secondStep = False
            g_vefiryResult = False
          FpLibXCapture1(0).CodeName = 2
            FpLibXCapture1(1).CodeName = 2
            ' hardware initialize
            ' FDP02(0), FDU02(1), FDU03(2), FDU04(3)
            Dim DefaultDevice As Integer
            DefaultDevice = 2   ' FDU03
'            For i = 0 To 4
'             If (Option1(i).value = True) Then
'                 DefaultDevice = i
'                 Exit For
'             End If
'            Next i
            FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ANSI378        ' ANSI 378
            FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
            FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
            '  FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
            ReDim g_FirstMinData(FpLibXCapture1(0).MinutiaeSize) As Byte
            ReDim g_SecondMinData(FpLibXCapture1(1).MinutiaeSize) As Byte
            'ReDim g_MatchingMinData(FpLibXCapture1(2).MinutiaeSize) As Byte

           ' FpLibXCapture1(2).CodeName = 2
            
 
            
            FpLibXCapture1(0).Refresh
            FpLibXCapture1(1).Refresh
           ' FpLibXCapture1(2).Refresh
    End If
     
    If lInfhotel Then
       cmdOpcion(0).Enabled = False
    End If
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Usuarios "
    'fraDetalle.Caption = Me.Caption
              
    'Llena todos los Combos
    LlenaCombos
            
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & IIf(frmUsuario.RsCabecera.RecordCount = 0, 0, frmUsuario.RsCabecera.AbsolutePosition) & " de " & frmUsuario.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
    
    
    'CESAR PERMISOS DE GRUPO
    Dim xGrupo As String
    Dim xRGrupoAcceso As Recordset
    xGrupo = Calcular("Select tGrupoUsuario As Codigo From TUSUARIO where tResumido = '" & sUsuario & "'", Cn)
        
    If xGrupo <> "00" Then
            Isql = "Select * From TGRUPOACCESO WHERE tGrupoUsuario = '" & xGrupo & "'"
            Set xRGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
            '10175010 AGREGAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10175010"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(0).Enabled = False
                Else
                   cmdOpcion(0).Enabled = True
                End If
            
            '10175020 MODIFICAR
                If xNuevo = "1" Then
                    cmdOpcion(1).Enabled = True
                Else
                    xRGrupoAcceso.Filter = "tCodigoAcceso=10175020"
                    If xRGrupoAcceso.RecordCount > 0 Then
                       cmdOpcion(1).Enabled = False
                    End If
                End If
                
            '10175030 MODIFICAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10175030"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(2).Enabled = False
                End If
                
                xNuevo = ""
    End If
    '---------------------------------------------

End Sub

Public Sub Grilla(lNuevo As String)
    If lNuevo = "1" Then
       xNuevo = "1"
    Else
       xNuevo = "0"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUsuarioDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    If lInfhotel Then
       cmdOpcion(0).Enabled = False
    Else
       cmdOpcion(0).Enabled = Activa
    End If
    cmdOpcion(2).Enabled = Activa
End Sub

Private Sub GrabaMozo()
    Dim sCorrela As String
    Dim tCodigo As String
    sCorrela = Calcular("Select tCodigo as Codigo From TTABLA Where tTabla = 'MOZO' And tValor = '" & txtCodigo.Text & "'", Cn)
    '!Codigo
    If sCorrela = "0" Then
        'Obtiene el Numero de Orden
        sCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='MOZO' ", Cn)
        tCodigo = Lib.Correlativo(sCorrela, 4)
        If IsNull(sCorrela) Or sCorrela = "" Then
            tCodigo = "0001"
        End If
        Sw = True
    Else
        tCodigo = sCorrela
        Sw = False
    End If
        
    If Sw Then
        
        Sw = False
        sBanda = Extrae(txtBanda.Text)
                  
        'Cambiar el SQL
        Isql = "insert into TTABLA( " & _
                "tTabla, tCodigo, tDetallado, tResumido, nboton, tValor, nValor, tIcono, nTamano, lActivo) " & _
                "values ('MOZO', " & _
                " '" & tCodigo & "', " & _
                " '" & txtDetallado.Text & "', " & _
                " '" & txtResumido.Text & "', " & _
                " 0, " & _
                " '" & txtCodigo.Text & "', " & _
                "0" & ", " & _
                " '" & Encapsula(sBanda) & "', " & _
                "0" & ", " & _
                chkActivo.Value & ") "
        Cn.Execute Isql
    Else
        sBanda = Extrae(txtBanda.Text)
        'Cambiar el SQL
        Isql = "update TTABLA set " & _
                "tDetallado ='" & txtDetallado.Text & "', " & _
                "tResumido ='" & txtResumido.Text & "', " & _
                "nBoton =" & "''" & ", " & _
                "tIcono = '" & Encapsula(sBanda) & "', " & _
                "nValor =" & "0" & ", " & _
                "nTamano =" & "0" & ", " & _
                "lActivo =" & chkActivo.Value & ", lReplica=1 " & _
                " where tTAbla = 'MOZO' and tCodigo = '" & sCorrela & "'"
                
        Cn.Execute Isql
    End If
End Sub




'HUELLA
Private Sub DrawPicture(ByVal Pict As IPictureDisp)
 Set HiddenPict.Picture = Pict
 Picture1.PaintPicture HiddenPict.Picture, _
       0, 0, Picture1.ScaleWidth, _
       Picture1.ScaleHeight, _
       0, 0, HiddenPict.ScaleWidth, _
       HiddenPict.ScaleHeight, vbSrcCopy
 Picture1.Picture = Picture1.Image
End Sub
Private Sub ReportStatus(ByVal str As String)
 ' Add string to list box.
 Status.AddItem (str)
 ' Move list box selection down.
 Status.ListIndex = Status.NewIndex
End Sub

Private Sub Capture_OnReaderConnect(ByVal ReaderSerNum As String)
 ReportStatus ("Lector de Huella Dactilar Conectado")
End Sub

Private Sub Capture_OnReaderDisconnect(ByVal ReaderSerNum As String)
 ReportStatus ("Lector de Huella Dactilar Desconectado")
End Sub

Private Sub Capture_OnFingerTouch(ByVal ReaderSerNum As String)
    lblHuella.Visible = False
 ReportStatus ("Lector de Huella Dactilar presionado")
End Sub
Private Sub Capture_OnFingerGone(ByVal ReaderSerNum As String)
 ReportStatus ("El dedo fue retirado del Lector de Huella Dactilar")
End Sub
Private Sub Capture_OnSampleQuality(ByVal ReaderSerNum As String, ByVal Feedback As DPFPCaptureFeedbackEnum)
 If Feedback = CaptureFeedbackGood Then
  ReportStatus ("La calidad de la huella obtenida es buena")
  Else
  ReportStatus ("La calidad de la huella obtenida es no es adecuada")
  End If
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
 lGrabaHuella = False
 Dim Feedback As DPFPCaptureFeedbackEnum
 ReportStatus ("Huella dactilar capturada")
 ' Draw fingerprint image.
 DrawPicture ConvertSample.ConvertToPicture(Sample)
 ' Process sample and create feature set for purpose of enrollment.
 Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeEnrollment)
 ' Quality of sample is not good enough to produce feature set.
 If Feedback = CaptureFeedbackGood Then
  ReportStatus ("Patrón de Huella Dactilar creado.")
  Prompt.Caption = "Presione el lector de Huella Dactilar con el mismo dedo"
  ' Add feature set to template.
  CreateTempl.AddFeatures CreateFtrs.FeatureSet
  ' Show number of samples needed to complete template.
  Samples.Caption = CreateTempl.FeaturesNeeded
  ' Check if template has been created.
  If CreateTempl.TemplateStatus = TemplateStatusTemplateReady Then
  '  MainFrame.SetTemplete CreateTempl.Template
    ' Template has been created, so stop capturing samples.
    Capture.StopCapture
    lGrabaHuella = True
    Prompt.Caption = "Puede guardar el registro."
    
       Me.SetTemplete CreateTempl.Template
   ' MsgBox "The fingerprint template was created."
  End If
 End If
 End Sub


Private Function arraytohex(arr() As Byte) As String

Dim templatestr As String
Dim tempstr As String
Dim i As Integer

templatestr = ""

For i = LBound(arr) To UBound(arr)

tempstr = Hex$(arr(i))

If Len(tempstr) = 1 Then tempstr = "0" + tempstr 'padhex
templatestr = templatestr + tempstr
Next i
arraytohex = templatestr

End Function
 
Public Sub hextoarray(inphex As String, outarray() As Byte)

ReDim outarray(0 To Len(inphex) / 2)

Dim i As Integer

For i = 1 To Len(inphex) Step 2
outarray(((i + 1) / 2) - 1) = val("&H" + Mid$(inphex, i, 2))
Next i

End Sub
Public Sub SetTemplete(ByVal Template As Object)
 Set Templ = Template
End Sub

Public Sub limpiaVariables()
    Picture1.Picture = Nothing
    lblHuella.Visible = False
    Prompt.Caption = "Toque el lector de Huella Dactilar"
    Samples.Caption = ""
    Status.Clear

 Set Capture = New DPFPCapture
 Capture.StartCapture
 Set CreateFtrs = New DPFPFeatureExtraction
 Set CreateTempl = New DPFPEnrollment
 Samples.Caption = CreateTempl.FeaturesNeeded
 Set ConvertSample = New DPFPSampleConversion
End Sub


Private Sub TemplateFormat_Click(Index As Integer)

    g_firstStep = False
    g_secondStep = False
    g_vefiryResult = False

    ' Set format of template to use.
    ' two modes of FpLibXVerify1 and FpLibXCapture1 MUST be the same minutiae mode
    ' otherwise comment out to use MinutiaeMode property
    Select Case Index
    Case 0
        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_SG400         ' By default
        FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_SG400
        FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_SG400
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_SG400
        
    Case 1
        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ANSI378        ' ANSI 378
        FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        
    Case 2
        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ISO19794      ' ISO 19794-2
        FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_ISO19794
        FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_ISO19794
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ISO19794
    End Select

    'Redim these variables after setting the MinutiaeMode,
    'because size of data is different for each MinutiaeMode
    ReDim g_FirstMinData(FpLibXCapture1(0).MinutiaeSize) As Byte
    ReDim g_SecondMinData(FpLibXCapture1(1).MinutiaeSize) As Byte
    ReDim g_MatchingMinData(FpLibXCapture1(2).MinutiaeSize) As Byte


End Sub

Private Sub FpLibXCapture1_Paint(Index As Integer, ByVal hwnd As stdole.OLE_HANDLE, ByVal hdc As stdole.OLE_HANDLE)
    Select Case Index
    Case 0
        If g_firstStep = False Then
            TextOut hdc, 30, 70, "Click Aquí..", 12
        End If
        
    Case 1
        If g_firstStep = True And g_secondStep = False Then
            TextOut hdc, 30, 70, "Click Aquí..", 12
        End If
        
    Case 2
        If g_firstStep = True And g_secondStep = True And g_vefiryResult = False Then
            TextOut hdc, 30, 70, "Click Aquí..", 12
        End If
    
    End Select
End Sub

Private Sub FpLibXCapture1_Click(Index As Integer)
On Error GoTo err2014
Dim register_quality, verify_quality As Integer
Dim Timeout As Integer

register_quality = 50
verify_quality = 30
Timeout = 6000 ' 6seconds

    Select Case (Index)
    Case 0
        Result = FpLibXCapture1(Index).LiveCapture(Timeout, register_quality)
        FpLibXCapture1(Index).GetMinutiaeData (g_FirstMinData)
        
        g_firstStep = True
        g_secondStep = False
        g_vefiryResult = False

'    Dim x As String
'     x = arraytohex(g_FirstMinData)
'    Debug.Print x
'
'    'Dim Cn As Connection
'
'    Set Cn = New Connection
'    Cn.Provider = "SQLOLEDB"
'
'    Cn.CursorLocation = adUseServer
'    Cn.ConnectionString = "User ID=sa" & _
'                          ";password=sistemas" & _
'                          ";Data Source=INFLIMDES02" & _
'                          ";Initial Catalog=Prueba_xxx"
'
'    Cn.CommandTimeout = 300
'    Cn.Open
'    Cn.Execute "UPDATE thuella set huella='" & x & "' where codigo='00001'"
        
        FpLibXCapture1(1).Clear
       ' FpLibXCapture1(2).Clear
        LabelMessage.Caption = ""
    
    Case 1
        If g_firstStep = True Then
            Result = FpLibXCapture1(Index).LiveCapture(Timeout, register_quality)
            FpLibXCapture1(Index).GetMinutiaeData (g_SecondMinData)
            
            'you can get the score
            score = FpLibXVerify1.GetMatchingScore(g_FirstMinData, g_SecondMinData)
            
            Result = FpLibXVerify1.Register(g_FirstMinData, g_SecondMinData)
            If Result = True Then
                g_secondStep = True
                LabelMessage.Caption = "Patrón de Huellas Correcto. Puede Guardar el Registro"
                 lGrabaHuella = True

                g_vefiryResult = False
              '  FpLibXCapture1(2).Clear
            Else
                g_secondStep = False
             '   FpLibXCapture1(2).Clear
                If FpLibXVerify1.ErrorCode = ERROR_VERIFY_FAKE Then
                    LabelMessage.Caption = "Retire el dedo y tocar de nuevo , por favor"
                     lGrabaHuella = False

                Else
                    LabelMessage.Caption = "Patrón de Huellas Incorrecto, Intente de Nuevo"
                     lGrabaHuella = False

                End If
            End If
            
        End If
        
    Case 2
        If g_secondStep = True Then
            bResult = FpLibXCapture1(Index).LiveCapture(Timeout, verify_quality)
            FpLibXCapture1(Index).GetMinutiaeData (g_MatchingMinData)
            
                Dim huella() As Byte
          
                   Set Cn = New Connection
                   Cn.Provider = "SQLOLEDB"
                   
                   Cn.CursorLocation = adUseServer
                   Cn.ConnectionString = "User ID=sa" & _
                                         ";password=sistemas" & _
                                         ";Data Source=INFLIMDES02" & _
                                         ";Initial Catalog=Prueba_xxx"
                                         
                   Cn.CommandTimeout = 300
                   Cn.Open
                   Dim rst1 As ADODB.Recordset
                   Set rst1 = New ADODB.Recordset
                   Dim p As String
                  Set rst1 = Cn.Execute("select huella from thuella where codigo='00001'")
                   p = rst1.Fields(0)
                    hextoarray p, huella()
                    
            
            
            
            
            
            Result = FpLibXVerify1.Verify(huella, g_MatchingMinData)
            
            
            If Result = True Then
                g_vefiryResult = True
                LabelMessage.Caption = "Identical"
            Else
                g_vefiryResult = False
                If FpLibXVerify1.ErrorCode = ERROR_VERIFY_FAKE Then
                    LabelMessage.Caption = "Release your finger and touch on again, please"
                Else
                    LabelMessage.Caption = "NOT Identical"
                End If
                
            End If
        End If
        
    End Select
 
    FpLibXCapture1(0).Refresh
    FpLibXCapture1(1).Refresh
 '   FpLibXCapture1(2).Refreshe
 Exit Sub
 
err2014:
   FpLibXCapture1(0).Refresh
    FpLibXCapture1(1).Refresh
End Sub

