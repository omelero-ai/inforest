VERSION 5.00
Begin VB.Form frmConsultaSaldo 
   Caption         =   "Consultar Saldo"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCombo 
      Caption         =   "Salir"
      Height          =   555
      Index           =   0
      Left            =   6075
      Picture         =   "frmConsultaSaldo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2025
      Width           =   1230
   End
   Begin VB.CommandButton cmdNumPad 
      Caption         =   "NumPad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4050
      TabIndex        =   0
      Top             =   270
      Width           =   1275
   End
   Begin VB.Label txtmoneda 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   315
      Left            =   2025
      TabIndex        =   10
      Top             =   690
      Width           =   660
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Número de Tarjeta :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   307
      Width           =   1785
   End
   Begin VB.Label txtBeneficiario 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUAN JOSE RAMOS ALBUQUERQUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   315
      Left            =   2025
      TabIndex        =   7
      Top             =   1530
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Beneficiario :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   750
      TabIndex        =   6
      Top             =   1567
      Width           =   1155
   End
   Begin VB.Label txtSocio 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JOSE LUIS RAMOS ONAGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   315
      Left            =   2025
      TabIndex        =   5
      Top             =   1110
      Width           =   5295
   End
   Begin VB.Label txtMonto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10,520.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2745
      TabIndex        =   4
      Top             =   690
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Socio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1290
      TabIndex        =   3
      Top             =   1147
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Saldo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1275
      TabIndex        =   2
      Top             =   727
      Width           =   630
   End
   Begin VB.Label txtNT 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1086"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2025
      TabIndex        =   1
      Top             =   270
      Width           =   1875
   End
End
Attribute VB_Name = "frmConsultaSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsConsulta As Recordset

Private Sub cmdCombo_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdNumPad_Click()
    sTipo = "Numero"
    frmNumPad.Show vbModal
    txtNT.Caption = IIf(wEnter = True, Mid(sDescrip, 1, 16), txtNT.Caption)
    If wEnter Then
       ConsultaSaldo
    End If
End Sub

Private Sub Form_Load()
   Centrar Me
   Limpiar
   txtmoneda.Caption = sMonN
End Sub

Public Sub ConsultaSaldo()
   sSYBASE = Trim(LeerIni(App.Path + "\INFOREST.INI", "CONEXION", "SYBASE", ""))
   Set CnSiab = New Connection
   CnSiab.Provider = SYPROV '"ASAProv.80"
   CnSiab.CursorLocation = adUseServer
   CnSiab.ConnectionString = sSYBASE
   CnSiab.CommandTimeout = 250
   On Error GoTo ErrorEcuador
   CnSiab.Open
   
   Isql = "select * from tarj_consumo_maestro where estado='A' and num_tarjeta = '" & txtNT.Caption & "'"
   Set RsConsulta = Lib.OpenRecordset(Isql, CnSiab)
   If RsConsulta.RecordCount > 0 Then
      MuestraResultados
   Else
      NoMuestraResultados
   End If
   CnSiab.Close
   On Error GoTo 0

Exit Sub
ErrorEcuador:
   MsgBox "Error Fatal: No existe conexión con el sistema SIAB", vbCritical, sMensaje
End Sub

Public Sub MuestraResultados()
   txtNT.ForeColor = &HFF0000
   
   txtMonto.Caption = Format(RsConsulta!Saldo, "##,##0.00")
   txtSocio.Caption = Calcular("select nom_cli as Codigo from clientes where cod_cli= '" & RsConsulta!Cod_cli & "'", CnSiab)
   txtBeneficiario.Caption = RsConsulta!Beneficiario
End Sub

Public Sub NoMuestraResultados()
   txtNT.ForeColor = &HFF&
   txtNT.Caption = "No Econtrado"
   txtMonto.Caption = "0.00"
   txtSocio.Caption = ""
   txtBeneficiario.Caption = ""
End Sub

Public Sub Limpiar()
   txtNT.Caption = ""
   txtMonto.Caption = "0.00"
   txtSocio.Caption = ""
   txtBeneficiario.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsConsulta = Nothing
    Set frmConsultaSaldo = Nothing
End Sub

