VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWinSock 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3855
      Begin VB.Label lblRespuesta 
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Conectando ........."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3900
      End
      Begin VB.Label lblespera 
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1080
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Scan 
      Left            =   600
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer 
      Left            =   120
      Top             =   1320
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Facturación Electronica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2490
   End
End
Attribute VB_Name = "frmWinSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ip As String
Public Puerto As Long
Public TipoOper As Integer
Public Resp As String
Dim Contador As Integer
Public TramaTCPIP As String
'Notas:
'tipooper=0 ----> Para FE Paperlees
Private Sub Form_Load()
    Resp = ""
    Contador = 1
    Me.Caption = "Conectando a Perlees :" & IPpape & ":" & PUERTOpape
    Centrar Me
    Scan.Close
    Scan.Connect IPpape, PUERTOpape
    Timer.Interval = 1000
    Timer.Enabled = True
    lblRespuesta.Caption = "Conectando........... " + IPpape & " : " & PUERTOpape
End Sub



Private Sub Scan_Connect()
    If TipoOper = 0 Then
        wEnter = True
        lblRespuesta.Caption = "Conectado!!!!!!!!! " + IPpape & " : " & PUERTOpape
        Scan.SendData TramaTCPIP
    End If
End Sub

Private Sub Scan_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   Resp = Number & " " & Description
   wEnter = False
   Scan.Close
   Unload Me
End Sub
Private Sub Scan_DataArrival(ByVal bytesTotal As Long)
On Error GoTo fin
    Dim Msg As String
    Scan.GetData Msg
    lblRespuesta.Caption = "Recibiendo respuesta............ " + IPpape & " : " & PUERTOpape
    Resp = Msg
    Scan.Close
    wEnter = True
    Unload Me
    Exit Sub
fin:
    Resp = "Error: " & Resp & vbNewLine & error
    wEnter = False
    Scan.Close
    Unload Me
End Sub
Private Sub Timer_Timer()
    If TipoOper = 0 Then
        lblespera.Caption = "Tiempo Espera de Conexion: " & Contador & " Seg"
        If Contador = 9 Then
            wEnter = False
            lblespera.Caption = ""
            Timer.Interval = 0
            Timer.Enabled = False
            Scan.Close
            Resp = "Tiempo de coneccion paperlees Agotado!!!"
            Unload Me
        Else
            Contador = Contador + 1
        End If
    End If
End Sub

'Private Sub OpenTCPIP()
'On Error GoTo fin
'        'Constant                   Value   Description
'        '-----------------------------------------------------
'        'sckClosed                  0       Default. Closed     --Defecto.Cerrado
'        'sckOpen                    1       Open                --Abierto
'        'sckListening               2       Listening           --Escuchando
'        'sckConnectionPending       3       Connection pending  --Conexión pendiente
'        'sckResolvingHost           4       Resolving host      --Resolviendo host
'        'sckHostResolved            5       Host resolved       --Host resuelto
'        'sckConnecting              6       Connecting          --Conectando
'        'sckConnected               7       Connected           --Conectado
'        'sckClosing                 8       Peer is closing the connection --Peer está cerrando la conexión
'        'sckError                   9       Error --Error
'End Sub
