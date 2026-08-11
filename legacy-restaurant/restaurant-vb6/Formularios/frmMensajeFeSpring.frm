VERSION 5.00
Begin VB.Form frmMensajeFeSpring 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFeSpring 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      Begin VB.CommandButton cmdEnvioFe 
         Appearance      =   0  'Flat
         Caption         =   "Re - Enviar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   2925
         Width           =   1800
      End
      Begin VB.TextBox txtRespuesta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   840
         Left            =   260
         MaxLength       =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   5800
      End
      Begin VB.CommandButton cmdEnvioFe 
         Caption         =   "QR Alterno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Top             =   2925
         Width           =   1800
      End
      Begin VB.CommandButton cmdEnvioFe 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   4425
         TabIndex        =   1
         Top             =   2925
         Width           =   1800
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   360
         TabIndex        =   12
         Top             =   1620
         Width           =   1080
      End
      Begin VB.Label lblDocumento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1485
         TabIndex        =   11
         Top             =   1575
         Width           =   1695
      End
      Begin VB.Label lblResultado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4440
         TabIndex        =   10
         Top             =   1575
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3405
         TabIndex        =   9
         Top             =   1620
         Width           =   960
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmMensajeFeSpring.frx":0000
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo hash almacenado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1185
         TabIndex        =   6
         Top             =   1155
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a spring."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1170
         TabIndex        =   5
         Top             =   870
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmMensajeFeSpring.frx":0213
         Top             =   855
         Width           =   240
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
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMensajeFeSpring.frx":0426
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmMensajeFeSpring.frx":0768
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a Spring......."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   435
         Width           =   6080
      End
   End
End
Attribute VB_Name = "frmMensajeFeSpring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oVenta As Integer
Public sMesa As String
Public sDocumento As String
Public sTipoDocumento As String
Public sCaja As String
Public sPedido As String
Public lQrInf As Boolean

Dim xUltimoCorrelativo As String
Dim RsRespuesta As ADODB.Recordset

Public lEnvio As Boolean

Private Sub cmdEnvioFe_Click(Index As Integer)
    
    txtRespuesta.Text = ""
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    
    Select Case Index
        Case Is = 2 'Re-envia a spring
            
            Dim Metodo As Integer
            If RsRespuesta!resultado = 2 Then
                Metodo = 2
                imgProceso(1).Visible = True
            ElseIf RsRespuesta!resultado = 3 Then
                Metodo = 3
                imgProceso(1).Visible = True
                MsgBox ("Metodo no disponible para respuesta 3.")
                Exit Sub
            Else
                Metodo = 1
            End If
            Sleep 1500
            If Not INSERTAFE_SPRING(sDocumento, Metodo) Then
                ResultadosErroneos
                lEnvio = False
                Exit Sub
            End If
            imgProceso(0).Visible = True
            imgProceso(1).Visible = True
            txtRespuesta.Text = "Se realizo el proceso correctamente."
            Sleep 1500
            
            lEnvio = True
            Unload Me
            
        Case Is = 1 'Genera QR alterno
            lQrInf = True
            lEnvio = True
            Unload Me
            
        Case Is = 0 'Salir
            If Not lEnvio Then
                RollBack
            End If
            Unload Me

    End Select
    
End Sub

Private Sub Form_Load()
    lQrInf = False
    lEnvio = True
    
    lblPaso1.Visible = True
    lblPaso2.Visible = True
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    
    Label2.Caption = "   Proceso de envio de documento a Spring...."
    lblPaso1.Caption = "Enviando información de documento a Spring."
    lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                    
    If Not INSERTAFE_SPRING(sDocumento, 1) Then
        ResultadosErroneos
        lEnvio = False
        Exit Sub
    End If
        
    imgProceso(0).Visible = True
    imgProceso(1).Visible = True
    txtRespuesta.Text = "Se realizo el proceso correctamente."
    lblResultado.Caption = "1"
    lblDocumento.Caption = sDocumento

    lEnvio = True
    cmdEnvioFe(0).Enabled = True
    cmdEnvioFe(1).Enabled = False
    cmdEnvioFe(2).Enabled = False
    
End Sub

Private Sub ResultadosErroneos()
    Set RsRespuesta = Lib.OpenRecordset("exec usp_TransSpring_ObtieneRespuesta '" & sDocumento & "'", Cn)
    If RsRespuesta.RecordCount > 0 Then
        If RsRespuesta!resultado = 2 Then
            imgProceso(1).Visible = True
            imgProceso(3).Visible = True
            cmdEnvioFe(0).Enabled = False
            cmdEnvioFe(2).Enabled = True
        ElseIf RsRespuesta!resultado = 3 Then
            imgProceso(1).Visible = True
            imgProceso(3).Visible = True
            cmdEnvioFe(0).Enabled = False
            cmdEnvioFe(2).Enabled = False
        Else
            imgProceso(2).Visible = True
            imgProceso(3).Visible = True
            If oVenta = 1 Or oVenta = 4 Then
                cmdEnvioFe(0).Visible = True
                'RollBack
            Else
                cmdEnvioFe(0).Visible = False
            End If
        End If
        txtRespuesta.Text = RsRespuesta!Mensaje
        lblResultado.Caption = RsRespuesta!resultado
        lblDocumento.Caption = RsRespuesta!Documento
    Else
        cmdEnvioFe(0).Enabled = False
        cmdEnvioFe(1).Enabled = False
        cmdEnvioFe(2).Enabled = False
    End If
End Sub

Private Sub RollBack()
    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub
