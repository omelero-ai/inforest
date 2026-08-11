VERSION 5.00
Begin VB.Form frmAcceso 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Acceso al Sistema"
   ClientHeight    =   9990
   ClientLeft      =   3435
   ClientTop       =   2820
   ClientWidth     =   12720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAcceso.frx":0442
   ScaleHeight     =   9990
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   9480
      TabIndex        =   13
      Top             =   5040
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   20160
      Top             =   6960
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H007EC642&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5400
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualización Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   6555
      End
      Begin VB.Label LblVer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   480
         Width           =   6615
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   20040
      Picture         =   "frmAcceso.frx":178EC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   1275
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   6600
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3720
      Width           =   3705
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   480
      Left            =   6600
      TabIndex        =   0
      Top             =   2880
      Width           =   3660
   End
   Begin VB.CommandButton cmdOpcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   4
      Left            =   20160
      Picture         =   "frmAcceso.frx":179EE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   11040
      Width           =   1080
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "PassWord"
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
      Index           =   3
      Left            =   20040
      TabIndex        =   6
      Top             =   9840
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Usuario"
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
      Index           =   2
      Left            =   20040
      TabIndex        =   5
      Top             =   10440
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   20040
      Picture         =   "frmAcceso.frx":18730
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   1275
   End
   Begin VB.TextBox txtCaja1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   19560
      TabIndex        =   2
      Top             =   7800
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   7950
      Picture         =   "frmAcceso.frx":18832
      Top             =   920
      Width           =   1500
   End
   Begin VB.Image imgVentaLogin 
      Height          =   9000
      Left            =   840
      Picture         =   "frmAcceso.frx":1A4B1
      Top             =   8880
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Image imgadministracionLogin 
      Height          =   9000
      Left            =   240
      Picture         =   "frmAcceso.frx":2C121
      Top             =   8400
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Image imgconsultaLogin 
      Height          =   9000
      Left            =   0
      Picture         =   "frmAcceso.frx":3F177
      Top             =   1300
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Image imginiciasesion 
      Height          =   735
      Left            =   6480
      Top             =   6120
      Width           =   4455
   End
   Begin VB.Image imgNewProceso 
      Height          =   735
      Left            =   6360
      Top             =   7080
      Width           =   4695
   End
   Begin VB.Image imgOpcion 
      Height          =   615
      Index           =   2
      Left            =   10320
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image ImagePais 
      Height          =   435
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image imgNewOpcion 
      Height          =   495
      Index           =   0
      Left            =   10080
      Top             =   9000
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   10920
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label txtCaja 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   405
      Width           =   2655
   End
   Begin VB.Image imgOpcion 
      Height          =   615
      Index           =   3
      Left            =   10320
      Top             =   3700
      Width           =   615
   End
   Begin VB.Image imgOpcion 
      Height          =   735
      Index           =   1
      Left            =   11280
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgOpcion 
      Height          =   1095
      Index           =   4
      Left            =   8280
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label txtBD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   8760
      Width           =   6375
   End
   Begin VB.Image imgInforest 
      Height          =   1290
      Left            =   12000
      Picture         =   "frmAcceso.frx":4E43F
      Top             =   2040
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Image imgAdministracion 
      Height          =   1290
      Left            =   0
      Picture         =   "frmAcceso.frx":52D02
      Top             =   0
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Image imgconsulta 
      Height          =   1290
      Left            =   12000
      Picture         =   "frmAcceso.frx":5A6D4
      Top             =   720
      Visible         =   0   'False
      Width           =   5400
   End
End
Attribute VB_Name = "frmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUsuario As Recordset
Dim RsAcceso As Recordset
Dim i As Integer
Dim validaCierreHK As Boolean
Dim fso As Object
'HUELLA
Dim lModulo As String
Dim marcacion As Integer



'TRANSPARENCIA
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
                
                Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 ' Aceptar
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                 
                'Chequea Datos
                
                wEnter = False
                
                If txtUsuario.Text = "" Then MsgBox "Ingrese su usuario", vbExclamation, sMensaje: txtUsuario.SetFocus: Exit Sub
                If txtPassword.Text = "" Then MsgBox "Ingrese su password", vbExclamation, sMensaje: txtPassword.SetFocus: Exit Sub
                RsUsuario.MoveFirst
                RsUsuario.Find ("tResumido = '" & Trim(txtUsuario.Text) & "' ")
                   
                If RsUsuario.EOF Then
                    i = i + 1
                    MsgBox "Usuario No Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtUsuario.SetFocus
                Else
                    If Desencapsula(RsUsuario!tpassword) = UCase(txtPassword.Text) Or (Desencapsula(RsUsuario!tBandaMagnetica) = UCase(Extrae(txtPassword.Text)) And RsUsuario!tBandaMagnetica <> "") Then
                        sPassword = UCase(txtPassword.Text)
                        sUsuario = UCase(txtUsuario.Text)
                        xUsuario = Mid(RsUsuario!tCodigoUsuario, 3, 3)
                        tcodigoUsuarioA = RsUsuario!tCodigoUsuario 'para controler

                        wEnter = True
                       
                        Open App.path & "\USUARIO.INI" For Output As #1
                        Print #1, IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
                        Close #1
                        If lHARDkey Then
                            validaCierreHK = False
                        End If
             
                        'audirotia
                        
                        registroAccesoAuditoria "I", sUsuario
                        If nCorrelativoAcceso = -1 Then
                        End
                        End If
                        
                        'auditoria
                        
                        Unload Me
                    Else
                        i = i + 1
                        MsgBox "Password Erroneo", vbCritical, sMensaje
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                    End If
                End If
                If i = 4 Then End
           
           Case Is = 1 ' Cancelar
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim Result As Boolean
                        Select Case sModulo
                            Case "INFOREST"
                                Result = hk.FinalizarConexion(Aplicacion.PuntoVenta) 'InfhotelHK.PuntoVenta)
                            Case "ADMINISTRACION"
                                Result = hk.FinalizarConexion(Aplicacion.Administracion) 'InfhotelHK.Administracion)
                            Case "CONSULTA"
                                Result = hk.FinalizarConexion(Aplicacion.Consultas) 'InfhotelHK.Consultas)
                            Case Else
                        End Select
                        End
                    End If
                '--------------------------------------------------------------
                End If
                End
                
           Case Is = 2 ' Usuario
                frmKeyBoard.txtResultado.Text = txtUsuario.Text
                frmKeyBoard.Show vbModal
                If wEnter Then
                   txtUsuario.Text = sDescrip
                End If
                wEnter = False
                
           Case Is = 3 ' Password
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   txtPassword.Text = sDescrip
                End If
                wEnter = False
                
           Case Is = 4 'HUELLA
                wenterHuellaSup = False
                lUsuarioHuella = True
                frmVerificacionHuellaSup.Opcion lModulo
                frmVerificacionHuellaSup.Show vbModal
                If wenterHuellaSup Then
                    wEnter = True
                    lUsuarioHuella = False
                    sUsuario = sVar1
                    Unload Me
                End If

    End Select
End Sub

Private Sub Form_Activate()
   On Error GoTo fin
   If sUsuario = "" Then
      txtUsuario.SetFocus
   Else
      txtPassword.SetFocus
   End If
fin:
   lUsuarioHuella = False

End Sub

Private Sub Form_Load()
  On Error GoTo Err:
  ElminaTriguer
  '************** del flash
'  If lAlmacen Then
'      Tiempo.Interval = 1
'   Else
'      Tiempo.Interval = 1500
'   End If
    Me.Width = 12000
    Me.Height = 9000
    If lAlmacen Then
      Actualiza
   End If
   If sModulo = "INTEGRACION" Then
      Integra
   End If
   
   If sModulo = "ADMINISTRACION" Then
     Me.imgAdministracion.Top = 0
     Me.imgAdministracion.Left = 0
     Me.imgAdministracion.Visible = True
     
    Me.imgadministracionLogin.Visible = True
    Me.imgadministracionLogin.Top = 1300
    Me.imgadministracionLogin.Left = 0
   End If
   If sModulo = "INFOREST" Then
    Me.imgInforest.Top = 0
    Me.imgInforest.Left = 0
    Me.imgInforest.Visible = True
    Me.Frame1.Visible = False
    
    Me.imgVentaLogin.Visible = True
    Me.imgVentaLogin.Top = 1300
    Me.imgVentaLogin.Left = 0
   End If
   If sModulo = "CONSULTA" Then
    Me.imgconsulta.Top = 0
    Me.imgconsulta.Left = 0
    Me.imgconsulta.Visible = True
    
    Me.imgconsultaLogin.Visible = True
    Me.imgconsultaLogin.Top = 1300
    Me.imgconsultaLogin.Left = 0
   End If
   'TIPO CAMBIO
   If pais = "002" Then
      RTipoCambio
   End If
  '*************************
  Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(App.path & "\bmps\Paises\" & pais & ".jpg") Then
               ImagePais.Picture = LoadPicture(App.path & "\bmps\Paises\" & pais & ".jpg")
    End If
    Set fso = Nothing
  
  If lHARDkey Then
        validaCierreHK = True
  End If
  Open App.path & "\USUARIO.INI" For Input As #1   ' Abre el archivo para recibir los datos.
  Do While Not EOF(1)                              ' Repite el bucle hasta el final del archivo.
     Input #1, sUsuario                            ' Lee el carácter en dos variables
  Loop
  Close #1
  txtUsuario.Text = sUsuario

  AccesoInicio
  Screen.MousePointer = vbDefault
Exit Sub

Err:
   txtUsuario.Text = ""
   AccesoInicio
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lHARDkey Then
        If validaCierreHK Then
            Dim Verifica As Boolean
            Select Case sModulo
                Case "INFOREST"
                    Verifica = hk.FinalizarConexion(Aplicacion.PuntoVenta)
                Case "ADMINISTRACION"
                    Verifica = hk.FinalizarConexion(Aplicacion.Administracion)
                Case "CONSULTA"
                    Verifica = hk.FinalizarConexion(Aplicacion.Consultas)
                Case Else
            End Select
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsUsuario = Nothing
    Set frmAcceso = Nothing
End Sub



Private Sub Image1_Click()
    frmAbout.Show vbModal
End Sub







Private Sub imginiciasesion_Click()
    If Trim(txtPassword.Text) <> "" Then
      imgOpcion_Click (0)
    End If
End Sub

Private Sub imgNewOpcion_Click(Index As Integer)
    
    If lCambiarContrasena Then
        Call Aplicar_Transparencia(Me.hwnd, CByte(200))
         frmCambiarContrasenia.Show vbModal
        Call Aplicar_Transparencia(Me.hwnd, CByte(250))
    Else
        MsgBox "Proceso no habilitado, indicar al administrador del sistema la activación en parametros generales!", vbExclamation, sMensaje
    End If
    
End Sub
Private Sub imgNewProceso_Click()
    marcacion = 1
    imgOpcion_Click 5
End Sub

Private Sub imgOpcion_Click(Index As Integer)
 Select Case Index
           Case Is = 0 ' Aceptar
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                 
                'Chequea Datos
                
                wEnter = False
                
                If txtUsuario.Text = "" Then MsgBox "Ingrese su usuario", vbExclamation, sMensaje: txtUsuario.SetFocus: Exit Sub
                If txtPassword.Text = "" Then MsgBox "Ingrese su password", vbExclamation, sMensaje: txtPassword.SetFocus: Exit Sub
                RsUsuario.MoveFirst
                RsUsuario.Find ("tResumido = '" & Trim(txtUsuario.Text) & "' ")
                   
                If RsUsuario.EOF Then
                    i = i + 1
                    MsgBox "Usuario No Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtUsuario.SetFocus
                Else
                    If Desencapsula(RsUsuario!tpassword) = UCase(txtPassword.Text) Or (Desencapsula(RsUsuario!tBandaMagnetica) = UCase(Extrae(txtPassword.Text)) And RsUsuario!tBandaMagnetica <> "") Then
                    
                        'Valida Version
                        Call VerificaVersionInfoRest(sModulo)

                        sPassword = UCase(txtPassword.Text)
                        sUsuario = UCase(txtUsuario.Text)
                        xUsuario = Mid(RsUsuario!tCodigoUsuario, 3, 3)
                        tcodigoUsuarioA = RsUsuario!tCodigoUsuario 'para controler

                        wEnter = True
                       
                        Open App.path & "\USUARIO.INI" For Output As #1
                        Print #1, IIf(Mid(sUsuario, 1, 1) = "*", Mid(sUsuario, 2, 15), sUsuario)
                        Close #1
                        If lHARDkey Then
                            validaCierreHK = False
                        End If
             
                        'audirotia
                        
                        registroAccesoAuditoria "I", sUsuario
                        If nCorrelativoAcceso = -1 Then
                        End
                        End If
                        'auditoria
                        
                        Unload Me
                    Else
                        i = i + 1
                        MsgBox "Password Erroneo", vbCritical, sMensaje
                        txtPassword.Text = ""
                        txtPassword.SetFocus
                    End If
                End If
                If i = 4 Then End
           
           Case Is = 1 ' Cancelar
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim Result As Boolean
                        Select Case sModulo
                            Case "INFOREST"
                                Result = hk.FinalizarConexion(Aplicacion.PuntoVenta) 'InfhotelHK.PuntoVenta)
                            Case "ADMINISTRACION"
                                Result = hk.FinalizarConexion(Aplicacion.Administracion) 'InfhotelHK.Administracion)
                            Case "CONSULTA"
                                Result = hk.FinalizarConexion(Aplicacion.Consultas) 'InfhotelHK.Consultas)
                            Case Else
                        End Select
                        End
                    End If
                '--------------------------------------------------------------
                End If
                End
                
           Case Is = 2 ' Usuario
                frmKeyBoard.txtResultado.Text = txtUsuario.Text
                frmKeyBoard.Show vbModal
                If wEnter Then
                   txtUsuario.Text = sDescrip
                End If
                wEnter = False
                
           Case Is = 3 ' Password
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   'txtPassword.SetFocus
                   txtPassword.Text = sDescrip
                  'cmdOpcion_Click (0)
                End If
                wEnter = False
                
           Case Is = 4 'HUELLA
                wenterHuellaSup = False
                lUsuarioHuella = True
                frmVerificacionHuellaSup.Opcion lModulo
                frmVerificacionHuellaSup.Show vbModal
                If wenterHuellaSup Then
                    wEnter = True
                    lUsuarioHuella = False
                    sUsuario = sVar1
                    Unload Me
                End If
           'HUELLA
           Case Is = 5
                wEnterHuella = False
                frmVerificarHuellaMarcacion.Show vbModal
                If wEnterHuella Then
                    'timSalida.Enabled = False
                    Exit Sub
                End If
                
                
    End Select
End Sub







Private Sub Timer1_Timer()
On Error GoTo fin

    Call ValidaVersiones
fin:
End Sub
Private Sub ValidaVersiones()
On Error GoTo fin
    
'     Dim CnSeg As New Connection
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

    If CnSeg.State = 0 Then
        CnSeg.Open
    End If
    Dim VersionInfoRest, VersionFEInfoHotel, VersionActualFEInfoHotel As String
    
    'VersionInfoRest = Calcular("select top 1 VersionInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)
    Open App.path & "\version.txt" For Input As #1   ' Abre el archivo para recibir los datos.
    Line Input #1, VersionInfoRest
    Close #1
    
    Dim sVersionExe As String
    sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
    
'    If VersionFEInfoHotel <> VersionActualFEInfoHotel Then
'        Label2(2).Caption = "Motor FE " & VersionFEInfoHotel & " Click Aqui"
'        Label2(2).Visible = True
'        Frame4.Visible = True
'    Else
'        If (VersionInfoHotel <> sVersionExe) Then
'            Label2(1).Visible = True
'            Label2(1).Caption = "InfoHotel " & VersionInfoHotel & " Click Aqui"
'            Frame4.Visible = True
'        End If
'    End If
    If VersionInfoRest <> "" Then
        If (VersionInfoRest <> sVersionExe) Then
            LblVer.Visible = True
            LblVer.Caption = "InfoRest " & VersionInfoRest & " "
            Frame4.Visible = True
        End If
    End If
Exit Sub
fin:
End Sub
Private Sub txtPassword_GotFocus()
   If Trim(txtPassword.Text) <> "" Then
      imgOpcion_Click (0)
   End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      imgOpcion_Click (0)
   End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      imgOpcion_Click (0)
   End If
End Sub

Private Sub txtUsuario_LostFocus()
   txtUsuario.Text = UCase(txtUsuario)
End Sub

Public Sub AccesoInicio()
    wEnter = False
    txtCaja.Caption = "CAJA  " & sCaja
    txtBD.Caption = UCase(sRuta) & " : " & UCase(sMDB)
    i = 1
    
    'HUELLA
    pTipo = "M"
    
    If sModulo = "INFOREST" Then
      ' Set RsUsuario = Lib.OpenRecordset("select * from vGrupoUSUARIO where lActivo=1 and ActivoGrupo=1 and lModulo01=1", Cn)
       lModulo = "01"
    ElseIf sModulo = "ADMINISTRACION" Then
       'Set RsUsuario = Lib.OpenRecordset("select * from vGrupoUSUARIO where lActivo=1 and ActivoGrupo=1 and lModulo02=1", Cn)
       lModulo = "02"
    Else
      ' Set RsUsuario = Lib.OpenRecordset("select * from vGrupoUSUARIO where lActivo=1 and ActivoGrupo=1 and lModulo03=1", Cn)
       lModulo = "03"
    End If
    Set RsUsuario = Lib.OpenRecordset("usp_Inforest_ObtieneUsuarios '" & sModulo & "'", Cn)
    If RsUsuario.RecordCount = 0 Then
       MsgBox "No existen Usuarios..!", vbCritical, sMensaje
       End
    End If
End Sub

Private Sub ValidacionEntrarConLicencia()
    If lHARDkey Then
        '----------Verifica Llave HK----------------------------------
        Dim verif As Boolean
        verif = hk.VerificaConexion
                        
        If verif = False Then
            Dim str As String
            str = ""
            Select Case sModulo
                Case "INFOREST"
                    str = hk.IniciaConexion(Aplicacion.PuntoVenta)
                Case "ADMINISTRACION"
                    str = hk.IniciaConexion(Aplicacion.Administracion)
                Case "CONSULTA"
                    str = hk.IniciaConexion(Aplicacion.Consultas)
                Case Else
                    
            End Select
            If str <> "" Then
                MsgBox str, vbCritical, "Aviso"
                End
            End If
        End If
        '--------------------------------------------------------------
    End If
End Sub
Public Sub Actualiza()
   Dim RsTemp As Recordset
   Screen.MousePointer = vbHourglass
   CnAlmacen.Execute "sp_ActualizaReceta"
   
   Cn.Execute "usp_Inforest_InicializaCostos"
    
    Dim oComandox As clsComando
    Set oComandox = New clsComando
    If Not oComandox.CreateCmdSp("usp_Inforest_ActualizaCostos", Cn) Then
       Set oComandox = Nothing
       Exit Sub
    End If
    oComandox.CreateParameter "@tNombreInforest", adVarChar, adParamInput, 50, sMDB
    If BDLink = "1" Then
        oComandox.CreateParameter "@tNombreAlmacen", adVarChar, adParamInput, 50, "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" 'CASO CORP.FERNANDEZ
    Else
        oComandox.CreateParameter "@tNombreAlmacen", adVarChar, adParamInput, 50, sAlmacenMDB   'CASO CORP.FERNANDEZ
    End If
    oComandox.CreateParameter "@tLocal", adVarChar, adParamInput, 5, sLocal
    If Not oComandox.GetParamOK Then
       Set oComandox = Nothing
       Exit Sub
    End If
    If Not oComandox.ExecSP Then
    Set oComandox = Nothing
    Exit Sub
    End If

    'Actualiza los precios de Venta de Transferencia a almacen

    If Not oComandox.CreateCmdSp("Usp_ActualizarPreciosTransferenciaAlmacen", Cn) Then
       Set oComandox = Nothing
       Exit Sub
    End If
    oComandox.CreateParameter "@SubGrupo", adVarChar, adParamInput, 50, ""
    If BDLink = "1" Then
        oComandox.CreateParameter "@BaseDatoAlmacen", adVarChar, adParamInput, 50, "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" 'CASO CORP.FERNANDEZ
    Else
        oComandox.CreateParameter "@BaseDatoAlmacen", adVarChar, adParamInput, 50, sAlmacenMDB   'CASO CORP.FERNANDEZ
    End If
    oComandox.CreateParameter "@tipooper", adInteger, adParamInput, 5, 2
    If Not oComandox.GetParamOK Then
       Set oComandox = Nothing
       Exit Sub
    End If
    If Not oComandox.ExecSP Then
        Set oComandox = Nothing
        Exit Sub
    End If

   Screen.MousePointer = vbDefault
   Exit Sub
End Sub
Public Sub Integra()

    'frmServidores.cargaModo False
    'frmServidores.llenaGrid
    
End Sub

Public Sub RTipoCambio()

   'TIPO DE CAMBIO
   Dim RsTipoCambio As Recordset
   
   Isql = "select * from TTIPOCAMBIO WHERE CONVERT(NVARCHAR,fFecha,103)= '" & FechaServidorTipoCambio() & "' "
   Set RsTipoCambio = Lib.OpenRecordset(Isql, Cn)
     
   If RsTipoCambio.RecordCount = 0 Then
        Dim oComando As clsComando
        Set oComando = New clsComando
        If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
           Set oComando = Nothing
           Exit Sub
        End If
    
        oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, 1
        oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, ""
        oComando.CreateParameter "@nTco", adDouble, adParamInput, 0, 0
        oComando.CreateParameter "@nTc2", adDouble, adParamInput, 0, 1
        oComando.CreateParameter "@nTc3", adDouble, adParamInput, 0, 1
        If Not oComando.GetParamOK Then
           Set oComando = Nothing
           Exit Sub
        End If
        If Not oComando.ExecSP Then
           Set oComando = Nothing
           Exit Sub
        End If
   End If
End Sub

Public Function Aplicar_Transparencia(ByVal hwnd As Long, _
                                      Valor As Integer) As Long
  
Dim Msg As Long
  
On Error Resume Next
  
If Valor < 0 Or Valor > 255 Then
   Aplicar_Transparencia = 1
Else
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   Msg = Msg Or WS_EX_LAYERED
     
   SetWindowLong hwnd, GWL_EXSTYLE, Msg
     
   SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
  
   Aplicar_Transparencia = 0
  
End If
  
  
If Err Then
   Aplicar_Transparencia = 2
End If
  
End Function

