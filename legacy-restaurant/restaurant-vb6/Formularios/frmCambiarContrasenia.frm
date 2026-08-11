VERSION 5.00
Begin VB.Form frmCambiarContrasenia 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contraseña de Acceso"
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtActual 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00393034&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   765
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1810
      Width           =   3870
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   465
      Left            =   2640
      TabIndex        =   5
      Top             =   6720
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
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
      Height          =   465
      Left            =   840
      TabIndex        =   4
      Top             =   6720
      Width           =   1140
   End
   Begin VB.TextBox txtConfirmaNueva 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00393034&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   765
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3975
      Width           =   3870
   End
   Begin VB.TextBox txtNueva 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00393034&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   765
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2890
      Width           =   3870
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00393034&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   765
      TabIndex        =   0
      Top             =   740
      Width           =   3870
   End
   Begin VB.Image ImageTecla4 
      Height          =   540
      Left            =   4800
      Picture         =   "frmCambiarContrasenia.frx":0000
      Top             =   3960
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image ImageTecla3 
      Height          =   540
      Left            =   4800
      Picture         =   "frmCambiarContrasenia.frx":08AE
      Top             =   2880
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image ImageTecla2 
      Height          =   540
      Left            =   4800
      Picture         =   "frmCambiarContrasenia.frx":115C
      Top             =   1800
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image ImageTecla1 
      Height          =   540
      Left            =   4800
      Picture         =   "frmCambiarContrasenia.frx":1A0A
      Top             =   720
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image ImageCancelar 
      Height          =   735
      Left            =   4800
      Picture         =   "frmCambiarContrasenia.frx":22B8
      Top             =   5400
      Width           =   810
   End
   Begin VB.Image ImageGrabar 
      Height          =   795
      Left            =   1725
      Picture         =   "frmCambiarContrasenia.frx":2BA9
      Top             =   5400
      Width           =   2280
   End
   Begin VB.Image ImageNuevaClaveC 
      Height          =   855
      Left            =   600
      Picture         =   "frmCambiarContrasenia.frx":3A06
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Image ImageNuevaClave 
      Height          =   855
      Left            =   600
      Picture         =   "frmCambiarContrasenia.frx":430D
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Image ImageContraseña 
      Height          =   855
      Left            =   600
      Picture         =   "frmCambiarContrasenia.frx":4C14
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Image ImageUsuario 
      Height          =   855
      Left            =   600
      Picture         =   "frmCambiarContrasenia.frx":551B
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lblInformacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* Dato Necesario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   360
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Label lblConfirmalNecesario 
      BackStyle       =   0  'Transparent
      Caption         =   " *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   210
      TabIndex        =   13
      Top             =   4035
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblNuevaNecesario 
      BackStyle       =   0  'Transparent
      Caption         =   " *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   210
      TabIndex        =   12
      Top             =   2955
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblActualNecesario 
      BackStyle       =   0  'Transparent
      Caption         =   " *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   210
      TabIndex        =   11
      Top             =   1875
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblUsuarioNece 
      BackStyle       =   0  'Transparent
      Caption         =   " *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   210
      TabIndex        =   10
      Top             =   795
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRMAR CONTRASEÑA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   780
      TabIndex        =   9
      Top             =   3600
      Width           =   3645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NUEVA CONTRASEÑA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   780
      TabIndex        =   8
      Top             =   2520
      Width           =   2805
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA ACTUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   780
      TabIndex        =   7
      Top             =   1440
      Width           =   2685
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   780
      TabIndex        =   6
      Top             =   315
      Width           =   1365
   End
End
Attribute VB_Name = "frmCambiarContrasenia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTipo As String
Dim mLogin As String


Public Sub obtieneTipo(ByVal tIp As String)
    mTipo = tIp
End Sub

Public Sub obtieneLogin(ByVal usu As String)
    mLogin = usu
End Sub

Private Sub CmdAceptar_Click()
    ocultarlabel
    
    If txtUsuario.Text = "" Then: Me.lblUsuarioNece.Visible = True: Exit Sub
    If Me.txtActual.Text = "" Then: Me.lblActualNecesario.Visible = True: Exit Sub
    If Me.txtNueva.Text = "" Then: Me.lblNuevaNecesario.Visible = True: Exit Sub
    If Me.txtConfirmaNueva.Text = "" Then: Me.lblConfirmalNecesario.Visible = True: Exit Sub
    If UCase(Me.txtActual.Text) = UCase(Me.txtNueva.Text) Then: Me.lblInformacion.Caption = "La nueva contraseña debe ser distinta a la actual": Me.lblInformacion.Visible = True: Exit Sub
    If UCase(Me.txtNueva.Text) <> UCase(Me.txtConfirmaNueva) Then: Me.lblInformacion.Caption = "Nueva contraseña no coincide con la confirmación": Me.lblInformacion.Visible = True: Exit Sub

    If Calcular("select count(*) codigo from tusuario where  tUsuario='" & Trim(txtUsuario.Text) & "' and lActivo=1", Cn) = 0 Then
        Me.lblInformacion.Caption = "El Usuario no esta activo o no esta registrado"
        Me.lblInformacion.Visible = True
        Exit Sub
    End If
    If Calcular("select count(*) codigo from tusuario where  tUsuario='" & Trim(txtUsuario.Text) & "' and lActivo=1", Cn) > 1 Then
        Me.lblInformacion.Caption = "El Login corresponde a más de un Usuario"
        Me.lblInformacion.Visible = True
        Exit Sub
    End If
    
    If Calcular("select count(*) codigo from tusuario where  tUsuario='" & Trim(txtUsuario.Text) & "' and tpassword= '" & Encapsula(Trim(Me.txtActual.Text)) & "' and lActivo=1", Cn) = 0 Then
        Me.lblInformacion.Caption = "Contraseña Actual Incorrecta"
        Me.lblInformacion.Visible = True
        Exit Sub
    End If
    
     If Calcular("select count(tPassword) as codigo from tusuario where tpassword='" & Encapsula(txtNueva.Text) & "'", Cn) > 0 Then
            MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
            Me.txtConfirmaNueva.Text = ""
            Me.txtNueva.Text = ""
            Me.txtNueva.SetFocus
            Exit Sub
     End If
     
    Dim UCodigo As String
    UCodigo = Calcular("select tcodigousuario codigo from tusuario where tUsuario='" & Trim(txtUsuario.Text) & "' and tpassword= '" & Encapsula(Trim(Me.txtActual.Text)) & "' and lActivo=1", Cn)
    sPasa = UCodigo
    
    'Inserta Movimiento auditoria
    
    If lModAuditoria Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TUSUARIO", "USUARIO", "01", txtUsuario.Text, sPasa, "", "TCODIGOUSUARIO", "CODIGO USUARIO", sPasa, "tPassword", "Contraseña", Encapsula(Me.txtNueva), "tUsuarioModifica", "Usuario Modifica", Me.txtUsuario.Text)
                        If lAuditoria = False Then
                            Screen.MousePointer = vbDefault
                                Exit Sub
                        End If
    End If
    Cn.Execute "update tusuario set tpassword='" & Encapsula(Trim(Me.txtNueva.Text)) & "', fregistro=getdate(), tusuariomodifica='" & Trim(Me.txtUsuario.Text) & "' where TCODIGOUSUARIO='" & UCodigo & "' AND tresumido='" & Trim(Me.txtUsuario.Text) & "' and lactivo=1"
    
    Unload Me

End Sub

Private Sub ocultarlabel()
    Me.lblActualNecesario.Visible = False
    Me.lblConfirmalNecesario.Visible = False
    Me.lblNuevaNecesario.Visible = False
    Me.lblUsuarioNece.Visible = False
    Me.lblInformacion.Visible = False
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtUsuario.Text = mLogin
    Me.txtUsuario.SelStart = Len(Me.txtUsuario.Text)
End Sub


Private Sub ImageCancelar_Click()
    Unload Me
End Sub

Private Sub ImageGrabar_Click()
    ocultarlabel
    
    If txtUsuario.Text = "" Then: Me.lblUsuarioNece.Visible = True: Exit Sub
    If Me.txtActual.Text = "" Then: Me.lblActualNecesario.Visible = True: Exit Sub
    If Me.txtNueva.Text = "" Then: Me.lblNuevaNecesario.Visible = True: Exit Sub
    If Me.txtConfirmaNueva.Text = "" Then: Me.lblConfirmalNecesario.Visible = True: Exit Sub
    If UCase(Me.txtActual.Text) = UCase(Me.txtNueva.Text) Then: Me.lblInformacion.Caption = "La nueva contraseña debe ser distinta a la actual": Me.lblInformacion.Visible = True: Exit Sub
    If UCase(Me.txtNueva.Text) <> UCase(Me.txtConfirmaNueva) Then: Me.lblInformacion.Caption = "Nueva contraseña no coincide con la confirmación": Me.lblInformacion.Visible = True: Exit Sub

    If Calcular("select count(*) codigo from tusuario where  tResumido='" & Trim(txtUsuario.Text) & "' and lActivo=1", Cn) = 0 Then
        Me.lblInformacion.Caption = "El Usuario no esta activo o no esta registrado"
        Me.lblInformacion.Visible = True
        Exit Sub
    End If
    If Calcular("select count(*) codigo from tusuario where  tResumido='" & Trim(txtUsuario.Text) & "' and lActivo=1", Cn) > 1 Then
        Me.lblInformacion.Caption = "El Login corresponde a más de un Usuario"
        Me.lblInformacion.Visible = True
        Exit Sub
    End If
    
    If Calcular("select count(*) codigo from tusuario where  tResumido='" & Trim(txtUsuario.Text) & "' and tpassword= '" & Encapsula(Trim(Me.txtActual.Text)) & "' and lActivo=1", Cn) = 0 Then
        Me.lblInformacion.Caption = "Contraseña Actual Incorrecta"
        Me.lblInformacion.Visible = True
        Exit Sub
    End If
    
     If Calcular("select count(tPassword) as codigo from tusuario where tpassword='" & Encapsula(txtNueva.Text) & "'", Cn) > 0 Then
            MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
            Me.txtConfirmaNueva.Text = ""
            Me.txtNueva.Text = ""
            Me.txtNueva.SetFocus
            Exit Sub
     End If

    Dim UCodigo As String
    UCodigo = Calcular("select tcodigousuario codigo from tusuario where tResumido='" & Trim(txtUsuario.Text) & "' and tpassword= '" & Encapsula(Trim(Me.txtActual.Text)) & "' and lActivo=1", Cn)
    sPasa = UCodigo
    
    'Inserta Movimiento auditoria
    
    If lModAuditoria Then
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TUSUARIO", "USUARIO", "01", txtUsuario.Text, sPasa, "", "TCODIGOUSUARIO", "CODIGO USUARIO", sPasa, "tPassword", "Contraseña", Encapsula(Me.txtNueva), "tUsuarioModifica", "Usuario Modifica", Me.txtUsuario.Text)
                    
                        If lAuditoria = False Then
                            Screen.MousePointer = vbDefault
                                Exit Sub
                        End If
    End If
    On Error GoTo Err
    Debug.Print "update tusuario set tpassword='" & Encapsula(Trim(Me.txtNueva.Text)) & "' where TCODIGOUSUARIO='" & UCodigo & "' AND tResumido='" & Trim(Me.txtUsuario.Text) & "' and lactivo=1"
    Cn.Execute "update tusuario set tpassword='" & Encapsula(Trim(Me.txtNueva.Text)) & "' where TCODIGOUSUARIO='" & UCodigo & "' AND tResumido='" & Trim(Me.txtUsuario.Text) & "' and lactivo=1"
    
    Unload Me
    Exit Sub
Err:
    Screen.MousePointer = vbDefault
    Me.lblInformacion.Caption = "Ocurrió un error al grabar Información."
    Me.lblInformacion.Visible = True
    Exit Sub
End Sub


'Private Sub ImageTecla1_Click()
'        frmKeyBoard.Show vbModal
'        If wEnter Then
'           txtUsuario.Text = sDescrip
'        End If
'        wEnter = False
'End Sub
'
'Private Sub ImageTecla2_Click()
'        frmPassword.cmdOpcion.Visible = False
'        frmPassword.ImageHuella.Visible = False
'        frmPassword.Show vbModal
'        If wEnter Then
'           txtActual.Text = sDescrip
'        End If
'End Sub
'
'Private Sub ImageTecla3_Click()
'        frmPassword.cmdOpcion.Visible = False
'        frmPassword.ImageHuella.Visible = False
'        frmPassword.Show vbModal
'        If wEnter Then
'           txtNueva.Text = sDescrip
'        End If
'End Sub
'
'Private Sub ImageTecla4_Click()
'        frmPassword.cmdOpcion.Visible = False
'        frmPassword.ImageHuella.Visible = False
'        frmPassword.Show vbModal
'        If wEnter Then
'           txtConfirmaNueva.Text = sDescrip
'        End If
'End Sub





Private Sub txtActual_Change()
ocultarlabel
End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.txtNueva.SetFocus
    
    End If
End Sub

Private Sub txtConfirmaNueva_Change()
ocultarlabel
End Sub

Private Sub txtConfirmaNueva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    
    End If
End Sub

Private Sub txtNueva_Change()
ocultarlabel
End Sub

Private Sub txtNueva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       Me.txtConfirmaNueva.SetFocus
    
    End If
End Sub

Private Sub txtUsuario_Change()
ocultarlabel
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Me.txtActual.SetFocus
    
    End If
End Sub

Private Sub txtUsuario_LostFocus()
    txtUsuario.Text = UCase(txtUsuario.Text)
End Sub
