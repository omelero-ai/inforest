VERSION 5.00
Begin VB.Form frmCrearIniTVS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion TVS"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnContinuar 
      Caption         =   "Continuar >>"
      Height          =   375
      Left            =   2175
      TabIndex        =   9
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Blue Vision]"
      Height          =   2130
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4050
      Begin VB.TextBox txtPuerto 
         Height          =   330
         Left            =   1620
         TabIndex        =   8
         Top             =   1605
         Width           =   975
      End
      Begin VB.TextBox txtPassword 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1635
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   765
         Width           =   2160
      End
      Begin VB.TextBox txtUsuario 
         Height          =   330
         Left            =   1635
         TabIndex        =   2
         Top             =   345
         Width           =   2175
      End
      Begin VB.TextBox txtServidor 
         Height          =   330
         Left            =   1635
         TabIndex        =   1
         Top             =   1185
         Width           =   2160
      End
      Begin VB.Label lblClearPassword 
         Caption         =   "Contraseña :"
         Height          =   270
         Left            =   180
         TabIndex        =   7
         Top             =   765
         Width           =   1650
      End
      Begin VB.Label lbllogin 
         Caption         =   "Usuario :"
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   405
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Servidor (IP) :"
         Height          =   270
         Left            =   180
         TabIndex        =   5
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Puerto :"
         Height          =   270
         Left            =   180
         TabIndex        =   4
         Top             =   1665
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmCrearIniTVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnContinuar_Click()
    Open App.Path + "\BlueVision.ini" For Append As #1
    Print #1, "[BlueVision]"
    Print #1, "login=" & Me.txtUsuario.Text
    Print #1, "ClearPassword=" & Me.txtPassword.Text
    Print #1, "url=" & "http://" & Me.txtServidor.Text & ":" & Me.txtPuerto.Text & "/"
    Close #1
    Unload Me
End Sub

