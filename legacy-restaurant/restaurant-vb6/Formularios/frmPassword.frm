VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00D8D3CF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   4470
   ClientLeft      =   690
   ClientTop       =   2685
   ClientWidth     =   11265
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   10110
      Picture         =   "frmPassword.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3540
      Width           =   1125
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   55
      Left            =   9300
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2775
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   45
      Left            =   7770
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2775
      Width           =   1525
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   54
      Left            =   9300
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   480
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   53
      Left            =   8535
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   52
      Left            =   7770
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   480
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   51
      Left            =   9300
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1245
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   50
      Left            =   8535
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1245
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   49
      Left            =   7770
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1245
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   48
      Left            =   9300
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2010
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   47
      Left            =   8535
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2010
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   46
      Left            =   7770
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2010
      Width           =   765
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "BackSpace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   42
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1245
      Width           =   1125
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H007EC642&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Index           =   43
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2010
      Width           =   1125
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   41
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1125
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "Space"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   44
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3540
      Width           =   10080
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   765
   End
   Begin VB.TextBox txtResultado 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   0
      Width           =   11205
   End
   Begin VB.TextBox txtTempo 
      Height          =   285
      Left            =   10470
      TabIndex        =   0
      Top             =   2580
      Width           =   765
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sTipoSeguridad As String

Private Sub cmdkey_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    Select Case Index
           Case Is = 41
                wEnter = False
                Unload Me
           Case Is = 42
                If Len(Trim(txtResultado.Text)) > 0 Then
                   txtResultado.Text = Mid(txtResultado.Text, 1, Len(txtResultado.Text) - 1)
                Else
                   txtResultado.Text = ""
                End If
                txtTempo.SetFocus
           Case Is = 43
                wEnter = True
                sDescrip = txtResultado.Text
                Unload Me
           Case Is = 44
                txtResultado.Text = txtResultado.Text + " "
    Case Else
         If cmdKey(Index).Caption = "&&" Then
            txtResultado.Text = txtResultado.Text + "&"
         Else
            txtResultado.Text = txtResultado.Text + cmdKey(Index).Caption
         End If
         txtTempo.SetFocus
    End Select
End Sub

'HUELLA
Public Sub TipoSeguridad(ByVal vNew As String)
   sTipoSeguridad = vNew
End Sub

Private Sub cmdOpcion_Click()
        wenterHuellaSup = False
        lVieneHuella = False
        frmVerificacionHuellaSup.Opcion sTipoSeguridad
        frmVerificacionHuellaSup.Show vbModal
        If wenterHuellaSup = False Then
            Exit Sub
        Else
            Unload Me
        End If
End Sub
'----------------------

Private Sub Form_Load()
    wEnterHuella = False
    lVieneHuella = False
    wenterHuellaSup = False
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   wEnter = False
   MatrizBotones 40, 10, 1, 1, cmdKey
   Asigna
End Sub

Public Sub Asigna()
   cmdKey(1).Caption = "!"
   cmdKey(2).Caption = "'"
   cmdKey(3).Caption = "#"
   cmdKey(4).Caption = "$"
   cmdKey(5).Caption = "%"
   cmdKey(6).Caption = "&&"
   cmdKey(7).Caption = "/"
   cmdKey(8).Caption = "("
   cmdKey(9).Caption = ")"
   cmdKey(10).Caption = "="
   cmdKey(11).Caption = "Q"
   cmdKey(12).Caption = "W"
   cmdKey(13).Caption = "E"
   cmdKey(14).Caption = "R"
   cmdKey(15).Caption = "T"
   cmdKey(16).Caption = "Y"
   cmdKey(17).Caption = "U"
   cmdKey(18).Caption = "I"
   cmdKey(19).Caption = "O"
   cmdKey(20).Caption = "P"
   cmdKey(21).Caption = "A"
   cmdKey(22).Caption = "S"
   cmdKey(23).Caption = "D"
   cmdKey(24).Caption = "F"
   cmdKey(25).Caption = "G"
   cmdKey(26).Caption = "H"
   cmdKey(27).Caption = "J"
   cmdKey(28).Caption = "K"
   cmdKey(29).Caption = "L"
   cmdKey(30).Caption = "Ñ"
   cmdKey(31).Caption = "Z"
   cmdKey(32).Caption = "X"
   cmdKey(33).Caption = "C"
   cmdKey(34).Caption = "V"
   cmdKey(35).Caption = "B"
   cmdKey(36).Caption = "N"
   cmdKey(37).Caption = "M"
   cmdKey(38).Caption = ","
   cmdKey(39).Caption = "."
   cmdKey(40).Caption = "-"
   cmdKey(41).Caption = "ESC"
   cmdKey(42).Caption = "BackSpace"
   cmdKey(43).Caption = "Enter"
   cmdKey(44).Caption = "SPACE"
   cmdKey(45).Caption = "0"
   cmdKey(46).Caption = "1"
   cmdKey(47).Caption = "2"
   cmdKey(48).Caption = "3"
   cmdKey(49).Caption = "4"
   cmdKey(50).Caption = "5"
   cmdKey(51).Caption = "6"
   cmdKey(52).Caption = "7"
   cmdKey(53).Caption = "8"
   cmdKey(54).Caption = "9"
   cmdKey(55).Caption = "."
End Sub

Private Sub Form_Paint()
On Error Resume Next
   txtTempo.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmPassword = Nothing
End Sub

Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

If Shift > 0 Then
      Select Case KeyCode
       Case 48
            txtResultado.Text = txtResultado.Text + "="
       Case 49
            txtResultado.Text = txtResultado.Text + "!"
       Case 50
            txtResultado.Text = txtResultado.Text + "'"
       Case 51
            txtResultado.Text = txtResultado.Text + "#"
       Case 52
            txtResultado.Text = txtResultado.Text + "$"
       Case 53
            txtResultado.Text = txtResultado.Text + "%"
       Case 54
            txtResultado.Text = txtResultado.Text + "&"
       Case 55
            txtResultado.Text = txtResultado.Text + "/"
       Case 56
            txtResultado.Text = txtResultado.Text + "("
       Case 57
            txtResultado.Text = txtResultado.Text + ")"
      End Select
   Else
      Select Case KeyCode
          Case 13
               Call cmdkey_Click(43)
          Case 27
               Call cmdkey_Click(41)
          Case 8
               Call cmdkey_Click(42)
          Case 32
               txtResultado.Text = txtResultado.Text + " "
          Case 65
               txtResultado.Text = txtResultado.Text + "A"
          Case 66
               txtResultado.Text = txtResultado.Text + "B"
          Case 67
               txtResultado.Text = txtResultado.Text + "C"
          Case 68
               txtResultado.Text = txtResultado.Text + "D"
          Case 69
               txtResultado.Text = txtResultado.Text + "E"
          Case 70
               txtResultado.Text = txtResultado.Text + "F"
          Case 71
               txtResultado.Text = txtResultado.Text + "G"
          Case 72
               txtResultado.Text = txtResultado.Text + "H"
          Case 73
               txtResultado.Text = txtResultado.Text + "I"
          Case 74
               txtResultado.Text = txtResultado.Text + "J"
          Case 75
               txtResultado.Text = txtResultado.Text + "K"
          Case 76
               txtResultado.Text = txtResultado.Text + "L"
          Case 77
               txtResultado.Text = txtResultado.Text + "M"
          Case 78
               txtResultado.Text = txtResultado.Text + "N"
          Case 79
               txtResultado.Text = txtResultado.Text + "O"
          Case 80
               txtResultado.Text = txtResultado.Text + "P"
          Case 81
               txtResultado.Text = txtResultado.Text + "Q"
          Case 82
               txtResultado.Text = txtResultado.Text + "R"
          Case 83
               txtResultado.Text = txtResultado.Text + "S"
          Case 84
               txtResultado.Text = txtResultado.Text + "T"
          Case 85
               txtResultado.Text = txtResultado.Text + "U"
          Case 86
               txtResultado.Text = txtResultado.Text + "V"
          Case 87
               txtResultado.Text = txtResultado.Text + "W"
          Case 88
               txtResultado.Text = txtResultado.Text + "X"
          Case 89
               txtResultado.Text = txtResultado.Text + "Y"
          Case 90
               txtResultado.Text = txtResultado.Text + "Z"
          Case 96, 48
               txtResultado.Text = txtResultado.Text + "0"
          Case 97, 49
               txtResultado.Text = txtResultado.Text + "1"
          Case 98, 50
               txtResultado.Text = txtResultado.Text + "2"
          Case 99, 51
               txtResultado.Text = txtResultado.Text + "3"
          Case 100, 52
               txtResultado.Text = txtResultado.Text + "4"
          Case 101, 53
               txtResultado.Text = txtResultado.Text + "5"
          Case 102, 54
               txtResultado.Text = txtResultado.Text + "6"
          Case 103, 55
               txtResultado.Text = txtResultado.Text + "7"
          Case 104, 56
               txtResultado.Text = txtResultado.Text + "8"
          Case 105, 57
               txtResultado.Text = txtResultado.Text + "9"
          Case 110, 190
               txtResultado.Text = txtResultado.Text + "."
          Case 106, 188
               txtResultado.Text = txtResultado.Text + ","
          Case 109, 189
               txtResultado.Text = txtResultado.Text + "-"
          Case 192
               txtResultado.Text = txtResultado.Text + "Ñ"
      End Select
    End If
End Sub


