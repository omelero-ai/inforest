VERSION 5.00
Begin VB.Form frmKeyBoard 
   BackColor       =   &H00D8D3CF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard"
   ClientHeight    =   5415
   ClientLeft      =   690
   ClientTop       =   2685
   ClientWidth     =   11250
   Icon            =   "frmKeyboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00F2FAED&
      Caption         =   "Suprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8980
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4725
      Width           =   2230
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00F2FAED&
      Caption         =   "@OUTLOOK.COM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5980
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4725
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F2FAED&
      Caption         =   "@HOTMAIL.COM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2980
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4725
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F2FAED&
      Caption         =   "@GMAIL.COM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4725
      Width           =   2980
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
      Top             =   3255
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
      Top             =   3255
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
      Top             =   960
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
      Top             =   960
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
      Top             =   960
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
      Top             =   1725
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
      Top             =   1725
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
      Top             =   1725
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
      Top             =   2490
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
      Top             =   2490
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
      Top             =   2490
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
      Top             =   1725
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
      Top             =   2490
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
      Top             =   960
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
      Height          =   675
      Index           =   44
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4050
      Width           =   11235
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
      Top             =   960
      Width           =   765
   End
   Begin VB.TextBox txtResultado 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      IMEMode         =   3  'DISABLE
      Left            =   0
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11175
   End
   Begin VB.TextBox txtTempo 
      Height          =   285
      Left            =   10470
      TabIndex        =   0
      Top             =   3060
      Width           =   765
   End
End
Attribute VB_Name = "frmKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tipo As String
 
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
        If tipo = "Cliente" And pais <> "003" And Index <> 40 Then
            If (Index >= 2 And Index <= 10) Or (Index >= 38 And Index <= 44) Or (Index = 55) Then
                txtTempo.SetFocus
                Exit Sub
            End If
        Else
            If (Index >= 41 And Index <= 44) Or (Index = 55) Then
                txtTempo.SetFocus
                Exit Sub
            End If
        End If
         If cmdKey(Index).Caption = "&&" Then
            txtResultado.Text = txtResultado.Text + "&"
         Else
            txtResultado.Text = txtResultado.Text + cmdKey(Index).Caption
         End If
         txtTempo.SetFocus
    End Select
End Sub

Private Sub Command1_Click()
 txtResultado.Text = txtResultado.Text + Command1.Caption
End Sub

Private Sub Command2_Click()
 txtResultado.Text = txtResultado.Text + Command2.Caption
End Sub

Private Sub Command3_Click()
 txtResultado.Text = txtResultado.Text + Command3.Caption
End Sub

Private Sub Command4_Click()
txtResultado.Text = ""
End Sub

Private Sub txtresultado_Change()
On Error GoTo fin
    frmKeyBoard.Caption = "Keyboard        /  " & Len(txtResultado.Text) & " Caracteres"
    txtResultado.Text = Replace(txtResultado.Text, vbNewLine, " ")
    'Me.txtResultado.Text = Trim(txtResultado.Text)
    Me.txtResultado.SelStart = Len(Me.txtResultado)
fin:
    
End Sub
Private Sub Form_Load()
   Centrar Me
   wEnter = False
   MatrizBotones 40, 10, 1, 1, cmdKey
   Asigna
   If sModulo = "ADICION" Then
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
   
   End If
   
End Sub

Public Sub Asigna()
   cmdKey(1).Caption = "@"
   cmdKey(2).Caption = "_"
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
   txtTempo.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    tipo = ""
   Set frmKeyBoard = Nothing
End Sub

Private Sub txtResultado_KeyPress(KeyAscii As Integer)
    If tipo = "Cliente" And pais <> "003" Then
        If NadaSimbolos(KeyAscii) = False Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   
   If Shift > 0 And tipo <> "Cliente" Then
      Select Case KeyCode
       Case 48
            txtResultado.Text = txtResultado.Text + "="
       Case 49
            txtResultado.Text = txtResultado.Text + "@"
       Case 50
            txtResultado.Text = txtResultado.Text + "_"
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
                If tipo <> "Cliente" Then
                    txtResultado.Text = txtResultado.Text + "."
                End If
          Case 106, 188
            If tipo <> "Cliente" Then
               txtResultado.Text = txtResultado.Text + ","
            End If
          Case 109
            If tipo <> "Cliente" Then
               txtResultado.Text = txtResultado.Text + "-"
            End If
          Case 189
               txtResultado.Text = txtResultado.Text + "-"
          Case 192
               txtResultado.Text = txtResultado.Text + "Ñ"
      End Select

    End If
End Sub


Private Sub txtTempo_KeyPress(KeyAscii As Integer)
    If tipo = "Cliente" And pais <> "003" Then
        If NadaSimbolos(KeyAscii) = False Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub
