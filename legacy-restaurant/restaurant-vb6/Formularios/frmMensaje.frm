VERSION 5.00
Begin VB.Form frmMensaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard"
   ClientHeight    =   6210
   ClientLeft      =   690
   ClientTop       =   2685
   ClientWidth     =   11250
   Icon            =   "frmMensaje.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKey 
      Caption         =   "F12"
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
      Index           =   71
      Left            =   10350
      TabIndex        =   29
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F11"
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
      Index           =   70
      Left            =   9400
      TabIndex        =   28
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F1"
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
      Index           =   60
      Left            =   0
      TabIndex        =   27
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F2"
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
      Index           =   61
      Left            =   940
      TabIndex        =   26
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F3"
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
      Index           =   62
      Left            =   1880
      TabIndex        =   25
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F4"
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
      Index           =   63
      Left            =   2820
      TabIndex        =   24
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F5"
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
      Index           =   64
      Left            =   3760
      TabIndex        =   23
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F6"
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
      Index           =   65
      Left            =   4700
      TabIndex        =   22
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F7"
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
      Index           =   66
      Left            =   5640
      TabIndex        =   21
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F8"
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
      Index           =   67
      Left            =   6580
      TabIndex        =   20
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F9"
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
      Index           =   68
      Left            =   7520
      TabIndex        =   19
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "F10"
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
      Index           =   69
      Left            =   8460
      TabIndex        =   18
      Top             =   4725
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FFC0C0&
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
      BackColor       =   &H00FF8080&
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
      BackColor       =   &H00FF8080&
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
      BackColor       =   &H00FF8080&
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
      TabIndex        =   3
      Top             =   4020
      Width           =   11235
   End
   Begin VB.CommandButton cmdKey 
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
      TabIndex        =   2
      Top             =   960
      Width           =   765
   End
   Begin VB.TextBox txtResultado 
      Enabled         =   0   'False
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
      Locked          =   -1  'True
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
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   12
      Left            =   10350
      TabIndex        =   41
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   11
      Left            =   9400
      TabIndex        =   40
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   10
      Left            =   8460
      TabIndex        =   39
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   9
      Left            =   7520
      TabIndex        =   38
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   8
      Left            =   6580
      TabIndex        =   37
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   7
      Left            =   5640
      TabIndex        =   36
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   6
      Left            =   4700
      TabIndex        =   35
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   5
      Left            =   3760
      TabIndex        =   34
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   4
      Left            =   2820
      TabIndex        =   33
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   3
      Left            =   1880
      TabIndex        =   32
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   2
      Left            =   940
      TabIndex        =   31
      Top             =   5580
      Width           =   810
   End
   Begin VB.Label lblfuncion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   1
      Left            =   0
      TabIndex        =   30
      Top             =   5580
      Width           =   810
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMensaje As Recordset
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
           Case 60 To 71
                txtResultado.Text = txtResultado.Text + lblfuncion(Index - 59)
    Case Else
         If cmdKey(Index).Caption = "&&" Then
            txtResultado.Text = txtResultado.Text + "&"
         Else
            txtResultado.Text = txtResultado.Text + cmdKey(Index).Caption
         End If
         txtTempo.SetFocus
    End Select
End Sub

Private Sub Form_Load()
   Centrar Me
   wEnter = False
   MatrizBotones 40, 10, 1, 1, cmdKey
   Asigna
   Set RsMensaje = Lib.OpenRecordset("select * from TMENSAJE", Cn)
   AsignaMensaje
End Sub

Public Sub AsignaMensaje()
   lblfuncion(1).Caption = IIf(IsNull(RsMensaje!tF1), "", RsMensaje!tF1)
   lblfuncion(2).Caption = IIf(IsNull(RsMensaje!tF2), "", RsMensaje!tF2)
   lblfuncion(3).Caption = IIf(IsNull(RsMensaje!tF3), "", RsMensaje!tF3)
   lblfuncion(4).Caption = IIf(IsNull(RsMensaje!tF4), "", RsMensaje!tF4)
   lblfuncion(5).Caption = IIf(IsNull(RsMensaje!tF5), "", RsMensaje!tF5)
   lblfuncion(6).Caption = IIf(IsNull(RsMensaje!tF6), "", RsMensaje!tF6)
   lblfuncion(7).Caption = IIf(IsNull(RsMensaje!tF7), "", RsMensaje!tF7)
   lblfuncion(8).Caption = IIf(IsNull(RsMensaje!tF8), "", RsMensaje!tF8)
   lblfuncion(9).Caption = IIf(IsNull(RsMensaje!tF9), "", RsMensaje!tF9)
   lblfuncion(10).Caption = IIf(IsNull(RsMensaje!tF10), "", RsMensaje!tF10)
   lblfuncion(11).Caption = IIf(IsNull(RsMensaje!tF11), "", RsMensaje!tF11)
   lblfuncion(12).Caption = IIf(IsNull(RsMensaje!tF12), "", RsMensaje!tF12)
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
   Set frmKeyBoard = Nothing
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
       'Case Else
       '     MsgBox KeyCode
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
          Case 112
               txtResultado.Text = txtResultado.Text + lblfuncion(1).Caption
          Case 113
               txtResultado.Text = txtResultado.Text + lblfuncion(2).Caption
          Case 114
               txtResultado.Text = txtResultado.Text + lblfuncion(3).Caption
          Case 115
               txtResultado.Text = txtResultado.Text + lblfuncion(4).Caption
          Case 116
               txtResultado.Text = txtResultado.Text + lblfuncion(5).Caption
          Case 117
               txtResultado.Text = txtResultado.Text + lblfuncion(6).Caption
          Case 118
               txtResultado.Text = txtResultado.Text + lblfuncion(7).Caption
          Case 119
               txtResultado.Text = txtResultado.Text + lblfuncion(8).Caption
          Case 120
               txtResultado.Text = txtResultado.Text + lblfuncion(9).Caption
          Case 121
               txtResultado.Text = txtResultado.Text + lblfuncion(10).Caption
          Case 122
               txtResultado.Text = txtResultado.Text + lblfuncion(11).Caption
          Case 123
               txtResultado.Text = txtResultado.Text + lblfuncion(12).Caption
          Case Else
               MsgBox KeyCode
      End Select
    End If
End Sub
