VERSION 5.00
Begin VB.Form frmTarjetaNumero 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Número de Tarjeta de Crédito"
   ClientHeight    =   3930
   ClientLeft      =   5760
   ClientTop       =   2670
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000011&
   Icon            =   "frmTarjetaNumero.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   13
      Left            =   1740
      TabIndex        =   15
      Top             =   3075
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   1725
      Index           =   12
      Left            =   2595
      TabIndex        =   14
      Top             =   2220
      Width           =   990
   End
   Begin VB.CommandButton cmdKey 
      Caption         =   "Sup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   2595
      TabIndex        =   13
      Top             =   1365
      Width           =   990
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   10
      Left            =   2595
      TabIndex        =   12
      Top             =   510
      Width           =   990
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   9
      Left            =   1740
      TabIndex        =   11
      Top             =   510
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   8
      Left            =   885
      TabIndex        =   10
      Top             =   510
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   7
      Left            =   30
      TabIndex        =   9
      Top             =   510
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   6
      Left            =   1740
      TabIndex        =   8
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   5
      Left            =   885
      TabIndex        =   7
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   4
      Left            =   30
      TabIndex        =   6
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   3
      Left            =   1740
      TabIndex        =   5
      Top             =   2220
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   2
      Left            =   885
      TabIndex        =   4
      Top             =   2220
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   2220
      Width           =   855
   End
   Begin VB.CommandButton cmdKey 
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
      Height          =   855
      Index           =   0
      Left            =   15
      TabIndex        =   2
      Top             =   3075
      Width           =   1725
   End
   Begin VB.TextBox txtTempo 
      Height          =   285
      Left            =   2820
      TabIndex        =   0
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label txtResultado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1234567890123456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   420
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   3510
   End
End
Attribute VB_Name = "frmTarjetaNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wPunto As Boolean
Dim sTemp As String

Private Sub cmdBorra_AfterClick()
   sTemp = "0"
   wPunto = False
   txtResultado.Caption = Format(sTemp, "################")
End Sub

Private Sub cmdEnter_AfterClick()
   wEnter = True
   Unload Me
End Sub

Private Sub cmdEsc_AfterClick()
   wEnter = False
   Unload Me
End Sub

Private Sub cmdkey_Click(Index As Integer)
    Select Case Index
           Case Is = 10 ' Esc
                wEnter = False
                Unload Me
                
           Case Is = 11 ' Supr
                wPunto = False
                txtResultado.Caption = Format(sTemp, "################")
                txtTempo.SetFocus
                                
           Case Is = 12 'Enter
                wEnter = True
                sDescrip = sTemp
                Unload Me
           
           Case Is = 13 'Punto
                If Not wPunto And sTipo <> "Numero" Then
                   sTemp = sTemp & "."
                   wPunto = True
                   txtTempo.SetFocus
                End If
           Case Else
                If (Not wPunto And Len(Trim(sTemp)) >= 16) Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 2 And sTipo = "") Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 3 And sTipo = "TC") Then
                   Beep
                   txtTempo.SetFocus
                Else
                   sTemp = IIf(sTemp = "0", cmdKey(Index).Caption, sTemp & cmdKey(Index).Caption)
                End If
                txtResultado.Caption = Format(sTemp, "################")
                txtTempo.SetFocus
    End Select

End Sub

Private Sub Form_Load()
'   Centrar Me
         
   wEnter = False
   wPunto = False
   sTemp = ""
   sDescrip = ""
   txtResultado.Caption = Format(sTemp, "################")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmNumPad = Nothing
End Sub

Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 13
        cmdkey_Click (12)
   Case 27
        cmdkey_Click (10)
   Case 46
        cmdkey_Click (11)
   Case 96, 48
        cmdkey_Click (0)
   Case 97, 49
        cmdkey_Click (1)
   Case 98, 50
        cmdkey_Click (2)
   Case 99, 51
        cmdkey_Click (3)
   Case 100, 52
        cmdkey_Click (4)
   Case 101, 53
        cmdkey_Click (5)
   Case 102, 54
        cmdkey_Click (6)
   Case 103, 55
        cmdkey_Click (7)
   Case 104, 56
        cmdkey_Click (8)
   Case 105, 57
        cmdkey_Click (9)
   Case 110, 190
        cmdkey_Click (13)
   End Select
End Sub
