VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2850
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Left            =   2295
      Top             =   90
   End
   Begin VB.Label lblMensaje 
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
      Left            =   585
      TabIndex        =   0
      Top             =   45
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmMsgBox.frx":0000
      Top             =   30
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Centrar Me
End Sub

Private Sub Timer_Timer()
   Unload Me
End Sub

Public Sub Centrar(frmNombre As Form)
    With frmNombre
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
    End With
End Sub

