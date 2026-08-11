VERSION 5.00
Begin VB.Form frmAdicion 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Acceso"
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   9075
      Left            =   0
      Picture         =   "frmAdicion.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5505
   End
End
Attribute VB_Name = "frmAdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
