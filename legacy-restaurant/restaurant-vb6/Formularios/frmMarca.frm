VERSION 5.00
Begin VB.Form frmMarca 
   BackColor       =   &H00E2DEDB&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5520
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido(a):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5985
   End
   Begin VB.Label lblNom 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
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
      ForeColor       =   &H0073B43C&
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   5985
   End
   Begin VB.Label lblRestultado 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Marcación Registrada Correctamente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   5985
   End
End
Attribute VB_Name = "frmMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public resultado As String
Public descripcion As String
Private Sub Form_Load()

 lblRestultado.Caption = resultado
 lblNom.Caption = descripcion 'RsMozo!descripcion
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Unload Me
End Sub
