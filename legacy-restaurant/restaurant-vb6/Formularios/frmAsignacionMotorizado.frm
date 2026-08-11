VERSION 5.00
Begin VB.Form frmAsignacionMotorizado 
   Caption         =   "Asignación Motorizado"
   ClientHeight    =   855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblAsignacion 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Asignación Maxima a Motorizados S/. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmAsignacionMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim nAsignacion As String
    nAsignacion = Calcular("select nAsignacionMotorizado as Codigo FROM TPARAMETRO", Cn)
    lblAsignacion = nAsignacion
    
End Sub
