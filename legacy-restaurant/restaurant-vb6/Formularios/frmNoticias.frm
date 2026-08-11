VERSION 5.00
Begin VB.Form frmAviso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificación"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton BtnAceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nota: Al Aceptar, saldra del formulario para la actualización de los Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Se realizo el Cambio de Precio, debido a la Programación realizada."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label LblHora 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label LblTitulo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAceptar_Click()
 wEnter = True
 Unload Me
End Sub

Private Sub BtnCancelar_Click()
 wEnter = False
 Unload Me
End Sub

Private Sub Form_Load()
   Centrar Me
 If PrograPre_Saber = "Principal" Then
    Label2.Visible = False
 Else
    Label2.Visible = True
 End If
 
 
 LblTitulo.Caption = PrograPre_Titulo
 LblFecha.Caption = PrograPre_Fecha
 LblHora.Caption = PrograPre_Hora
 
End Sub

Private Sub Label3_Click()

End Sub

