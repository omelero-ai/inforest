VERSION 5.00
Begin VB.Form FrmActualizacion 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Información de Actualización"
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdActualizarMotor 
      Caption         =   "Actualizar Motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   2
      Top             =   1920
      Width           =   3400
   End
   Begin VB.CommandButton cmdActualizarInfoRest 
      Caption         =   "Actualizar InfoRest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   3400
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "FrmActualizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public tipo As String
Private Sub cmdActualizarInfoRest_Click()
On Error GoTo fin
    If moduloUso = "Consulta" Then
        Call IniciarActualizador("CONSULTAS")
        End
    End If
    
    If moduloUso = "PuntoVenta" Then
        Call IniciarActualizador("INFOREST")
        End
    End If
    
     If moduloUso = "Administracion" Then
        Call IniciarActualizador("ADMINISTRACION")
        End
    End If
    
    
    
Exit Sub
fin:
    MsgBox "Error: " & error
End Sub

Private Sub cmdActualizarMotor_Click()
On Error GoTo fin:
    Dim respuesta As String
    respuesta = EjecutaActualizadorInfoRest
    If respuesta <> "OK" Then
         MsgBox "Se ha encontrado un problema en la ejecucion: " & vbNewLine & respuesta & vbNewLine & "Favor de contactar con el administrador de sistemas y/o informar sobre una nueva actualización de Facturacion electronica que esta pendiente." & vbNewLine & "Estado: Muy Critico, favor de informar lo mas pronto posible!!!"
    Else
        MsgBox "La Actualizacion del motor ha sido satisfactorio!!!!"
        Me.Hide
    End If
Exit Sub
fin:
    MsgBox "error: " & error
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  If (tipo = "FE") Then
        cmdActualizarInfoRest.Visible = False
    End If
    If (tipo = "INFOREST") Then
        cmdActualizarMotor.Visible = False
    End If
End Sub

