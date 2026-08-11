VERSION 5.00
Begin VB.Form frmSolicitudBoleta 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSDatos 
      Caption         =   "SIN DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnCDatos 
      Caption         =   "CON DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "EMITIR BOLETA"
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
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frmSolicitudBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCDatos_Click()
    Me.Hide
    frmBusquedaRapida.Show vbModal
End Sub

Private Sub btnSDatos_Click()
    
    If sCodigoClienteGeneral = "" Or sCodigoClienteGeneral = "0" Then
        MsgBox "Codigo de cliente general no configurado en el INI.", vbInformation, sMensaje
        Exit Sub
    End If
    
    Me.Hide
    sCodigo = sCodigoClienteGeneral 'Calcular("select Codigo FROM vCliente where tIdentidad = '" & sCodigoClienteGeneral & "'", Cn)
    sDescrip = "Cliente General"

    frmBusquedaRapida.EjecutarDocumentacion sCodigo, sDescrip

End Sub
