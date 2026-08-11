VERSION 5.00
Begin VB.Form frmFiltroRecibo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emite de Recibos"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frmFiltroRecibo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optTipoReporte 
      Caption         =   "Emite formato reporte de recibos de egreso por compra de mercaderia."
      Height          =   435
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   5775
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar"
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
      Index           =   2
      Left            =   3135
      Picture         =   "frmFiltroRecibo.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   450
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Emite"
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
      Index           =   0
      Left            =   225
      Picture         =   "frmFiltroRecibo.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   450
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Impresora"
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
      Index           =   1
      Left            =   1680
      Picture         =   "frmFiltroRecibo.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   450
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Salir"
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
      Index           =   3
      Left            =   4590
      Picture         =   "frmFiltroRecibo.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   450
      Width           =   1455
   End
End
Attribute VB_Name = "frmFiltroRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCanal As String
Dim xCriterio As String

Private Sub cmdOpcion_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   wEnter = True
   Select Case Index
          Case Is = 0 ' Pantalla
               sTipo = "Pantalla"
                If optTipoReporte(1).value Then
                     sTipo = "PantallaCompra"
                End If
          Case Is = 1 ' Impresora
               sTipo = "Impresora"
                If optTipoReporte(1).value Then
                     sTipo = "ImpresoraCompra"
                End If
          Case Is = 2 ' XLS
               sTipo = "Excel"
                If optTipoReporte(1).value Then
                     sTipo = "ExcelCompra"
                End If
          Case Is = 3 ' Salir
               Screen.MousePointer = vbDefault
               wEnter = False
               sTipo = "Salir"
   End Select

   Unload Me
End Sub

Private Sub Form_Load()
   Centrar Me
   If lAlmacen Then
        optTipoReporte(1).Enabled = True
   Else
        optTipoReporte(1).Enabled = False
   End If
   
End Sub


