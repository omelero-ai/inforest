VERSION 5.00
Begin VB.Form frmMarcar 
   Caption         =   "Marcar Facturas"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   Icon            =   "frmMarcar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1410
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   3465
      Picture         =   "frmMarcar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   855
      Width           =   1245
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   0
      Left            =   4755
      Picture         =   "frmMarcar.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   855
      Width           =   1245
   End
   Begin VB.TextBox txtDestinatario 
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   427
      Width           =   4830
   End
   Begin VB.TextBox txtDocumento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      TabIndex        =   0
      Top             =   67
      Width           =   4830
   End
   Begin VB.Label Label 
      Caption         =   "Destinatario :"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   450
      Width           =   960
   End
   Begin VB.Label Label 
      Caption         =   "Documento :"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "frmMarcar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 0 Then
      Cn.Execute "update ddocumento set tGuia='MARCADO' where tDocumento='" & sDescrip & "'"
      MsgBox "Factura marcarda", vbExclamation, sMensaje
   End If
   Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    sDescrip = ""
    Isql = "SELECT distinct dbo.MDOCUMENTO.tDocumento as Descripcion, dbo.MDOCUMENTO.fFecha, dbo.vCliente.Descripcion as Cliente, dbo.vCliente.tIdentidad, dbo.MDOCUMENTO.nVenta, vCliente.Codigo AS Codigo " & _
           "FROM   dbo.MDOCUMENTO INNER JOIN dbo.DDOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DDOCUMENTO.tDocumento LEFT OUTER JOIN dbo.vCliente ON dbo.MDOCUMENTO.tCodigoCliente = dbo.vCliente.Codigo " & _
           "WHERE (ISNULL(tCodigoCliente,'')<>'' or len(ltrim(tCodigoCliente)) > 0) and isnull(tguia,'')='' and tEstadoDocumento<>'04'"
           
    frmBusca.cboCriterio.Enabled = True
    frmBusca.nPredeterm = 0
    Call ConfGrilla(5, frmBusca.grdGrilla, "Documento", 2, "Descripcion", 1500, 0, 0, "", _
                                           "Fecha", 2, "fFecha", 1100, 0, 0, "dd/MM/yyyy", _
                                           "Monto", 2, "nVenta", 1000, 1, 0, "###,##0.00", _
                                           "Identificador", 2, "tIdentidad", 1100, 0, 0, "", _
                                           "Razón Comercial", 2, "Cliente", 2600, 0, 0, "")
    frmBusca.Show vbModal
    If Not wEnter Then
       cmdOpcion(0).Enabled = False
       Exit Sub
    End If
    sCliente = sCodigo
    txtDocumento.Text = sDescrip
    txtDestinatario.Text = Calcular("select Descripcion as codigo from vCliente where Codigo='" & sCliente & "'", Cn)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmMarcar = Nothing
End Sub
