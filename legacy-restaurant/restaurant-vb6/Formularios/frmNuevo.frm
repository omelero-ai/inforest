VERSION 5.00
Begin VB.Form frmNuevo 
   Caption         =   "Datos de la Guía"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "frmNuevo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Tienda"
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
      Index           =   5
      Left            =   7230
      TabIndex        =   14
      Top             =   1830
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de la Guía "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   7095
      Begin VB.TextBox txtCorrelativo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   13
         Top             =   300
         Width           =   1485
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   765
         Width           =   5385
      End
      Begin VB.TextBox txtSerie 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1530
         TabIndex        =   0
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox txtReferencia 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "- "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2370
         TabIndex        =   11
         Top             =   383
         Width           =   135
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Número de Guía :"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   383
         Width           =   1275
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Index           =   1
         Left            =   855
         TabIndex        =   9
         Top             =   848
         Width           =   570
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tienda :"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   1313
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Correlativo"
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
      Index           =   3
      Left            =   7230
      TabIndex        =   1
      Top             =   675
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cliente"
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
      Index           =   4
      Left            =   7230
      TabIndex        =   2
      Top             =   1245
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Numero de Serie"
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
      Index           =   2
      Left            =   7230
      TabIndex        =   3
      Top             =   90
      Width           =   1275
   End
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
      Left            =   4530
      Picture         =   "frmNuevo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1830
      Width           =   1275
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
      Left            =   5850
      Picture         =   "frmNuevo.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1830
      Width           =   1275
   End
End
Attribute VB_Name = "frmNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCliente As Recordset
Dim sFlag As Boolean

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Aceptar
               If txtSerie.Text = "" Then
                  MsgBox "Número de Serie no asignado", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If txtCorrelativo.Text = "" Then
                  MsgBox "Número Correlativo no asignado", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If txtCliente.Text = "" Then
                  MsgBox "Cliente no asignado", vbExclamation, sMensaje
                  Exit Sub
               End If
                                                          
               'Consistencia Documento Repetido
               sTemp = "G-" & Trim(txtSerie.Text) & "-" & Trim(txtCorrelativo.Text)
               If sFlag Then
                  If Not Calcular("select tCodigoPedido as Codigo from MPEDIDO where tTipoPedido='04' and tEstadoPedido<>'03' and tObservacion ='" & sTemp & "'", Cn) = 0 Then
                     MsgBox "Guía de Remisión ya registrada", vbExclamation, sMensaje
                     Exit Sub
                  End If
                  '"select tCodigoPedido as Codigo from MPEDIDO where tTipoPedido='04' and tEstadoPedido<>'03' and tObservacion ='" & sTemp & "' and tClienteCorp = '" & sCliente & "'"
               Else
                  If Calcular("select count(tCodigoPedido) as Codigo from MPEDIDO where tTipoPedido='04' and tEstadoPedido<>'03' and tCodigoPedido <>'" & sPedido & " ' and tObservacion ='" & sTemp & "'", Cn) > 0 Then
                     MsgBox "Guía de Remisión ya registrada", vbExclamation, sMensaje
                     Exit Sub
                  End If
                  '"select count(tCodigoPedido) as Codigo from MPEDIDO where tTipoPedido='04' and tEstadoPedido<>'03' and tCodigoPedido <>'" & sPedido & " ' and tObservacion ='" & sTemp & "' and tClienteCorp = '" & sCliente & "'"
               End If
                             
               wEnter = True
               sCodigo = sCliente
               Unload Me
          
          Case Is = 1 ' Cancelar
               wEnter = False
               Unload Me
          
          Case Is = 2 'Numero de Serie
               sTipo = "Numero"
               frmNumPad.Show vbModal
               If wEnter Then
                  txtSerie.Text = Mid("00000", 1, 5 - Len(Trim(sDescrip))) & Trim(sDescrip)
               End If
               
          Case Is = 3 'Correlativo
               sTipo = "Numero"
               frmNumPad.Show vbModal
               If Len(Trim(sDescrip)) > 5 Then
                  MsgBox "Numero correlativo de Guia excede a lo permitido", vbExclamation, sMensaje
                  Exit Sub
               End If
               If wEnter Then
                  txtCorrelativo.Text = Mid("000000000", 1, 9 - Len(Trim(sDescrip))) & Trim(sDescrip)
               End If
               
          Case Is = 4 'Cliente
               sCodigo = ""
               sDescrip = ""
               sTemp = ""
                            
               Isql = "SELECT * from vCompania where lActivo = 1 Order by Descripcion"
               frmBusquedaRapida.nPredeterm = 2
               Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                               "Ruc", 2, "tIdentidad", 1600, 2, 0, "", _
                                                               "Cliente", 2, "Descripcion", 5500, 0, 0, "")
               frmBusquedaRapida.Show vbModal
               If wEnter = True And sCodigo <> "" Then
                  sCliente = sCodigo
                  txtCliente.Text = sDescrip
                  sTienda = ""
                  txtReferencia.Text = ""
               Else
                  Exit Sub
               End If
                              
          Case Is = 5 'Tienda
               sTemp = ""
               sCodigo = ""
               sDescrip = ""
               Isql = "SELECT tCodigoTienda as Codigo, tNombre as Descripcion, tDireccion from TTIENDA where lActivo = 1 and tCodigoCompania = '" & sCliente & "' Order by Descripcion"
               Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                               "Descripcion", 2, "Descripcion", 1600, 2, 0, "", _
                                                               "Direccion", 2, "tDireccion", 5500, 0, 0, "")
               
               frmBusquedaRapida.Show vbModal
               
               If wEnter = True And sCodigo <> "" Then
                  sTienda = sCodigo
                  txtReferencia.Text = sDescrip
               End If
               
   End Select
End Sub

Private Sub Form_Load()
   Centrar Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNuevo = Nothing
    Set RsCliente = Nothing
End Sub

