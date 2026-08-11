VERSION 5.00
Object = "{B1C46850-3E6A-11D2-8FEB-00104B9E07A7}#3.0#0"; "SSDW3AO.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Begin VB.Form frmBuscaPedido 
   Caption         =   "Busca Pedidos"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdUltimo 
      Height          =   555
      Left            =   6120
      TabIndex        =   0
      Top             =   2790
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Ultimo Registro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":0000
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPrimero 
      Height          =   555
      Left            =   6120
      TabIndex        =   1
      Top             =   15
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Primer Registro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":08DA
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdAceptar 
      Height          =   555
      Left            =   6120
      TabIndex        =   2
      Top             =   3900
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":11B4
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdCancelar 
      Height          =   555
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   3345
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Cancela"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":14CE
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPrevio 
      Height          =   555
      Left            =   6120
      TabIndex        =   4
      Top             =   1125
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Previo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":1DA8
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdSiguiente 
      Height          =   555
      Left            =   6120
      TabIndex        =   5
      Top             =   1680
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Siguient"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":2682
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPaginaUp 
      Height          =   555
      Left            =   6120
      TabIndex        =   6
      Top             =   570
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Página Arriba"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":2F5C
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdPaginaDn 
      Height          =   555
      Left            =   6120
      TabIndex        =   7
      Top             =   2235
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Página Abajo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBuscaPedido.frx":3836
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdGrilla 
      Height          =   4470
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6045
      ScrollBars      =   2
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets.count =   1
      stylesets(0).Name=   "Cabecera"
      stylesets(0).ForeColor=   16711680
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmBuscaPedido.frx":4110
      stylesets(0).AlignmentText=   2
      HeadFont3D      =   3
      Font3D          =   1
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      RowNavigation   =   3
      HeadStyleSet    =   "Cabecera"
      ForeColorEven   =   0
      BackColorEven   =   -2147483624
      BackColorOdd    =   -2147483624
      RowHeight       =   423
      ExtraHeight     =   79
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   10663
      _ExtentY        =   7885
      _StockProps     =   79
      Caption         =   "Listado de Pedidos"
      ForeColor       =   0
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBuscaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsGrilla As Recordset
Public fAceptar As Boolean
Public sResultado As String

Private Sub cmdAceptar_AfterClick()
   fAceptar = True
   sResultado = RsGrilla!Codigo
   Unload Me
End Sub

Private Sub cmdCancelar_AfterClick(Index As Integer)
   fAceptar = False
   Unload Me
End Sub

Private Sub cmdPaginaDn_AfterClick()
   RsGrilla.Move 10
   If RsGrilla.EOF Then
      RsGrilla.MoveLast
   End If
End Sub

Private Sub cmdPaginaUp_AfterClick()
   RsGrilla.Move -10
   If RsGrilla.BOF Then
      RsGrilla.MoveFirst
   End If
End Sub

Private Sub cmdPrevio_AfterClick()
   RsGrilla.MovePrevious
   If RsGrilla.BOF Then
      RsGrilla.MoveFirst
   End If
End Sub

Private Sub cmdPrimero_AfterClick()
   RsGrilla.MoveFirst
End Sub

Private Sub cmdSiguiente_AfterClick()
   RsGrilla.MoveNext
   If RsGrilla.EOF Then
      RsGrilla.MoveLast
   End If
End Sub

Private Sub cmdUltimo_AfterClick()
   RsGrilla.MoveLast
End Sub

Private Sub Form_Load()
   Centrar Me
   fAceptar = False
   Isql = "select Codigo, TipoPedido, Mesa, Mozo, Motorizado from vPedidoCabecera where lDocumento = False "
   Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
   Call Lib.ConfGrid(5, grdGrilla, "Pedido", 2, "Codigo", 1000, 2, 0, "", _
                                   "Tipo", 2, "TipoPedido", 1000, 2, 0, "", _
                                   "Mesa", 2, "Mesa", 1000, 2, 0, "", _
                                   "Mozo", 2, "Mozo", 1000, 2, 0, "", _
                                   "Motorizado", 2, "Motorizado", 1000, 0, 0, "")
   Set grdGrilla.DataSource = RsGrilla
End Sub


