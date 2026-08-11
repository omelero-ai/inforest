VERSION 5.00
Begin VB.Form frmPreciosDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios del Producto"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00E2DEDB&
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
      Left            =   1620
      Picture         =   "frmPreciosDetalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6120
      Width           =   1245
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00E2DEDB&
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
      Left            =   2925
      Picture         =   "frmPreciosDetalle.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Frame fraDetalle 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Detalle del Producto "
      ForeColor       =   &H00800080&
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   4155
      Begin VB.CommandButton cmdPrecio 
         BackColor       =   &H00F2FAED&
         Caption         =   "Precio"
         Height          =   585
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   300
         Width           =   1155
      End
      Begin VB.Frame fraImpuesto 
         BackColor       =   &H00E2DEDB&
         Caption         =   " Impuestos "
         ForeColor       =   &H00800080&
         Height          =   975
         Left            =   120
         TabIndex        =   31
         Top             =   1695
         Width           =   3960
         Begin VB.CommandButton cmdImpuesto 
            BackColor       =   &H00F2FAED&
            Height          =   585
            Index           =   2
            Left            =   2730
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   270
            Width           =   1155
         End
         Begin VB.CommandButton cmdImpuesto 
            BackColor       =   &H00F2FAED&
            Height          =   585
            Index           =   1
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   270
            Width           =   1155
         End
         Begin VB.CommandButton cmdImpuesto 
            BackColor       =   &H00F2FAED&
            Height          =   585
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label lblDesc 
            Height          =   615
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraPrecio 
         BackColor       =   &H00E2DEDB&
         Caption         =   " Precio de Venta "
         ForeColor       =   &H00800080&
         Height          =   3210
         Left            =   90
         TabIndex        =   6
         Top             =   2700
         Width           =   3990
         Begin VB.Label txtCantidad 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   30
            Top             =   2580
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   375
            TabIndex        =   29
            Top             =   2625
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recargo :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   390
            TabIndex        =   28
            Top             =   855
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio Venta :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   120
            TabIndex        =   27
            Top             =   2340
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   240
            TabIndex        =   26
            Top             =   585
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio Oficial :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   90
            TabIndex        =   25
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto 2 :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   240
            TabIndex        =   24
            Top             =   1725
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto 1 :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   240
            TabIndex        =   23
            Top             =   1455
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto 3 :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   240
            TabIndex        =   22
            Top             =   2010
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   18
            Left            =   600
            TabIndex        =   21
            Top             =   2925
            Width           =   390
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio Neto :"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   9
            Left            =   165
            TabIndex        =   20
            Top             =   1185
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   1590
            TabIndex        =   19
            Top             =   570
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   1590
            TabIndex        =   18
            Top             =   840
            Width           =   150
         End
         Begin VB.Label txtImpuesto1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   17
            Top             =   1410
            Width           =   1365
         End
         Begin VB.Label txtImpuesto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   16
            Top             =   1695
            Width           =   1365
         End
         Begin VB.Label txtImpuesto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   1065
            TabIndex        =   15
            Top             =   1965
            Width           =   1365
         End
         Begin VB.Label txtDPorcentaje 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1065
            TabIndex        =   14
            Top             =   540
            Width           =   510
         End
         Begin VB.Label txtRPorcentaje 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1065
            TabIndex        =   13
            Top             =   810
            Width           =   510
         End
         Begin VB.Label txtRImporte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1785
            TabIndex        =   12
            Top             =   810
            Width           =   645
         End
         Begin VB.Label txtDImporte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1785
            TabIndex        =   11
            Top             =   540
            Width           =   645
         End
         Begin VB.Label txtNeto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   1065
            TabIndex        =   10
            Top             =   1140
            Width           =   1365
         End
         Begin VB.Label txtPVenta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   9
            Top             =   2310
            Width           =   1365
         End
         Begin VB.Label txtOficial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   8
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label txtVenta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1065
            TabIndex        =   7
            Top             =   2850
            Width           =   1365
         End
      End
      Begin VB.Frame fraDescuento 
         BackColor       =   &H00E2DEDB&
         Caption         =   " Descuentos / Recargos Unitario"
         ForeColor       =   &H00800080&
         Height          =   1515
         Left            =   1440
         TabIndex        =   1
         Top             =   180
         Width           =   2640
         Begin VB.CommandButton cmdDescuento 
            BackColor       =   &H00F2FAED&
            Caption         =   "( % ) del Recargo"
            Height          =   555
            Index           =   3
            Left            =   1380
            MaskColor       =   &H00F2FAED&
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   870
            Width           =   1155
         End
         Begin VB.CommandButton cmdDescuento 
            BackColor       =   &H00F2FAED&
            Caption         =   "( % ) del Descuento"
            Height          =   555
            Index           =   1
            Left            =   1380
            MaskColor       =   &H00F2FAED&
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   270
            Width           =   1155
         End
         Begin VB.CommandButton cmdDescuento 
            BackColor       =   &H00F2FAED&
            Caption         =   "Monto del Recargo"
            Height          =   555
            Index           =   2
            Left            =   150
            MaskColor       =   &H00F2FAED&
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   870
            Width           =   1155
         End
         Begin VB.CommandButton cmdDescuento 
            BackColor       =   &H00F2FAED&
            Caption         =   "Monto del Descuento"
            Height          =   555
            Index           =   0
            Left            =   150
            MaskColor       =   &H00F2FAED&
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   135
         Left            =   720
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPreciosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsProducto As Recordset
Dim xCambio As String
'Variables Detalles
Dim nPVenta As Double
Dim nPBase As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nRecargo As Double
Dim nDescuento As Double
Dim nOficial As Double
Dim nCantidad As Double
Dim xPedido As String
Dim xItem As String
Dim xTipo As String

'OBTENER DATOS
Dim dxPBase As Double
Dim dxnImpuesto1 As Double
Dim dxnImpuesto2 As Double
Dim dxnImpuesto3 As Double
Dim dxPNVenta As Double
Dim dxPVenta As Double

'RESULTADO DE DATOS
Dim rslPBase As Double
Dim rslImpuesto1 As Double
Dim rslImpuesto2 As Double
Dim rslImpuesto3 As Double
Dim rslPNVenta As Double
Dim rslPVenta As Double
'SABER SI ESTA AFECTO A IGV
Dim xsaber As Recordset

Dim x_cambio_precio As Integer


Private Sub cmdOpcion_Click(Index As Integer)

   If Index = 0 Then
        If x_cambio_precio = 0 Then
             If CDbl(txtRImporte.Caption) + CDbl(txtDImporte.Caption) = 0 Then
                 MsgBox "El Descuento o  Recargo debe ser mayor a 0", vbCritical, sMensaje
                 Unload Me
            Else
                GrabaProducto
                wEnter = True
            End If
        Else
            GrabaProducto
            wEnter = True
        End If
       
         
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   Centrar Me
   wEnter = False
   xPedido = sCodigo
   xItem = sDescrip
   If sTipo = "CajaRapida" Then
      xTipo = sTipo

      Isql = "SELECT [" & frmCajaRapida.sDetalle & "].*, dbo.TPRODUCTO.tDetallado AS Producto, dbo.vCortesia.Descripcion AS Cortesia, dbo.TPRODUCTO.lDescuento AS lDescuento, CASE [" & frmCajaRapida.sDetalle & "].nDescuento WHEN 0 THEN 0 ELSE [" & frmCajaRapida.sDetalle & "].nDescuento * 100 / [" & frmCajaRapida.sDetalle & "].nPrecioOficial END AS Descuento, " & _
             "dbo.TPRODUCTO.lModificable AS lModificable, CONVERT(bit, ISNULL(DATALENGTH([" & frmCajaRapida.sDetalle & "].tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad, isnull(tAplicaExoneracion,0)tAplicaExoneracion, isnull(tAplicaInafectacion,0)tAplicaInafectacion " & _
             "FROM [" & frmCajaRapida.sDetalle & "] LEFT OUTER JOIN (SELECT tItem, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad FROM [" & frmCajaRapida.sProductoPropiedad & "] Group by tItem) T1 " & _
             "ON [" & frmCajaRapida.sDetalle & "].tItem = T1.tItem LEFT OUTER JOIN dbo.vCortesia ON [" & frmCajaRapida.sDetalle & "].tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN " & _
             "dbo.TPRODUCTO ON [" & frmCajaRapida.sDetalle & "].tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
            "  WHERE [" & frmCajaRapida.sDetalle & "].tEstadoItem ='N' and [" & frmCajaRapida.sDetalle & "].tItem ='" & xItem & "' " & _
             " ORDER BY [" & frmCajaRapida.sDetalle & "].tItem"
             
   Else
      Isql = "select * from vPedidoDetalle " & _
             "where tEstadoItem ='N' and tCodigoPedido ='" & xPedido & "' and tItem='" & xItem & "' " & _
             "ORDER BY tCodigoPedido, tItem"
   End If
   Set RsProducto = Lib.OpenRecordset(Isql, Cn)
   Label2.Caption = RsProducto!tCodigoProducto
   Impuesto
   AsignaProducto
   x_cambio_precio = 0
End Sub

Private Sub cmdImpuesto_Click(Index As Integer)
    Dim nTotalImp As Double
    Dim nImp1 As Double
    Dim nImp2 As Double
    Dim nImp3 As Double
    x_cambio_precio = 1
    nImp1 = IIf(nImpuesto1 = 0, 0, nPorcentaje1)
    nImp2 = IIf(nImpuesto2 = 0, 0, nPorcentaje2)
    nImp3 = IIf(nImpuesto3 = 0, 0, nPorcentaje3)
    
    Select Case Index
        Case 0
           nImp1 = IIf(nImpuesto1 = 0, nPorcentaje1, 0)
        Case 1
           nImp2 = IIf(nImpuesto2 = 0, nPorcentaje2, 0)
        Case 2
           nImp3 = IIf(nImpuesto3 = 0, nPorcentaje3, 0)
    End Select
    
    nTotalImp = nImp1 + nImp2 + imp3
    
    nPBase = nPVenta / (1 + (nTotalImp / 100))
          
    nImpuesto1 = nPBase * nImp1 / 100
    nImpuesto2 = nPBase * nImp2 / 100
    nImpuesto3 = nPBase * nImp3 / 100
       
    txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
    txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
    txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
    
      rslPBase = nPBase
      rslImpuesto1 = nImpuesto1
      rslImpuesto2 = nImpuesto2
      rslImpuesto3 = nImpuesto3
     
        'GCAA 02032021
    rslPNVenta = txtPVenta.Caption 'nPVenta
    rslPVenta = (rslPNVenta * nCantidad) '(nPVenta * nCantidad)
End Sub

Private Sub cmdPrecio_Click()
   If nDecimal > 2 Then
      sTipo = "Decimal4"
   Else
      sTipo = ""
   End If
   frmNumPad.Show vbModal
   If wEnter Then
      x_cambio_precio = 1
      nPVenta = Val(sDescrip)
      nOficial = nPVenta
      nDescuento = 0
      nRecargo = 0
      txtDPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtRPorcentaje.Caption = Format(0, "###,###,###,##0.00")
      txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
      txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
      Select Case pais
            Case "001" 'Bolivia
                Acumulado = 0
                Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = (Acumulado / 100)
                nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
            
            Case Else 'Peru, Ecuador
                Acumulado = 0
                Acumulado = IIf(txtImpuesto1.Caption <> 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(txtImpuesto2.Caption <> 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(txtImpuesto3.Caption <> 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = 1 + (Acumulado / 100)
                nImpuesto1 = IIf(txtImpuesto1.Caption <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(txtImpuesto2.Caption <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(txtImpuesto3.Caption <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        End Select
      txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
      txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
      txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
      txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
      txtOficial.Caption = Format(nOficial, "###,###,##0.00")
      txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
      txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
      
      rslPBase = nPBase
      rslImpuesto1 = nImpuesto1
      rslImpuesto2 = nImpuesto2
      rslImpuesto3 = nImpuesto3
         'GCAA 02032021
        rslPNVenta = txtPVenta.Caption 'nPVenta
        rslPVenta = (rslPNVenta * nCantidad) '(nPVenta * nCantidad)
      
   End If
End Sub


Private Sub cmdDescuento_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Dscto. Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nDescuento = Val(sDescrip)
                     If nDescuento = 0 Then
                        MsgBox "Monto no válido, debe ser mayor a 0", vbCritical, sMensaje
                        Exit Sub
                     End If
                     If nDescuento > Val(Me.txtOficial.Caption) Then
                        MsgBox "Monto no válido", vbCritical, sMensaje
                        nDescuento = 0
                        Exit Sub
                     End If
                     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                     CalculaPrecio1
                  End If
               End If
            x_cambio_precio = 0
          Case Is = 1 ' Dscto. Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                        If Val(sDescrip) < 100 And Val(sDescrip) > 0 Then
                            txtDPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                            nDescuento = nOficial * Val(sDescrip) / 100
                            txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
                            CalculaPrecio
                        Else
                            MsgBox "Porcentaje no válido, debe ser mayor a 0", vbCritical, sMensaje
                            Exit Sub
                        End If
                  End If
               End If
          x_cambio_precio = 0
          Case Is = 2 ' Recargo Monto
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nRecargo = Val(sDescrip)
                     If nRecargo = 0 Then
                        MsgBox "Monto no válido, debe ser mayor a 0", vbCritical, sMensaje
                        Exit Sub
                     End If
                     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                     CalculaPrecio1
                  End If
               End If
          x_cambio_precio = 0
          Case Is = 3 ' Recargo Porcentaje
               If nPBase > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     If Val(sDescrip) < 100 And Val(sDescrip) > 0 Then
                        txtRPorcentaje.Caption = Format(sDescrip, "###,###,###,##0.00")
                        nRecargo = nOficial * Val(sDescrip) / 100
                        txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
                        CalculaPrecio
                     Else
                        MsgBox "Porcentaje no válido, debe ser mayor a 0", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
               End If
          x_cambio_precio = 0
    End Select
End Sub

Public Sub CalculaPrecio()
    Acumulado = 0
    If nPVenta = 0 Then
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       nRecargo = 0
       nDescuento = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
      Select Case pais
            Case "001" 'Bolivia
                    nPVenta = nOficial - nDescuento + nRecargo
                    Acumulado = IIf(dxnImpuesto1 <> "0.00", Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(dxnImpuesto2 <> "0.00", Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(dxnImpuesto3 <> "0.00", Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    nImpuesto1 = IIf(dxnImpuesto1 <> "0.00", nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(dxnImpuesto2 <> "0.00", nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(dxnImpuesto3 <> "0.00", nPVenta * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
                    
            Case Else 'Peru, Ecuador
                    nPVenta = nOficial - nDescuento + nRecargo
                    Acumulado = IIf(dxnImpuesto1 <> "0.00", Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(dxnImpuesto2 <> "0.00", Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(dxnImpuesto3 <> "0.00", Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    nImpuesto1 = IIf(dxnImpuesto1 <> "0.00", nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(dxnImpuesto2 <> "0.00", nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(dxnImpuesto3 <> "0.00", nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
        
        End Select
    End If
    txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
    txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
    txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    
    txtNeto.Caption = Format(nPBase, "###,###,##0.00")
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
    
    rslPBase = nPBase
    rslImpuesto1 = nImpuesto1
    rslImpuesto2 = nImpuesto2
    rslImpuesto3 = nImpuesto3
    'GCAA 02032021
    rslPNVenta = txtPVenta.Caption 'nPVenta
    rslPVenta = (rslPNVenta * nCantidad) '(nPVenta * nCantidad)
    
End Sub


Public Sub CalculaPrecio1()
    Acumulado = 0
    If nPVenta = 0 Then
       txtDPorcentaje.Caption = "0.00"
       txtRPorcentaje.Caption = "0.00"
       nRecargo = 0
       nDescuento = 0
       nImpuesto1 = 0
       nImpuesto2 = 0
       nImpuesto3 = 0
    Else
      Select Case pais
            Case "001" 'Bolivia
                    nPVenta = txtVenta.Caption - nDescuento + nRecargo
                    Acumulado = IIf(dxnImpuesto1 <> "0.00", Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(dxnImpuesto2 <> "0.00", Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(dxnImpuesto3 <> "0.00", Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = (Acumulado / 100)
                    nImpuesto1 = IIf(dxnImpuesto1 <> "0.00", nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(dxnImpuesto2 <> "0.00", nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(dxnImpuesto3 <> "0.00", nPVenta * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / txtVenta.Caption, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / txtVenta.Caption, "###,###,###,##0.00")
            
            Case Else 'Peru, Ecuador
            
                    'GCAA N019-2021
                     'SI ES EXONERADO O INAFECTO
'                     sImpuesto1 = Calcular("select tImpuesto1 as codigo from TPARAMETRO", Cn)
'                    If RsProducto!nPrecioImpuesto1 = 0 Then
'                         dxnImpuesto1 = "0.00"
'                    ElseIf RsProducto!nPrecioImpuesto1 > 0 Or RsProducto!tAplicaExoneracion = True Then
'                         dxnImpuesto1 = "0.00"
'                    ElseIf RsProducto!nPrecioImpuesto1 > 0 Or RsProducto!tAplicaInafectacion = True Then
'                         dxnImpuesto1 = "0.00"
'                    End If
            
            
                    'nPVenta = txtVenta.Caption - nDescuento + nRecargo
                    'nPVenta = (RsProducto!nPrecioOficial * nCantidad) - nDescuento + nRecargo
                    nPVenta = nOficial - nDescuento + nRecargo
                    Acumulado = IIf(dxnImpuesto1 <> "0.00", Acumulado + nPorcentaje1, Acumulado)
                    Acumulado = IIf(dxnImpuesto2 <> "0.00", Acumulado + nPorcentaje2, Acumulado)
                    Acumulado = IIf(dxnImpuesto3 <> "0.00", Acumulado + nPorcentaje3, Acumulado)
                    Acumulado = 1 + (Acumulado / 100)
                    nImpuesto1 = IIf(dxnImpuesto1 <> "0.00", nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(dxnImpuesto2 <> "0.00", nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(dxnImpuesto3 <> "0.00", nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                    txtDPorcentaje.Caption = Format(nDescuento * 100 / txtVenta.Caption, "###,###,###,##0.00")
                    txtRPorcentaje.Caption = Format(nRecargo * 100 / txtVenta.Caption, "###,###,###,##0.00")
        
        End Select
    End If
    
    txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
    txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
    txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
    
    txtNeto.Caption = Format(nPBase, "###,###,##0.00")
    txtPVenta.Caption = Format(nPVenta, "###,###,##0.00")
    txtVenta.Caption = Format((nPVenta * nCantidad), "###,###,###,##0.00")
    
    rslPBase = nPBase
    rslImpuesto1 = nImpuesto1
    rslImpuesto2 = nImpuesto2
    rslImpuesto3 = nImpuesto3
    'GCAA 02032021
    rslPNVenta = txtPVenta.Caption 'nPVenta
    rslPVenta = (rslPNVenta * nCantidad) '(nPVenta * nCantidad)
    
    
'    txtImpuesto1.Caption = Format(nImpuesto1 / nCantidad, "###,###,###,##0.00")
'    txtImpuesto2.Caption = Format(nImpuesto2 / nCantidad, "###,###,###,##0.00")
'    txtImpuesto3.Caption = Format(nImpuesto3 / nCantidad, "###,###,###,##0.00")
'    txtNeto.Caption = Format(nPBase / nCantidad, "###,###,###,##0.00")
'    txtPVenta.Caption = Format(nPVenta / nCantidad, "###,###,##0.00")
'    txtVenta.Caption = Format((nPVenta), "###,###,###,##0.00")
'    nDescuento = nDescuento / nCantidad
'    nRecargo = nRecargo / nCantidad
'
'    rslPBase = nPBase / nCantidad
'    rslImpuesto1 = nImpuesto1 / nCantidad
'    rslImpuesto2 = nImpuesto2 / nCantidad
'    rslImpuesto3 = nImpuesto3 / nCantidad
'    'GCAA 02032021
'    rslPNVenta = txtPVenta.Caption 'nPVenta
'    rslPVenta = txtVenta.Caption 'rslPNVenta * nCantidad '(nPVenta * nCantidad)
End Sub



Public Sub Impuesto()
    'GCAA N019-2021
    'SI ES EXONERADO O INAFECTO
    sImpuesto1 = Calcular("select tImpuesto1 as codigo from TPARAMETRO", Cn)
    
   If RsProducto!nprecioImpuesto1 = 0 Then
        lblDesc.Caption = "Producto no esta afecto a, " & sImpuesto1
        sImpuesto1 = ""
   ElseIf (RsProducto!nprecioImpuesto1 > 0 Or RsProducto!nprecioImpuesto1 = 0) And RsProducto!tAplicaExoneracion = True Then
        sImpuesto1 = ""
        lblDesc.Caption = "Producto es Exonerado"
   ElseIf (RsProducto!nprecioImpuesto1 > 0 Or RsProducto!nprecioImpuesto1 = 0) And RsProducto!tAplicaInafectacion = True Then
        sImpuesto1 = ""
        lblDesc.Caption = "Producto es Inafecto"
   End If
        
        
        Label1(10).Caption = sImpuesto1 & " :"
        Label1(11).Caption = sImpuesto2 & " :"
        Label1(12).Caption = sImpuesto3 & " :"
        
        Label1(10).Visible = IIf(sImpuesto1 = "", False, True)
        Label1(11).Visible = IIf(sImpuesto2 = "", False, True)
        Label1(12).Visible = IIf(sImpuesto3 = "", False, True)
        
        txtImpuesto1.Visible = IIf(sImpuesto1 = "", False, True)
        txtImpuesto2.Visible = IIf(sImpuesto2 = "", False, True)
        txtImpuesto3.Visible = IIf(sImpuesto3 = "", False, True)
        
        cmdImpuesto(0).Caption = sImpuesto1
        cmdImpuesto(1).Caption = sImpuesto2
        cmdImpuesto(2).Caption = sImpuesto3
        
        cmdImpuesto(0).Visible = IIf(sImpuesto1 = "", False, True)
        cmdImpuesto(1).Visible = IIf(sImpuesto2 = "", False, True)
        cmdImpuesto(2).Visible = IIf(sImpuesto3 = "", False, True)
End Sub

Public Sub GrabaProducto()
   Screen.MousePointer = vbHourglass
   
   If xTipo = "CajaRapida" Then

      Isql = "Update " & frmCajaRapida.sDetalle & " Set nPrecioNeto = " & rslPBase & ", " & _
              "nDescuento = " & nDescuento & ", " & _
              "nRecargo = " & nRecargo & ", " & _
              "nPrecioOficial = " & nOficial & ", " & _
              "nprecioImpuesto1 = " & rslImpuesto1 & ", " & _
              "nprecioImpuesto2 = " & rslImpuesto2 & ", " & _
              "nprecioImpuesto3 = " & rslImpuesto3 & ", " & _
              "nPrecioVenta = " & rslPNVenta & ", " & _
              "nventa = " & rslPVenta & ", " & _
              "nCantidad = " & nCantidad & ", " & _
              "nImpuesto1 = " & (rslImpuesto1 * nCantidad) & ", " & _
              "nImpuesto2 = " & (rslImpuesto2 * nCantidad) & ", " & _
              "nImpuesto3 = " & (rslImpuesto3 * nCantidad) & " " & _
              "where tItem = '" & xItem & "'"
              
                  
       xCambio = "insert TLOG_MODPRECIO select 'CR" & xPedido & "',  '" & xItem & "', ndescuento, " & nDescuento & ", nrecargo, " & nRecargo & ", npreciooficial," & nOficial & ", " & _
                 " nprecioimpuesto1, " & rslImpuesto1 & ", nprecioimpuesto2," & rslImpuesto2 & ", nprecioimpuesto3," & rslImpuesto3 & ", nprecioventa, " & rslPNVenta & ", nventa, " & rslPVenta & ",  " & _
                 " ncantidad, " & nCantidad & ", nimpuesto1," & rslImpuesto1 * nCantidad & ", nimpuesto2, " & rslImpuesto2 * nCantidad & ",nimpuesto3, " & rslImpuesto3 * nCantidad & ", '" & tUsuarioAutoriza & "','" & sUsuario & "',GETDATE() from " & frmCajaRapida.sDetalle & " where tItem = '" & xItem & "'    "
   Else

      Isql = "Update DPEDIDO Set nPrecioNeto = " & rslPBase & ", " & _
              "nDescuento = " & nDescuento & ", " & _
              "nRecargo = " & nRecargo & ", " & _
              "nPrecioOficial = " & nOficial & ", " & _
              "nprecioImpuesto1 = " & rslImpuesto1 & ", " & _
              "nprecioImpuesto2 = " & rslImpuesto2 & ", " & _
              "nprecioImpuesto3 = " & rslImpuesto3 & ", " & _
              "nPrecioVenta = " & rslPNVenta & ", " & _
              "nventa = " & rslPVenta & ", " & _
              "nCantidad = " & nCantidad & ", " & _
              "nImpuesto1 = " & (rslImpuesto1 * nCantidad) & ", " & _
              "nImpuesto2 = " & (rslImpuesto2 * nCantidad) & ", " & _
              "nImpuesto3 = " & (rslImpuesto3 * nCantidad) & " " & _
              "where tItem = '" & xItem & "' and tCodigoPedido ='" & xPedido & "'"
        If sModulo = "INFOREST" Then
                    xCambio = "insert TLOG_MODPRECIO select '" & xPedido & "',  '" & xItem & "', ndescuento, " & nDescuento & ", nrecargo, " & nRecargo & ", npreciooficial," & nOficial & ", " & _
                             " nprecioimpuesto1, " & rslImpuesto1 & ", nprecioimpuesto2," & rslImpuesto2 & ", nprecioimpuesto3," & rslImpuesto3 & ", nprecioventa, " & rslPNVenta & ", nventa, " & rslPVenta & ",  " & _
                             " ncantidad, " & nCantidad & ", nimpuesto1," & rslImpuesto1 * nCantidad & ", nimpuesto2, " & rslImpuesto2 * nCantidad & ",nimpuesto3, " & rslImpuesto3 * nCantidad & ", '" & tUsuarioAutoriza & "','" & sUsuario & "', getdate() from dpedido where tItem = '" & xItem & "' and tCodigoPedido ='" & xPedido & "'  "
        Else
        
                    xCambio = "insert TLOG_MODPRECIO select '" & xPedido & "',  '" & xItem & "', ndescuento, " & nDescuento & ", nrecargo, " & nRecargo & ", npreciooficial," & nOficial & ", " & _
                             " nprecioimpuesto1, " & rslImpuesto1 & ", nprecioimpuesto2," & rslImpuesto2 & ", nprecioimpuesto3," & rslImpuesto3 & ", nprecioventa, " & rslPNVenta & ", nventa, " & rslPVenta & ",  " & _
                             " ncantidad, " & nCantidad & ", nimpuesto1," & rslImpuesto1 * nCantidad & ", nimpuesto2, " & rslImpuesto2 * nCantidad & ",nimpuesto3, " & rslImpuesto3 * nCantidad & ", '" & tUsuarioAutoriza & "','" & sMozo & "', getdate() from dpedido where tItem = '" & xItem & "' and tCodigoPedido ='" & xPedido & "'  "
        End If
   End If
   
    If xCambio <> "" Then
        On Error Resume Next
     Cn.Execute xCambio
    End If
   Cn.Execute Isql

   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaProducto()
     

     fraPrecio.Caption = IIf(IsNull(RsProducto!Producto), "", " " & RsProducto!Producto & " ")
     cmdPrecio.Enabled = IIf(IsNull(RsProducto!lModificable), False, RsProducto!lModificable)
     
     
    fraDescuento.Enabled = IIf(IsNull(RsProducto!lModificable), False, RsProducto!lModificable)
    
    
    
    
     nDescuento = IIf(IsNull(RsProducto!nDescuento), 0, RsProducto!nDescuento)
     nRecargo = IIf(IsNull(RsProducto!nRecargo), 0, RsProducto!nRecargo)
      
     nImpuesto1 = IIf(IsNull(RsProducto!nprecioImpuesto1), 0, RsProducto!nprecioImpuesto1)
     nImpuesto2 = IIf(IsNull(RsProducto!nprecioImpuesto2), 0, RsProducto!nprecioImpuesto2)
     nImpuesto3 = IIf(IsNull(RsProducto!nprecioImpuesto3), 0, RsProducto!nprecioImpuesto3)

     nOficial = IIf(IsNull(RsProducto!nPrecioOficial), 0, RsProducto!nPrecioOficial)
     nPBase = IIf(IsNull(RsProducto!nPrecioNeto), 0, RsProducto!nPrecioNeto)
     nPVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
     nCantidad = IIf(IsNull(RsProducto!nCantidad), 0, RsProducto!nCantidad)
                    
     txtOficial.Caption = Format(nOficial, "###,###,###,##0.00")
     txtNeto.Caption = Format(nPBase, "###,###,###,##0.00")
     txtDImporte.Caption = Format(nDescuento, "###,###,###,##0.00")
     txtRImporte.Caption = Format(nRecargo, "###,###,###,##0.00")
     txtImpuesto1.Caption = Format(nImpuesto1, "###,###,###,##0.00")
     txtImpuesto2.Caption = Format(nImpuesto2, "###,###,###,##0.00")
     txtImpuesto3.Caption = Format(nImpuesto3, "###,###,###,##0.00")
     txtPVenta.Caption = Format(nPVenta, "###,###,###,##0.00")
     txtCantidad.Caption = Format(nCantidad, "##,##0.00")
     txtVenta.Caption = Format(nPVenta * nCantidad, "###,###,###,##0.00")
    
     If nOficial = 0 Then
        txtDPorcentaje.Caption = "0.00"
        txtRPorcentaje.Caption = "0.00"
     Else
        txtDPorcentaje.Caption = Format(nDescuento * 100 / nOficial, "###,###,###,##0.00")
        txtRPorcentaje.Caption = Format(nRecargo * 100 / nOficial, "###,###,###,##0.00")
     End If
     
      dxPBase = nPBase
      dxnImpuesto1 = nImpuesto1
      dxnImpuesto2 = nImpuesto2
      dxnImpuesto3 = nImpuesto3
      dxPNVenta = nPVenta
      dxPVenta = (nPVenta * nCantidad)
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   sTipo = xTipo
   Set frmPreciosDetalle = Nothing
End Sub



