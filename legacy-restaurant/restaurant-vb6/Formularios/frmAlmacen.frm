VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlmacen 
   BackColor       =   &H80000004&
   Caption         =   "Descargos Almacén"
   ClientHeight    =   2250
   ClientLeft      =   5715
   ClientTop       =   6015
   ClientWidth     =   7245
   Icon            =   "frmAlmacen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   7245
   Begin VB.Frame fraOpcion 
      Caption         =   " Rango de Fechas a Descargar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   7185
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1530
         TabIndex        =   2
         Top             =   810
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60162049
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1530
         TabIndex        =   0
         Top             =   375
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   60162049
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   3510
         TabIndex        =   1
         Top             =   390
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   60162051
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   3510
         TabIndex        =   3
         Top             =   825
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   60162051
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   465
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   900
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Descargar"
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
      Left            =   4275
      Picture         =   "frmAlmacen.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1605
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
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
      Index           =   0
      Left            =   5760
      Picture         =   "frmAlmacen.frx":1B7C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1605
      Width           =   1455
   End
End
Attribute VB_Name = "frmAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCriterio       As String
Dim RsTransferencia As Recordset
Dim RsCierre        As Recordset
Dim RsTemporal      As Recordset
Dim xTotal          As Double
Dim lCierre         As Boolean
Dim fFechaCierre    As Date
Dim lFechaCierre    As Boolean
Dim clsDescAlmacen  As clsAlmacen

Private Sub cmdOpcion_Click()
    Dim cArea As String
    
   Screen.MousePointer = vbHourglass
   mdiAdministracion.StatusBar.Panels.Item(2).Text = "Iniciando Descargo..."
   xTotal = 0
   Dim Pedido As String
   sTemporal = dbTemporal(sCaja, 14, "tCodigoPedido", "nVarChar(10)", _
                                     "fFecha", "smalldatetime", _
                                     "Plato", "nVarChar(7)", _
                                     "nCantidad", "Float", _
                                     "tItem", "nVarChar(3)", _
                                     "tDescargo", "nVarChar(1)", _
                                     "tEnlace", "nVarChar(7)", _
                                     "tTipoPedido", "nVarChar(2)", _
                                     "tCodigoProducto", "nVarChar(7)", _
                                     "nRecetaCantidad", "Float", _
                                     "tSubAreaAlm", "nVarChar(3)", _
                                     "tSubAreaInf", "nVarChar(3)", _
                                     "lRecetaBase", "bit", _
                                     "lProducto", "bit")
                  
   '"tSubArea", "nVarChar(3)", _

   'Inserta el Cierre
   lCierre = Calcular("select lCierre as Codigo from tParametro", CnAlmacen)
   fFechaCierre = Calcular("select getdate() as Codigo", Cn)
   If lCierre Then
      If Calcular("select count(fCierre) as Codigo from mCierre where fCierre='" & Format(fFechaCierre, "YYYY/mm/dd") & "'", CnAlmacen) = 0 Then
         Isql = "SpInsmCierre '" & Format(fFechaCierre, "YYYY/mm/dd") & "', '" & sUsuario & "', 0"
         CnAlmacen.Execute Isql
      End If
      Isql = "SpLismCierre '" & Format(fFechaCierre, "YYYY/mm/dd") & "'"
      Set RsCierre = Lib.OpenRecordset(Isql, CnAlmacen)
      fFechaCierre = Format(dtpFecFin.value, "YYYY/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm:ss")
   End If
   
   'Ventas a descargar
   'Venta por Receta
   Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,''), " & sAlmacenMDB & ".dbo.TPRODUCTO.lRecetaBase, " & sAlmacenMDB & ".dbo.DRECETAVENTA.lProducto " & _
          "FROM (SELECT dbo.DPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.DPEDIDO.tCodigoProducto, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubAlmacen FROM dbo.TPRODUCTO RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
          " where (MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and isnull(lTransferido,0) = 0 and tEstadoItem='N' and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') T1 " & _
          "INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAVENTA ON " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta AND " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal " & _
          "INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto " & _
          "WHERE (" & sAlmacenMDB & ".dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '01' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal1 = 1) or (T1.tTipoPedido = '02' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '03' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal3= 1) or (T1.tTipoPedido = '04' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal4 = 1) or (T1.tTipoPedido = '05' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal5 = 1)) and " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal='" & sLocal & "' And " & sAlmacenMDB & ".dbo.MRECETAVENTA.lActivo = 1"
   Cn.Execute "insert into " & sTemporal & " " & Isql
   
   '(" & sAlmacenMDB & ".dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '01' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal1 = 1) or (T1.tTipoPedido = '02' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '03' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal3= 1))
   'Venta por Descargo Directo
   Isql = "SELECT dbo.DPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.DPEDIDO.tCodigoProducto as Plato, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTO.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTO.tEnlace as tCodigoProducto, 1 as nRecetaCantidad, dbo.vArea.tValor AS tSubArea, ISNULL(dbo.DPEDIDO.tSubAlmacen,''), 0, 0 " & _
          "FROM dbo.TPRODUCTO LEFT OUTER JOIN dbo.vArea ON dbo.TPRODUCTO.tArea = dbo.vArea.Codigo RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.TPRODUCTO.tCodigoProducto = dbo.DPEDIDO.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
          "where (MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and isnull(lTransferido,0) = 0 and tEstadoItem='N' and tDescargo = 'D' and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' And Isnull(tEnlace,'')<>''"
   Cn.Execute "insert into " & sTemporal & " " & Isql
   
   'Quita los Sin de las Ventas
   Cn.Execute "delete from " & sTemporal & " where tCodigoPedido + tItem + Plato + tCodigoProducto " & _
              "in (select dbo.TPRODUCTOPROPIEDAD.tCodigoPedido + TPRODUCTOPROPIEDAD.tItem + TPRODUCTOPROPIEDAD.tProducto + TPRODUCTOPROPIEDAD.tEnlace " & _
              "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN dbo.MPEDIDO ON dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.DPEDIDO ON dbo.TPRODUCTOPROPIEDAD.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.TPRODUCTOPROPIEDAD.tItem = dbo.DPEDIDO.tItem " & _
              " where tCodigoPropiedad='9999' and (dbo.MPEDIDO.tEstadoPedido = '02' or dbo.MPEDIDO.tEstadoPedido = '04' or dbo.MPEDIDO.tEstadoPedido = '05') and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' and isnull(lTransferido,0) = 0 and DPEDIDO.tEstadoItem='N')"
   
   'Propiedades con Receta
'   Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,''), 0, 0 " & _
'           "FROM (SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.TPRODUCTOPROPIEDAD.tProducto AS Plato, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubAlmacen " & _
'           "      FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.TPRODUCTO ON dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPRODUCTO.tCodigoProducto " & _
'           "      WHERE (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.tEstadoPedido = '02' or dbo.MPEDIDO.tEstadoPedido = '04' or dbo.MPEDIDO.tEstadoPedido = '05') AND (dbo.MPEDIDO.fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "')) T1 " & _
'           " INNER JOIN  " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD ON " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tLocal " & _
'           "WHERE (" & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.lDescargo = 1) and " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal='" & sLocal & "'"

   'DETP 12/04/2011 No se estaba Tomando en Cuenta en las recetas de propiedades los descargo por insumo
   Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,''), " & sAlmacenMDB & ".dbo.TPRODUCTO.lRecetaBase, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.lProducto " & _
           "FROM (SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.TPRODUCTOPROPIEDAD.tProducto AS Plato, dbo.DPEDIDO.nCantidad, dbo.DPEDIDO.tItem, dbo.TPRODUCTO.tDescargo, dbo.TPRODUCTOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubAlmacen " & _
           "      FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.TPRODUCTO ON dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPRODUCTO.tCodigoProducto " & _
           "      WHERE (ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0) AND (dbo.MPEDIDO.tEstadoPedido = '02' or dbo.MPEDIDO.tEstadoPedido = '04' or dbo.MPEDIDO.tEstadoPedido = '05') AND (dbo.MPEDIDO.fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "') AND (dbo.MPEDIDO.fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "')) T1 " & _
           " INNER JOIN  " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD ON " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tLocal " & _
           " INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto " & _
           "WHERE (" & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.lDescargo = 1) and " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal='" & sLocal & "'"
   Cn.Execute "insert into " & sTemporal & " " & Isql
      
   'Propiedades con Descargo Directo
   'dbo.vArea.tValor
   Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.DPEDIDO.tCodigoProducto AS Plato, dbo.DPEDIDO.nCantidad, dbo.TPRODUCTOPROPIEDAD.tItem, 'D' AS tDescargo, dbo.TPRODUCTOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TPRODUCTOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.tPropiedad.tArea AS tSubArea, ISNULL(dbo.DPEDIDO.tSubAlmacen,''), 0, 0 " & _
          "FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPROPIEDAD INNER JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TPRODUCTOPROPIEDAD.tProducto ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem INNER JOIN dbo.vArea ON dbo.TPROPIEDAD.tArea = dbo.vArea.Codigo " & _
          "where LEN(dbo.TPROPIEDAD.tEnlace)=7 and (MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and isnull(DPEDIDO.lTransferido,0)=0 and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "'"
   Cn.Execute "insert into " & sTemporal & " " & Isql
      
   'Combos por Recetas
   Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.tCodigoProducto AS Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAVENTA.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAVENTA.tSubArea, ISNULL(T1.tSubAlmacen,''), " & sAlmacenMDB & ".dbo.TPRODUCTO.lRecetaBase, " & sAlmacenMDB & ".dbo.DRECETAVENTA.lProducto " & _
          "FROM (SELECT dbo.DPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.CPEDIDO.tProductoCombo AS tCodigoProducto, dbo.CPEDIDO.nCantidad as nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubAlmacen FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
          "where (MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and isnull(lTransferido,0) = 0 and tEstadoItem='N' and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') T1 " & _
          "INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAVENTA ON " & sAlmacenMDB & ".dbo.MRECETAVENTA.tRecetaVenta = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tRecetaVenta AND " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tLocal " & _
          "INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto = " & sAlmacenMDB & ".dbo.DRECETAVENTA.tCodigoProducto " & _
          "WHERE (" & sAlmacenMDB & ".dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '01' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal1 = 1) or (T1.tTipoPedido = '02' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '03' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal3= 1) or (T1.tTipoPedido = '04' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal4 = 1) or (T1.tTipoPedido = '05' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal5 = 1)) and " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal='" & sLocal & "'"
   Cn.Execute "insert into " & sTemporal & " " & Isql
          '"WHERE (" & sAlmacenMDB & ".dbo.DRECETAVENTA.lDescargo = 1) and ((T1.tTipoPedido = '01' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal1 = 1) or (T1.tTipoPedido = '02' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal2= 1) or (T1.tTipoPedido = '03' AND " & sAlmacenMDB & ".dbo.DRECETAVENTA.lCanal3= 1)) and " & sAlmacenMDB & ".dbo.MRECETAVENTA.tLocal='" & sLocal & "'"
      
   'Combo por Descargo Directo
   Isql = "SELECT dbo.DPEDIDO.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad as nCantidad, dbo.DPEDIDO.tItem, TPRODUCTO_1.tDescargo, TPRODUCTO_1.tEnlace, dbo.MPEDIDO.tTipoPedido, TPRODUCTO_1.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.vArea.tValor AS tSubArea, ISNULL(dbo.DPEDIDO.tSubAlmacen,''), 0, 0 " & _
          "FROM dbo.CPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto INNER JOIN dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProductoCombo = TPRODUCTO_1.tCodigoProducto INNER JOIN dbo.vArea ON TPRODUCTO_1.tArea = dbo.vArea.Codigo RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
          "WHERE (dbo.MPEDIDO.tEstadoPedido = '02' OR dbo.MPEDIDO.tEstadoPedido = '04' OR dbo.MPEDIDO.tEstadoPedido = '05') AND ISNULL(dbo.DPEDIDO.lTransferido, 0) = 0 AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (TPRODUCTO_1.tDescargo = 'D') AND fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "'"
   Cn.Execute "insert into " & sTemporal & " " & Isql

   'Quita los Sin de los Combos
   Cn.Execute "delete from " & sTemporal & " where tcodigoPedido + tItem + Plato + tCodigoProducto " & _
              "in (select dbo.TCOMBOPROPIEDAD.tCodigoPedido + tItem + tProducto + tEnlace from dbo.TCOMBOPROPIEDAD INNER JOIN  dbo.MPEDIDO ON dbo.TCOMBOPROPIEDAD.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
              " where tCodigoPropiedad='9999' and (dbo.MPEDIDO.tEstadoPedido = '02' or dbo.MPEDIDO.tEstadoPedido = '04' or dbo.MPEDIDO.tEstadoPedido = '05') and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "')"
   
   'Propiedades de los Combos con Recetas
   Isql = "SELECT T1.tCodigoPedido, T1.fFecha, T1.Plato, T1.nCantidad, T1.tItem, T1.tDescargo, T1.tEnlace, T1.tTipoPedido, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tCodigoProducto, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.nCantidad AS nRecetaCantidad, " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tSubArea, ISNULL(T1.tSubAlmacen,''), 0, 0 " & _
          "FROM (SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad, dbo.TCOMBOPROPIEDAD.tItem, dbo.TPRODUCTO.tDescargo, dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.DPEDIDO.tSubAlmacen FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido INNER JOIN dbo.TPRODUCTO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBOPROPIEDAD.tProducto INNER JOIN dbo.CPEDIDO ON dbo.TCOMBOPROPIEDAD.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.TCOMBOPROPIEDAD.tItem = dbo.CPEDIDO.tItem AND dbo.TCOMBOPROPIEDAD.tItemCombo = dbo.CPEDIDO.tItemCombo ON dbo.DPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido And dbo.DPEDIDO.tItem = dbo.CPEDIDO.tItem " & _
          "  where (isnull(DPEDIDO.lTransferido,0)=0 and MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') T1 " & _
          "INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD ON T1.tEnlace = " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad INNER JOIN " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD ON " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tRecetaPropiedad = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tRecetaPropiedad AND " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal = " & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.tLocal " & _
          "WHERE (" & sAlmacenMDB & ".dbo.DRECETAPROPIEDAD.lDescargo = 1) and " & sAlmacenMDB & ".dbo.MRECETAPROPIEDAD.tLocal='" & sLocal & "'"
   Cn.Execute "insert into " & sTemporal & " " & Isql
      
   'Propiedades de los Combos con Descargo Directo
   Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.MPEDIDO.fFecha, dbo.CPEDIDO.tProductoCombo AS Plato, dbo.CPEDIDO.nCantidad, dbo.TCOMBOPROPIEDAD.tItem, 'D' as tDescargo, dbo.TCOMBOPROPIEDAD.tEnlace, dbo.MPEDIDO.tTipoPedido, dbo.TCOMBOPROPIEDAD.tEnlace AS tCodigoProducto, 1 AS nRecetaCantidad, dbo.TPROPIEDAD.tArea AS tSubArea, ISNULL(dbo.DPEDIDO.tSubAlmacen,''), 0, 0 " & _
          "FROM dbo.CPEDIDO INNER JOIN dbo.TCOMBOPROPIEDAD ON dbo.CPEDIDO.tCodigoPedido = dbo.TCOMBOPROPIEDAD.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.TCOMBOPROPIEDAD.tItem AND dbo.CPEDIDO.tItemCombo = dbo.TCOMBOPROPIEDAD.tItemCombo INNER JOIN dbo.TPROPIEDAD ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto INNER JOIN dbo.DPEDIDO ON dbo.CPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido AND dbo.CPEDIDO.tItem = dbo.DPEDIDO.tItem INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido " & _
          "where isnull(DPEDIDO.lTransferido,0)=0 and LEN(dbo.TPROPIEDAD.tEnlace)=7 and (MPEDIDO.tEstadoPedido = '02' or MPEDIDO.tEstadoPedido = '04' or MPEDIDO.tEstadoPedido = '05') and fFecha >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fFecha <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "'"
   Cn.Execute "insert into " & sTemporal & " " & Isql
   
   Isql = "SELECT " & sTemporal & ".tCodigoPedido, fFecha, Plato, nCantidad, tItem, tDescargo, tEnlace, tTipoPedido, " & sTemporal & ".tCodigoProducto, nRecetaCantidad, tSubAreaAlm, tSubAreaInf, nFactor, nPrecioPromedio, " & sTemporal & ".lRecetaBase, " & sTemporal & ".lProducto  FROM " & sTemporal & " INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO ON " & sTemporal & ".tCodigoProducto = " & sAlmacenMDB & ".dbo.TPRODUCTO.tCodigoProducto order by tCodigoPedido, tItem"
   
   Set RsTransferencia = Lib.OpenRecordset(Isql, Cn)
   If RsTransferencia.RecordCount > 0 Then
      RsTransferencia.MoveFirst
      xTotal = RsTransferencia.RecordCount
            
      Do While Not RsTransferencia.EOF
         Pedido = RsTransferencia!tCodigoPedido
         mdiAdministracion.StatusBar.Panels.Item(2).Text = "Transfiriendo Movimiento " & RsTransferencia.AbsolutePosition & " de " & xTotal
         
         If RsTransferencia!lRecetaBase And RsTransferencia!lProducto = False Then
            DescargoInsumoRecetaBase RsTransferencia!tCodigoPedido, RsTransferencia!tCodigoProducto, (RsTransferencia!nRecetaCantidad / RsTransferencia!nFactor) * RsTransferencia!nCantidad, IIf(IsNull(RsTransferencia!fFecha), Date, RsTransferencia!fFecha)
            'Numero de Pedido              Codigo de la Receta               (Cantidad de Receta / Factor)*La cantidad del Pedido= Nueva Cantidad
         Else
            'Linea Agregada por Luis Chumpitaz 08/02/2011
            If Len(Trim(RsTransferencia!tSubAreaInf)) = 0 Then
                cArea = RsTransferencia!tSubAreaAlm
            Else
                cArea = RsTransferencia!tSubAreaInf
            End If
            'Fin Linea Agregada 08/02/2011
            
            DescargoInsumo RsTransferencia!tCodigoPedido, cArea, RsTransferencia!tCodigoProducto, RsTransferencia!nCantidad, RsTransferencia!nRecetaCantidad, RsTransferencia!nFactor, RsTransferencia!tDescargo, IIf(IsNull(RsTransferencia!nPrecioPromedio), 0, RsTransferencia!nPrecioPromedio), IIf(IsNull(RsTransferencia!fFecha), Date, RsTransferencia!fFecha)  ' DETP Para que el preciopromedio no pase en Nulo 16/01/2009
         End If
         
         RsTransferencia.MoveNext
         If Not RsTransferencia.EOF Then
            If RsTransferencia!tCodigoPedido <> Pedido Then
               Cn.Execute "update DPEDIDO set lTransferido = 1 where tCodigoPedido = '" & Pedido & "'"
            End If
         End If
      Loop
      Screen.MousePointer = vbHourglass
      
      RsTransferencia.MoveLast
      Cn.Execute "update DPEDIDO set lTransferido = 1 where tCodigoPedido = '" & RsTransferencia!tCodigoPedido & "'"
      
      Screen.MousePointer = vbDefault
      MsgBox "Se realizó " & Trim(str(RsTransferencia.RecordCount)) & " Transferencia(s) con Exito", vbInformation, "Transferencia"
   Else
      Screen.MousePointer = vbDefault
      MsgBox "No existe Ventas nuevas a transferir", vbExclamation, "Transferencia"
   End If
        
   Set RsTransferencia = Nothing
   
   If lCierre Then
      Isql = "SpUpdmCierre '" & Format(RsCierre!fCierre, "YYYY/mm/dd") & "', '" & Format(RsCierre!fRegistro, "YYYY/mm/dd HH:mm:ss") & "', '" & RsCierre!tUsuario & "', 1"
      CnAlmacen.Execute Isql
      Set RsCierre = Nothing
   End If
   Cn.Execute "drop table " & sTemporal
   mdiAdministracion.StatusBar.Panels.Item(2).Text = "Caja : " & sCaja
End Sub

Private Sub cmdSalir_Click(Index As Integer)
   Unload Me
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   dtpFecIni.value = Date
   dtpFecFin.value = Date
   dtpFecIni.Enabled = True
   dtpFecFin.Enabled = True
   dtpHorIni.Enabled = True
   dtpHorFin.Enabled = True
End Sub

Public Sub DescargoInsumo(Documento As String, Area As String, Insumo As String, Cantidad As Double, CantidadReceta As Double, Factor As Double, Tipo As String, PrecioPromedio As Double, Fecha As String)
   Dim nStock    As Double
   Dim nCorrela  As Double
   Dim nCantidad As Double
   Dim sCodProd  As String
      
   Screen.MousePointer = vbHourglass
   If IsNull(Area) Or Len(Area) < 1 Then
      Exit Sub
   End If
   
   If IsNull(Insumo) Or Len(Insumo) < 1 Then
      Exit Sub
   End If
   nCantidad = (Cantidad * CantidadReceta) / IIf(Tipo = "D", 1, Factor)
   'Descargo hacia el Almacen Central
   On Error GoTo ErrorAlmacen
   CnAlmacen.BeginTrans
   Set clsDescAlmacen = New clsAlmacen
   If Area = "000" Then
      sCodProd = clsDescAlmacen.FunInsertamKardex(Insumo, 0, "95", True, Area, Documento, 0, PrecioPromedio * nCantidad, nCantidad, 0, 0, "01", PrecioPromedio, PrecioPromedio, Format(IIf(lCierre, Fecha, fFechaCierre), "YYYY/mm/dd HH:mm:ss"), sUsuario, "")
      If sCodProd = "" Then
         GoTo ErrorAlmacen
      End If
   Else
      sCodProd = clsDescAlmacen.FunInsertamSubKardex(Area, Insumo, 0, "", Format(IIf(lCierre, Fecha, fFechaCierre), "YYYY/mm/dd HH:mm:ss"), sUsuario, "95", True, Documento, 0, nCantidad, PrecioPromedio * nCantidad, 0, 0, "01")
      If sCodProd = "" Then
         GoTo ErrorAlmacen
      End If
   End If
   Set clsDescAlmacen = Nothing
   CnAlmacen.CommitTrans
   Screen.MousePointer = vbDefault
   Exit Sub

ErrorAlmacen:
   CnAlmacen.RollbackTrans
   'MsgBox "Ocurrio un error al momento de descargar el Pedido " & Trim(Documento), vbCritical, "Transferencia"
   mdiAdministracion.StatusBar.Panels.Item(2).Text = "Ocurrio un error al momento de descargar el Pedido " & Trim(Documento)
   Screen.MousePointer = vbDefault
End Sub

Sub DescargoInsumoRecetaBase(ByVal diCodigoPedido As String, ByVal diCodigoProducto As String, ByVal diCantidad As Double, ByVal diFecha As String)
    Dim RsDescargoReceta As ADODB.Recordset
    
    Isql = "SELECT RB.tRecetaBase, RB.tSubArea, RB.tCodigoProducto, isnull(PR.lRecetaBase, 0) as lRecetaBase, isnull(RB.lProducto, 0) as lProducto, isnull(RB.nCantidad, 0) as nCantidadReceta, isnull(PR.nFactor, 0) as nFactor, ISNULL(PR.nPrecioPromedio, 0) AS nPrecioPromedio " & _
           "FROM " & sAlmacenMDB & ".dbo.DRECETABASE RB INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO PR ON RB.tCodigoProducto = PR.tCodigoProducto INNER JOIN " & sAlmacenMDB & ".dbo.MRECETABASE MRB ON RB.tRecetaBase = MRB.tRecetaBase " & _
           "WHERE MRB.tCodigoProducto='" & diCodigoProducto & "' AND (RB.lDescargo=1)"

    Set RsDescargoReceta = Lib.OpenRecordset(Isql, Cn)
    If RsDescargoReceta.RecordCount > 0 Then
       RsDescargoReceta.MoveFirst
       Do While Not RsDescargoReceta.EOF
          If RsDescargoReceta!lRecetaBase And RsDescargoReceta!lProducto = False Then
             DescargoInsumoRecetaBase Trim(Left(Left(diCodigoPedido, 10) & Trim(RsDescargoReceta!tRecetaBase), 15)), RsDescargoReceta!tCodigoProducto, (RsDescargoReceta!nCantidadReceta / RsDescargoReceta!nFactor) * diCantidad, diFecha
          Else
             DescargoInsumo Trim(Left(Left(diCodigoPedido, 10) & Trim(RsDescargoReceta!tRecetaBase), 15)), RsDescargoReceta!tSubArea, RsDescargoReceta!tCodigoProducto, diCantidad, RsDescargoReceta!nCantidadReceta, RsDescargoReceta!nFactor, "R", RsDescargoReceta!nPrecioPromedio, diFecha
          End If
          RsDescargoReceta.MoveNext
       Loop
    End If
    Set RsDescargoReceta = Nothing
End Sub
