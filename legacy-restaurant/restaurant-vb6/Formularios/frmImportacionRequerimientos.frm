VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmImportacionRequerimientos 
   Caption         =   "Importacion de Pedidos"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1365
      ScaleWidth      =   12300
      TabIndex        =   2
      Top             =   6465
      Width           =   12360
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Importar"
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
         Left            =   7440
         Picture         =   "frmImportacionRequerimientos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Generar Pedido"
         Top             =   100
         Width           =   1170
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
         Left            =   8760
         Picture         =   "frmImportacionRequerimientos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   100
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Ver Detalle"
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
         Left            =   6120
         Picture         =   "frmImportacionRequerimientos.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ver Detalle"
         Top             =   100
         Width           =   1170
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   165
         TabIndex        =   12
         Top             =   650
         Width           =   3660
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   2205
            TabIndex        =   13
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
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
            Format          =   41156609
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   450
            TabIndex        =   14
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
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
            Format          =   41156609
            CurrentDate     =   37539
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " al "
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
            Index           =   1
            Left            =   1845
            TabIndex        =   16
            Top             =   270
            Width           =   225
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Del "
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
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   315
            Width           =   405
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   3855
         Picture         =   "frmImportacionRequerimientos.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Filtrar"
         Top             =   700
         Width           =   555
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5805
         TabIndex        =   3
         Top             =   30
         Width           =   5865
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4170
            Picture         =   "frmImportacionRequerimientos.frx":0738
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4710
            Picture         =   "frmImportacionRequerimientos.frx":0C7A
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5250
            Picture         =   "frmImportacionRequerimientos.frx":11BC
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmImportacionRequerimientos.frx":16FE
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmImportacionRequerimientos.frx":1C40
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmImportacionRequerimientos.frx":2182
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Caption         =   "Registro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1650
            TabIndex        =   10
            Top             =   150
            Width           =   2505
         End
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   6405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   5550
         Left            =   90
         TabIndex        =   1
         Top             =   780
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9790
         _LayoutType     =   4
         _RowHeight      =   21
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).FooterText=   "Footer Text"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   3
         Splits(0).AllowColMove=   -1  'True
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   3
         Splits(0).DividerColor=   32768
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Calibri"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Calibri"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Correlativo"
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DataView        =   2
         GroupByCaption  =   "Arrastre aqui la columna a agrupar"
         DeadAreaBackColor=   -2147483637
         ScrollTrack     =   -1  'True
         RowDividerColor =   12632256
         RowSubDividerColor=   -2147483648
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   16711935
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(38)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(42)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(43)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(50)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(55)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(59)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(60)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(63)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(64)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(67)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(71)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(72)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(73)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(74)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(75)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=13,.fontname=Arial Black"
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   225
         Top             =   2115
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo cboCanal 
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Canal de Venta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmImportacionRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim nColumna As Integer

Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Dim i As Integer

'GENERAR MPEDIDO
Dim RsDetalleR As Recordset

'DETALLE DPEDIDO
Dim RsDetallePedido As Recordset

Dim nPVenta As Double
Dim nPBase As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nRecargo As Double
Dim nDescuento As Double
Dim nOficial As Double
Dim sitem As String

Dim RsProductoX As Recordset
Dim RsProducto As Recordset
Dim codProducto As String
Dim cantProducto As Double
Dim RsCanal As Recordset

Sub LlenaCombo()
    With cboCanal
         Isql = "Select * From dbo.vTipoPedido Where lActivo=1 Order by Codigo"
         Set RsCanal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCanal
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
        If RsCanal.RecordCount > 0 Then
            cboCanal.BoundText = sTipoPedidoPD
        End If
    End With

End Sub



Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
               Case Is = 1  'Detalle
                    If RsCabecera.RecordCount > 0 Then
                       sRequerimiento = grdGrilla.Columns(0).value
                       frmImportacionRequerimientoDetalle.Show vbModal
                    Else
                       MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                    End If
                    
               Case Is = 2  'Importar Generar Pedido

                    'CANAL DE VENTA
                    If cboCanal.Text = "" Then
                       MsgBox "Seleccionar un Canal de Venta", vbExclamation, sMensaje
                       Exit Sub
                    Else
                       sTipoPedido = cboCanal.BoundText
                    End If
                 
                    If RsCabecera.RecordCount > 0 Then
                          Dim oComando As clsComando
                          Dim ePedido As String
                          
                          'DETALLE DEL REQUERIMIENTO
                          Isql = "Select CodArt,Articulo,nAtendido As Cantidad From vrequerimiento Where Rq = '" & RsCabecera!Rq & "'"
                          Set RsDetalleR = Lib.OpenRecordset(Isql, CnAlmacen)
                          
                          If RsDetalleR.RecordCount = 0 Then
                             MsgBox "No Existen Productos en Requerimiento", vbExclamation, sMensaje
                          End If
                          
                          'VALIDAR SI YA GENERO PEDIDO
                          ePedido = Calcular(" Select ISNULL(tPedido,'') As Codigo From MREQUERIMIENTO Where tRequerimiento = '" & RsCabecera!Rq & "'", CnAlmacen)
                          If ePedido <> "" Then
                             MsgBox "Existe Pedido Generado Mediante Requerimiento", vbExclamation, sMensaje
                             Exit Sub
                          End If
                          
                          
                          'GENERACION DE PEDIDO
                          Set oComando = New clsComando
                          If Not oComando.CreateCmdSp("spIns_MPEDIDO", Cn) Then
                             Set oComando = Nothing
                             Exit Sub
                          End If
                          oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, ""
                          oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                          oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, 0
                          oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, ""
                          oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, ""
                          oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, ""
                          oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, "0000"
                          oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                          oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 2, sSalon
                          oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
                          oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, "Req: " & RsCabecera!Rq
                          oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
                          oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, Right(sUsuario, 15)
                          oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 10, 0
                          oComando.CreateParameter "@nNino", adInteger, adParamInput, 10, 0
                          oComando.CreateParameter "@nMesa", adInteger, adParamInput, 10, 0
                          oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, ""
                          oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
                          oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
                          oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
                          oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
                          oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
                          oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, 0
                          oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, ""
                          oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, ""
                          oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, ""
                          oComando.CreateParameter "@nTiempoDelivery", adInteger, adParamInput, 10, 0
                          oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, ""
                          oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
                          oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, Null
        
                          oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, ""
                          oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, ""
                                               'entregara
                          oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, ""
                          oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, 0
                          oComando.CreateParameter "@nMontoMaximo", adInteger, adParamInput, 250, 0
                          oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, sPedido
                          oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, ""
                                            
                          If Not oComando.GetParamOK Then
                             Set oComando = Nothing
                             Exit Sub
                          End If
                          If Not oComando.ExecSP Then
                             Set oComando = Nothing
                             Exit Sub
                          Else
                             sPedido = oComando.GetParameterValue("@tPedido")
                          End If
                          
                           
                           'DETALLE
                           For i = 1 To RsDetalleR.RecordCount
                                 codProducto = Calcular("Select Codigo as codigo From vProducto Where tEnlace='" & RsDetalleR!CodArt & "'", Cn)
                                 cantProducto = RsDetalleR!Cantidad
                                 
                                 'ENLACE
                                 If codProducto = "" Or codProducto = "0" Then
                                    Isql = "update MPEDIDO set tEstadoPedido = '03', tTurnoAnulado='" & sTurno & "',tMotivoAnulacion='000',tObservacionAnulado='Error Rq. " & RsCabecera!Rq & "', tUsuarioAnulado='" & Right(sUsuario, 15) & "',fRegAnulado=GETDATE() where tCodigoPedido='" & sPedido & "'"
                                    Cn.Execute Isql
                                    MsgBox "No existe enlace de producto de venta", vbExclamation, sMensaje
                                    Exit Sub
                                 End If
                                 
                                 InsertaProducto cantProducto, codProducto
                                 
                                 RsDetalleR.MoveNext
                           Next i
                           
                           'Bloqueo el Requerimiento Asociado con PEDIDO UPDATE MREQ.
                           Isql = "Update mRequerimiento Set lPedido = 1, tPedido = '" & sPedido & "' Where tRequerimiento = '" & RsCabecera!Rq & "'"
                           CnAlmacen.Execute Isql
                           RsCabecera.Requery
                           MsgBox "Pedido '" & sPedido & "' Generado Correctamente", vbExclamation, sMensaje
                            If sTipoPedido <> "" Then
                                cboCanal.BoundText = sTipoPedido
                            Else
                                cboCanal.BoundText = sTipoPedidoPD
                            End If
                    Else
                       MsgBox "No Existe Datos a Importar", vbExclamation, sMensaje
                    End If
        

               Case Is = 3  'Salir
                    Unload Me
                    
    End Select
End Sub

Private Sub cmdProcesa_Click()
    Isql = "SELECT DISTINCT v.Rq , v.Fecha, v.Area, v.Responsable, v.Estado, ISNULL(v.tPedido,'') As tPedido FROM vRequerimiento v INNER JOIN  TRUTAAREA T ON v.CodigoArea = T.tCodigoArea " & _
           "where v.CodEstado='02' And v.Fecha >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and v.Fecha <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' And T.lImportarPV=1 "
    
    Set RsCabecera = Lib.OpenRecordset(Isql, CnAlmacen)
    
    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 2000 > 0, Me.Height - 2000, 0)
   fraGrilla.Width = IIf(Me.Width - 300 > 0, Me.Width - 300, 0)
   
   grdGrilla.Height = IIf(fraGrilla.Height - 850 > 0, fraGrilla.Height - 850, 0)
   grdGrilla.Width = IIf(fraGrilla.Width - 300 > 0, fraGrilla.Width - 300, 0)
End Sub

Private Sub grdGrilla_FilterChange()
    On Error GoTo errHandler
    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col
    grdGrilla.HoldFields
    RsCabecera.Filter = getFilter(cols)
    grdGrilla.col = c
    grdGrilla.EditActive = True
    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdGrilla
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    ' Ingrese el Titulo
    Me.Caption = " Correlativo de Requerimientos "
    grdGrilla.Caption = Me.Caption
    
    Centrar Me
    nColumna = 0
    
    LlenaCombo
    
    dtpFecIni.value = FechaServidor() - nDias
    dtpFecFin.value = FechaServidor()
    
    
    Isql = "SELECT DISTINCT v.Rq , v.Fecha, v.Area, v.Responsable, v.Estado, ISNULL(v.tPedido,'') As tPedido FROM vRequerimiento v INNER JOIN  TRUTAAREA T ON v.CodigoArea = T.tCodigoArea " & _
           "where v.CodEstado='02' And v.Fecha >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " 00:00' and v.Fecha <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " 23:59' And T.lImportarPV=1 "
    
    Set RsCabecera = Lib.OpenRecordset(Isql, CnAlmacen)
    
    Call ConfGrilla(6, grdGrilla, "Requerimiento", 2, "Rq", 1800, 2, 0, "", _
                                  "Fecha", 2, "Fecha", 1400, 0, 0, "dd/mm/yyyy", _
                                  "Area", 2, "Area", 2200, 0, 0, "", _
                                  "Responsable", 2, "Responsable", 2400, 0, 0, "", _
                                  "Estado", 2, "Estado", 2200, 0, 0, "", _
                                  "Pedido", 2, "tPedido", 1800, 2, 0, "")
            
    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Set rsReporte = Nothing
    Set RsCabecera = Nothing
    Set frmImportacionRequerimientos = Nothing
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub

Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
                fImpresionDiaContable = obtieneDiaContable
End Function

Public Sub InsertaProducto(xCantidad As Double, sProducto As String)
    'HabilitaTimerColor (False)

    Dim precioventa As Double
    Screen.MousePointer = vbHourglass
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim RsOrd As Recordset
    Dim nOrden As Integer

    'MULTIAREAPRODUCCION
    Dim lProductoMultiArea As Boolean
    Dim tsubalmacen As String
    Dim tAreaProduccion As String
    
    'ITEM
    Isql = "select * from DPEDIDO Where tCodigoPedido = '" & sPedido & "'"
    Set RsDetallePedido = Lib.OpenRecordset(Isql, Cn)
    
    sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
    If RsDetallePedido.RecordCount = 0 Then
       'sitem = "001"
       nOrden = 1
    Else
       'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
    End If
    
    'PRODUCTO
     Isql = "select * from vProducto where Codigo = '" & sProducto & "'"
     Set RsProducto = Lib.OpenRecordset(Isql, Cn)
    
    'Precios con Recargos / Descargos por Tipo de Pedido
    nRecargo = 0
    nDescuento = 0
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
    
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3
        
    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    ElseIf sTipoPedido = "04" Then
       If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    ElseIf sTipoPedido = "05" Then
       If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto13, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto14, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto15, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto13
          lImp2 = RsProducto!lImpuesto14
          lImp3 = RsProducto!lImpuesto15
       End If
    Else
       nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    End If
    nPVenta = nOficial
    nOficial = IIf(RsProducto!tMoneda = "02", nOficial * nTC, nOficial)
        
             
    Select Case pais 'ok
        Case "001" 'Bolivia
                nValor = (nValor / 100)
                nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                nValor = 1 + (nValor / 100)
                nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
    End Select
    
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
    ElseIf sTipoPedido = "02" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
       nMObra = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
    ElseIf sTipoPedido = "03" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
       nMObra = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
    ElseIf sTipoPedido = "04" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
       nMObra = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
    Else
       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
       nMObra = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
    End If
    

    'MULTIAREA PRODUCCION
    lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    
    If lProductoMultiArea = False Then
        tsubalmacen = ""
    Else
        tsubalmacen = ""
        If lMultiAreaSubGrupo = True Then
            tAreaProduccion = Calcular("select isnull(tarea,'') codigo from TAREASUBGRUPO where tcaja='" & sCaja & "' and tSubGrupo='" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If
        If lMultiAreaCaja = True Then
            tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If
        
        If tsubalmacen = "0" Then
            tsubalmacen = ""
        End If
    
    End If
    

    'DPEDIDO
    Isql = "insert into DPEDIDO " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, tEstadoItem, tComanda, fRegistro, tMozoD, tUsuarioD, nInsumo, nGasto, nManoObra, nOrden, tUnidadNegocio, tOferta, tsubalmacen,fdiacontable,tcajad ) " & _
           "Values(   '" & sPedido & "', " _
                   & "'" & sTipoPedido & "', '" & sitem & "', '" & sProducto & "', " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nPBase & ", " & nRecargo & ", " & nDescuento & ", " & nOficial & ", " _
                   & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                   & nPVenta & ", " & xCantidad & ", " & nPVenta * xCantidad & ", " _
                   & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                   & "0, '" & RsProducto!tArea & "', " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                   & IIf(RsProducto!lCombinacion, -1, 0) & "," _
                   & IIf(IsNull(RsProducto!nCombinacion), 0, RsProducto!nCombinacion) & "," _
                   & "'N', '', getDate(), '', '" & sUsuario & "', " _
                   & nInsumo & ", " _
                   & nGasto & ", " _
                   & nMObra & ", " _
                   & nOrden & ", '" & sUnidadNegocio & "','','" & tsubalmacen & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', '" & sCaja & "')"
    Cn.Execute Isql
    Screen.MousePointer = vbDefault
End Sub


