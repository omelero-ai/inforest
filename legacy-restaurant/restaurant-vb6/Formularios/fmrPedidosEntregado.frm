VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fmrPedidosEntregado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Entregados"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   13515
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   77266945
      CurrentDate     =   44164
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
      Index           =   6
      Left            =   12360
      Picture         =   "fmrPedidosEntregado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1170
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   6915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   12197
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
   Begin VB.Label Label1 
      Caption         =   "Filtrar Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
End
Attribute VB_Name = "fmrPedidosEntregado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCabecera As Recordset

Private Sub cmdOpcion_Click(Index As Integer)
Unload Me
End Sub



Private Sub DTPicker1_Change()
    LlenarGrilla Format(DTPicker1.Value, "yyyyMMdd")
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
 
    ' Ingrese el Titulo
    Me.Caption = " Pedidos entregados "
        
    grdGrilla.Caption = Me.Caption
    Centrar Me
    
 'Configuración de la Grilla
    Call ConfGrilla(7, grdGrilla, "Pedido", 2, "tCodigoPedido", 1200, 2, 0, "", _
                                   "Fecha", 2, "fFecha", 1200, 0, 0, "dd/MM/yyyy ", _
                                   "Hora Envio", 2, "fAsignacion", 1000, 2, 0, "HH:mm:ss", _
                                   "Cliente", 2, "Cliente", 2800, 0, 0, "", _
                                   "Telefono", 2, "tTelefo", 1500, 0, 0, "", _
                                   "Monto", 2, "Monto", 1500, 1, 0, "###,##0.00", _
                                   "Motorizado", 2, "Motorizado", 2600, 0, 0, "")
                                   
    LlenarGrilla Format(Date, "yyyyMMdd")

    DTPicker1.Value = Date
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub LlenarGrilla(ByVal Fecha As String)
On Error GoTo fin
    Set RsCabecera = Nothing
    Set grdGrilla.DataSource = Nothing
        ' Ingrese la Vista
    Isql = "select * from vDespachador where convert(date,ffecha) ='" & Fecha & "' and tTipoPedido in (select Codigo from vTipoPedido where lActivo =1 and lCanalDelivery=1 ) and tEstadoPedido='02' and CodigoOrigenVenta in (select CodOrigenVenta from vOrigenVenta where lVisualizaDespachador =1 and Activo =1) and isnull(fLlegada,0)<>0 order by tCodigoPedido"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
   
                                           
    Set grdGrilla.DataSource = RsCabecera
    Exit Sub
fin:
MsgBox "Error:" & Error


End Sub

