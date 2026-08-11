VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B52C1CDE-38E9-11D5-98EC-00C0F01C6C81}#1.0#0"; "ifepson.ocx"
Begin VB.Form frmLiquidacionDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Turno       "
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   Icon            =   "frmLiquidacionDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cierre X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   6
      Left            =   10320
      Picture         =   "frmLiquidacionDetalle.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   232
      Top             =   5760
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cierre Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   10320
      Picture         =   "frmLiquidacionDetalle.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   6480
      Visible         =   0   'False
      Width           =   1275
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal PrinterFiscalArgentina 
      Left            =   12000
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame frmDoc 
      Caption         =   "Documentos No Enviados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   210
      Top             =   7920
      Visible         =   0   'False
      Width           =   11655
      Begin MSComctlLib.ProgressBar pgbEnvio 
         Height          =   375
         Left            =   240
         TabIndex        =   216
         Top             =   7920
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdPape 
         Caption         =   "Refrescar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   0
         Left            =   6840
         Picture         =   "frmLiquidacionDetalle.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   215
         ToolTipText     =   "Emite"
         Top             =   240
         Width           =   1515
      End
      Begin VB.CommandButton cmdPape 
         Caption         =   "Enviar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   1
         Left            =   8520
         Picture         =   "frmLiquidacionDetalle.frx":0A40
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton cmdPape 
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
         Height          =   1035
         Index           =   2
         Left            =   10080
         Picture         =   "frmLiquidacionDetalle.frx":0B42
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   211
         Text            =   "frmLiquidacionDetalle.frx":0C34
         Top             =   240
         Width           =   6255
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6075
         Left            =   240
         TabIndex        =   212
         Top             =   1440
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   10716
         _LayoutType     =   4
         _RowHeight      =   26
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   3
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   2
         Splits(0).DividerColor=   32768
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
         _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=Arial"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
         _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(40)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(41)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(42)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(43)  =   "Named:id=34:Heading"
         _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=34,.fontname=Arial"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(53)  =   ":id=36,.fontname=Arial"
         _StyleDefs(54)  =   "Named:id=37:Caption"
         _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(56)  =   "Named:id=38:HighlightRow"
         _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=675"
         _StyleDefs(58)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=38,.fontname=Small Fonts"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblProgreso 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   217
         Top             =   7680
         Visible         =   0   'False
         Width           =   11055
      End
   End
   Begin VB.CommandButton cmdDescargo 
      Caption         =   "Descargar Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   10320
      Picture         =   "frmLiquidacionDetalle.frx":0CC5
      Style           =   1  'Graphical
      TabIndex        =   209
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Observacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   7
      Left            =   10305
      Picture         =   "frmLiquidacionDetalle.frx":1107
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3660
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Paloteo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   10305
      Picture         =   "frmLiquidacionDetalle.frx":1249
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2580
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Liquidación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   10320
      Picture         =   "frmLiquidacionDetalle.frx":177B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1905
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   10305
      Picture         =   "frmLiquidacionDetalle.frx":1CAD
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4335
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   10305
      Picture         =   "frmLiquidacionDetalle.frx":21DF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5010
      Width           =   1275
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
      Height          =   600
      Index           =   1
      Left            =   10305
      Picture         =   "frmLiquidacionDetalle.frx":22E1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7290
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   8370
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   14764
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Resumen"
      TabPicture(0)   =   "frmLiquidacionDetalle.frx":23D3
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label(13)"
      Tab(0).Control(1)=   "Label(12)"
      Tab(0).Control(2)=   "Label(32)"
      Tab(0).Control(3)=   "txtTotalIngresoN"
      Tab(0).Control(4)=   "Label(23)"
      Tab(0).Control(5)=   "txtPuntoN"
      Tab(0).Control(6)=   "txtInicioN"
      Tab(0).Control(7)=   "txtEfectivoN"
      Tab(0).Control(8)=   "txtChequeN"
      Tab(0).Control(9)=   "txtMN"
      Tab(0).Control(10)=   "txtTotalTarjeta"
      Tab(0).Control(11)=   "txtTotalOtroN"
      Tab(0).Control(12)=   "txtPuntoE"
      Tab(0).Control(13)=   "Label(31)"
      Tab(0).Control(14)=   "txtSaldoE"
      Tab(0).Control(15)=   "txtSaldoN"
      Tab(0).Control(16)=   "Label(27)"
      Tab(0).Control(17)=   "txtTotalOtroE"
      Tab(0).Control(18)=   "Label(16)"
      Tab(0).Control(19)=   "txtME"
      Tab(0).Control(20)=   "txtTotalEfectivoN"
      Tab(0).Control(21)=   "txtTotalEfectivoE"
      Tab(0).Control(22)=   "txtEfectivoE"
      Tab(0).Control(23)=   "txtInicioE"
      Tab(0).Control(24)=   "txtTotalTarjetaPropina"
      Tab(0).Control(25)=   "txtChequeE"
      Tab(0).Control(26)=   "Label(10)"
      Tab(0).Control(27)=   "Label(11)"
      Tab(0).Control(28)=   "Label(8)"
      Tab(0).Control(29)=   "Label(0)"
      Tab(0).Control(30)=   "txtTotalIngresoE"
      Tab(0).Control(31)=   "txtRetiroN"
      Tab(0).Control(32)=   "txtRetiroE"
      Tab(0).Control(33)=   "lblReciboN"
      Tab(0).Control(34)=   "lblReciboE"
      Tab(0).Control(35)=   "cmdMovimiento(8)"
      Tab(0).Control(36)=   "cmdMovimiento(7)"
      Tab(0).Control(37)=   "cmdMovimiento(6)"
      Tab(0).Control(38)=   "cmdMovimiento(2)"
      Tab(0).Control(39)=   "cmdMovimiento(4)"
      Tab(0).Control(40)=   "cmdMovimiento(11)"
      Tab(0).Control(41)=   "cmdMovimiento(10)"
      Tab(0).Control(42)=   "cmdMovimiento(1)"
      Tab(0).Control(43)=   "cmdMovimiento(0)"
      Tab(0).Control(44)=   "frameOME"
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "Tarjetas de Crédito"
      TabPicture(1)   =   "frmLiquidacionDetalle.frx":23EF
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTarjeta(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtTPropina"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtTTarjeta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblTarjeta(7)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtPropina(7)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtTarjeta(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label(20)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblTarjeta(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblTarjeta(8)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblTarjeta(5)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblTarjeta(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblTarjeta(6)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblTarjeta(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblTarjeta(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtPropina(8)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtPropina(6)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtPropina(5)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtPropina(4)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtPropina(3)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtPropina(2)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtPropina(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtTarjeta(8)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtTarjeta(6)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtTarjeta(5)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtTarjeta(4)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtTarjeta(3)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtTarjeta(2)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtTarjeta(1)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label(14)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmdTarjeta(7)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "cmdPropina(7)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmdTarjeta(1)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmdPropina(1)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "cmdTarjeta(2)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "cmdPropina(2)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "cmdTarjeta(3)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "cmdPropina(3)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "cmdTarjeta(4)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "cmdPropina(4)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "cmdTarjeta(5)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "cmdPropina(5)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "cmdTarjeta(6)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "cmdPropina(6)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "cmdTarjeta(8)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "cmdPropina(8)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "Otros Tipos de Pago"
      TabPicture(2)   =   "frmLiquidacionDetalle.frx":240B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtOtroE(21)"
      Tab(2).Control(1)=   "txtOtroN(21)"
      Tab(2).Control(2)=   "lblOtro(21)"
      Tab(2).Control(3)=   "SSTab1"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Observación"
      TabPicture(3)   =   "frmLiquidacionDetalle.frx":2427
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtObservacion"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Otras tarjetas 9-16"
      TabPicture(4)   =   "frmLiquidacionDetalle.frx":2443
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtTarjeta(9)"
      Tab(4).Control(1)=   "txtPropina(9)"
      Tab(4).Control(2)=   "lblTarjeta(9)"
      Tab(4).Control(3)=   "txtTarjeta(10)"
      Tab(4).Control(4)=   "txtPropina(10)"
      Tab(4).Control(5)=   "lblTarjeta(10)"
      Tab(4).Control(6)=   "txtTarjeta(13)"
      Tab(4).Control(7)=   "txtPropina(13)"
      Tab(4).Control(8)=   "lblTarjeta(11)"
      Tab(4).Control(9)=   "txtTarjeta(11)"
      Tab(4).Control(10)=   "txtPropina(11)"
      Tab(4).Control(11)=   "lblTarjeta(12)"
      Tab(4).Control(12)=   "txtTarjeta(12)"
      Tab(4).Control(13)=   "txtPropina(12)"
      Tab(4).Control(14)=   "lblTarjeta(13)"
      Tab(4).Control(15)=   "Label(1)"
      Tab(4).Control(16)=   "Label(2)"
      Tab(4).Control(17)=   "txtTarjeta(16)"
      Tab(4).Control(18)=   "txtPropina(16)"
      Tab(4).Control(19)=   "lblTarjeta(14)"
      Tab(4).Control(20)=   "txtTarjeta(14)"
      Tab(4).Control(21)=   "txtPropina(14)"
      Tab(4).Control(22)=   "lblTarjeta(15)"
      Tab(4).Control(23)=   "txtTarjeta(15)"
      Tab(4).Control(24)=   "txtPropina(15)"
      Tab(4).Control(25)=   "lblTarjeta(16)"
      Tab(4).Control(26)=   "cmdPropina(9)"
      Tab(4).Control(27)=   "cmdTarjeta(9)"
      Tab(4).Control(28)=   "cmdPropina(10)"
      Tab(4).Control(29)=   "cmdTarjeta(10)"
      Tab(4).Control(30)=   "cmdPropina(13)"
      Tab(4).Control(31)=   "cmdTarjeta(13)"
      Tab(4).Control(32)=   "cmdPropina(11)"
      Tab(4).Control(33)=   "cmdTarjeta(11)"
      Tab(4).Control(34)=   "cmdPropina(12)"
      Tab(4).Control(35)=   "cmdTarjeta(12)"
      Tab(4).Control(36)=   "cmdPropina(16)"
      Tab(4).Control(37)=   "cmdTarjeta(16)"
      Tab(4).Control(38)=   "cmdPropina(14)"
      Tab(4).Control(39)=   "cmdTarjeta(14)"
      Tab(4).Control(40)=   "cmdPropina(15)"
      Tab(4).Control(41)=   "cmdTarjeta(15)"
      Tab(4).ControlCount=   42
      TabCaption(5)   =   "Otras Tarjetas 17-24"
      TabPicture(5)   =   "frmLiquidacionDetalle.frx":245F
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdTarjeta(24)"
      Tab(5).Control(1)=   "cmdPropina(24)"
      Tab(5).Control(2)=   "cmdTarjeta(23)"
      Tab(5).Control(3)=   "cmdPropina(23)"
      Tab(5).Control(4)=   "cmdTarjeta(22)"
      Tab(5).Control(5)=   "cmdPropina(22)"
      Tab(5).Control(6)=   "cmdTarjeta(21)"
      Tab(5).Control(7)=   "cmdPropina(21)"
      Tab(5).Control(8)=   "cmdTarjeta(20)"
      Tab(5).Control(9)=   "cmdPropina(20)"
      Tab(5).Control(10)=   "cmdTarjeta(18)"
      Tab(5).Control(11)=   "cmdPropina(18)"
      Tab(5).Control(12)=   "cmdTarjeta(17)"
      Tab(5).Control(13)=   "cmdPropina(17)"
      Tab(5).Control(14)=   "cmdTarjeta(19)"
      Tab(5).Control(15)=   "cmdPropina(19)"
      Tab(5).Control(16)=   "lblTarjeta(24)"
      Tab(5).Control(17)=   "txtPropina(24)"
      Tab(5).Control(18)=   "txtTarjeta(24)"
      Tab(5).Control(19)=   "lblTarjeta(23)"
      Tab(5).Control(20)=   "txtPropina(23)"
      Tab(5).Control(21)=   "txtTarjeta(23)"
      Tab(5).Control(22)=   "lblTarjeta(22)"
      Tab(5).Control(23)=   "txtPropina(22)"
      Tab(5).Control(24)=   "txtTarjeta(22)"
      Tab(5).Control(25)=   "lblTarjeta(21)"
      Tab(5).Control(26)=   "txtPropina(21)"
      Tab(5).Control(27)=   "txtTarjeta(21)"
      Tab(5).Control(28)=   "lblTarjeta(20)"
      Tab(5).Control(29)=   "txtPropina(20)"
      Tab(5).Control(30)=   "txtTarjeta(20)"
      Tab(5).Control(31)=   "lblTarjeta(19)"
      Tab(5).Control(32)=   "txtPropina(18)"
      Tab(5).Control(33)=   "txtTarjeta(18)"
      Tab(5).Control(34)=   "lblTarjeta(18)"
      Tab(5).Control(35)=   "txtPropina(17)"
      Tab(5).Control(36)=   "txtTarjeta(17)"
      Tab(5).Control(37)=   "Label(4)"
      Tab(5).Control(38)=   "Label(3)"
      Tab(5).Control(39)=   "lblTarjeta(17)"
      Tab(5).Control(40)=   "txtPropina(19)"
      Tab(5).Control(41)=   "txtTarjeta(19)"
      Tab(5).ControlCount=   42
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   24
         Left            =   -71940
         TabIndex        =   313
         Top             =   5940
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   24
         Left            =   -67125
         TabIndex        =   312
         Top             =   5940
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   23
         Left            =   -71940
         TabIndex        =   308
         Top             =   5235
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   23
         Left            =   -67125
         TabIndex        =   307
         Top             =   5235
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   22
         Left            =   -71940
         TabIndex        =   303
         Top             =   4545
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   22
         Left            =   -67125
         TabIndex        =   302
         Top             =   4545
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   21
         Left            =   -71940
         TabIndex        =   298
         Top             =   3855
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   21
         Left            =   -67125
         TabIndex        =   297
         Top             =   3855
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   20
         Left            =   -71940
         TabIndex        =   293
         Top             =   3150
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   20
         Left            =   -67125
         TabIndex        =   292
         Top             =   3150
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   18
         Left            =   -71940
         TabIndex        =   288
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   18
         Left            =   -67125
         TabIndex        =   287
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   17
         Left            =   -71940
         TabIndex        =   283
         Top             =   1065
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   17
         Left            =   -67125
         TabIndex        =   282
         Top             =   1065
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   19
         Left            =   -71940
         TabIndex        =   276
         Top             =   2460
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   19
         Left            =   -67125
         TabIndex        =   275
         Top             =   2460
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   15
         Left            =   -71940
         TabIndex        =   271
         Top             =   5235
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   15
         Left            =   -67125
         TabIndex        =   270
         Top             =   5235
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   14
         Left            =   -71940
         TabIndex        =   266
         Top             =   4545
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   14
         Left            =   -67125
         TabIndex        =   265
         Top             =   4545
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   16
         Left            =   -71940
         TabIndex        =   261
         Top             =   5940
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   16
         Left            =   -67125
         TabIndex        =   260
         Top             =   5940
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   12
         Left            =   -71940
         TabIndex        =   254
         Top             =   3150
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   12
         Left            =   -67125
         TabIndex        =   253
         Top             =   3150
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   11
         Left            =   -71940
         TabIndex        =   249
         Top             =   2460
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   11
         Left            =   -67125
         TabIndex        =   248
         Top             =   2460
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   13
         Left            =   -71940
         TabIndex        =   244
         Top             =   3855
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   13
         Left            =   -67125
         TabIndex        =   243
         Top             =   3855
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   10
         Left            =   -71940
         TabIndex        =   239
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   10
         Left            =   -67125
         TabIndex        =   238
         Top             =   1800
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   9
         Left            =   -71940
         TabIndex        =   234
         Top             =   1065
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   9
         Left            =   -67125
         TabIndex        =   233
         Top             =   1065
         Width           =   1275
      End
      Begin VB.Frame frameOME 
         Caption         =   "Frame1"
         Height          =   7455
         Left            =   -68000
         TabIndex        =   218
         Top             =   640
         Width           =   3135
         Begin VB.Label lblReciboE3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   1680
            TabIndex        =   231
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lblReciboE2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   120
            TabIndex        =   230
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label txtSaldoE3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1680
            TabIndex        =   228
            Top             =   5940
            Width           =   1410
         End
         Begin VB.Label txtSaldoE2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   227
            Top             =   5940
            Width           =   1410
         End
         Begin VB.Label txtTotalEfectivoE3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1680
            TabIndex        =   226
            Top             =   5460
            Width           =   1410
         End
         Begin VB.Label txtTotalEfectivoE2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   225
            Top             =   5460
            Width           =   1410
         End
         Begin VB.Label txtTotalIngresoE3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Left            =   1680
            TabIndex        =   224
            Top             =   5040
            Width           =   1410
         End
         Begin VB.Label txtTotalIngresoE2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Left            =   120
            TabIndex        =   223
            Top             =   5040
            Width           =   1410
         End
         Begin VB.Label txtEfectivoE3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   1680
            TabIndex        =   222
            Top             =   1320
            Width           =   1410
         End
         Begin VB.Label txtEfectivoE2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Left            =   120
            TabIndex        =   221
            Top             =   1320
            Width           =   1410
         End
         Begin VB.Label txtME2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   220
            Top             =   240
            Width           =   690
         End
         Begin VB.Label txtME3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   219
            Top             =   240
            Width           =   690
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   104
         Top             =   660
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   12515
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Otros Pagos"
         TabPicture(0)   =   "frmLiquidacionDetalle.frx":247B
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblOtro(9)"
         Tab(0).Control(1)=   "lblOtro(10)"
         Tab(0).Control(2)=   "txtOtroE(10)"
         Tab(0).Control(3)=   "txtOtroN(10)"
         Tab(0).Control(4)=   "txtOtroE(9)"
         Tab(0).Control(5)=   "txtOtroN(9)"
         Tab(0).Control(6)=   "lblOtro(7)"
         Tab(0).Control(7)=   "txtOtroN(7)"
         Tab(0).Control(8)=   "lblOtro(4)"
         Tab(0).Control(9)=   "lblOtro(8)"
         Tab(0).Control(10)=   "lblOtro(5)"
         Tab(0).Control(11)=   "lblOtro(3)"
         Tab(0).Control(12)=   "lblOtro(6)"
         Tab(0).Control(13)=   "lblOtro(1)"
         Tab(0).Control(14)=   "txtOtroN(8)"
         Tab(0).Control(15)=   "txtOtroN(6)"
         Tab(0).Control(16)=   "txtOtroN(5)"
         Tab(0).Control(17)=   "txtOtroN(4)"
         Tab(0).Control(18)=   "txtOtroN(3)"
         Tab(0).Control(19)=   "txtOtroN(1)"
         Tab(0).Control(20)=   "lblMontoN(0)"
         Tab(0).Control(21)=   "txtOtroE(1)"
         Tab(0).Control(22)=   "txtOtroE(3)"
         Tab(0).Control(23)=   "txtOtroE(4)"
         Tab(0).Control(24)=   "txtOtroE(5)"
         Tab(0).Control(25)=   "txtOtroE(6)"
         Tab(0).Control(26)=   "txtOtroE(8)"
         Tab(0).Control(27)=   "txtOtroE(7)"
         Tab(0).Control(28)=   "txtOtroE(2)"
         Tab(0).Control(29)=   "txtOtroN(2)"
         Tab(0).Control(30)=   "lblOtro(2)"
         Tab(0).Control(31)=   "lblMontoE(0)"
         Tab(0).Control(32)=   "cmdOtroE(10)"
         Tab(0).Control(33)=   "cmdOtroN(10)"
         Tab(0).Control(34)=   "cmdOtroE(9)"
         Tab(0).Control(35)=   "cmdOtroN(9)"
         Tab(0).Control(36)=   "cmdOtroN(7)"
         Tab(0).Control(37)=   "cmdOtroN(1)"
         Tab(0).Control(38)=   "cmdOtroN(3)"
         Tab(0).Control(39)=   "cmdOtroN(4)"
         Tab(0).Control(40)=   "cmdOtroN(5)"
         Tab(0).Control(41)=   "cmdOtroN(6)"
         Tab(0).Control(42)=   "cmdOtroN(8)"
         Tab(0).Control(43)=   "cmdOtroE(8)"
         Tab(0).Control(44)=   "cmdOtroE(6)"
         Tab(0).Control(45)=   "cmdOtroE(5)"
         Tab(0).Control(46)=   "cmdOtroE(4)"
         Tab(0).Control(47)=   "cmdOtroE(3)"
         Tab(0).Control(48)=   "cmdOtroE(1)"
         Tab(0).Control(49)=   "cmdOtroE(7)"
         Tab(0).Control(50)=   "cmdOtroE(2)"
         Tab(0).Control(51)=   "cmdOtroN(2)"
         Tab(0).ControlCount=   52
         TabCaption(1)   =   "Otros Pagos"
         TabPicture(1)   =   "frmLiquidacionDetalle.frx":2497
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblOtro(19)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblOtro(20)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "txtOtroE(20)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "txtOtroN(20)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "txtOtroE(19)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtOtroN(19)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lblOtro(17)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtOtroN(17)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "lblOtro(14)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "lblOtro(18)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "lblOtro(15)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "lblOtro(13)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "lblOtro(16)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "lblOtro(11)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtOtroN(18)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txtOtroN(16)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "txtOtroN(15)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "txtOtroN(14)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "txtOtroN(13)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "txtOtroN(11)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "lblMontoN(1)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "txtOtroE(11)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "txtOtroE(13)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "txtOtroE(14)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "txtOtroE(15)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "txtOtroE(16)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "txtOtroE(18)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "txtOtroE(17)"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "txtOtroE(12)"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "txtOtroN(12)"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "lblOtro(12)"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "lblMontoE(1)"
         Tab(1).Control(31).Enabled=   0   'False
         Tab(1).Control(32)=   "cmdOtroE(20)"
         Tab(1).Control(32).Enabled=   0   'False
         Tab(1).Control(33)=   "cmdOtroN(20)"
         Tab(1).Control(33).Enabled=   0   'False
         Tab(1).Control(34)=   "cmdOtroE(19)"
         Tab(1).Control(34).Enabled=   0   'False
         Tab(1).Control(35)=   "cmdOtroN(19)"
         Tab(1).Control(35).Enabled=   0   'False
         Tab(1).Control(36)=   "cmdOtroN(17)"
         Tab(1).Control(36).Enabled=   0   'False
         Tab(1).Control(37)=   "cmdOtroN(11)"
         Tab(1).Control(37).Enabled=   0   'False
         Tab(1).Control(38)=   "cmdOtroN(13)"
         Tab(1).Control(38).Enabled=   0   'False
         Tab(1).Control(39)=   "cmdOtroN(14)"
         Tab(1).Control(39).Enabled=   0   'False
         Tab(1).Control(40)=   "cmdOtroN(15)"
         Tab(1).Control(40).Enabled=   0   'False
         Tab(1).Control(41)=   "cmdOtroN(16)"
         Tab(1).Control(41).Enabled=   0   'False
         Tab(1).Control(42)=   "cmdOtroN(18)"
         Tab(1).Control(42).Enabled=   0   'False
         Tab(1).Control(43)=   "cmdOtroE(18)"
         Tab(1).Control(43).Enabled=   0   'False
         Tab(1).Control(44)=   "cmdOtroE(16)"
         Tab(1).Control(44).Enabled=   0   'False
         Tab(1).Control(45)=   "cmdOtroE(15)"
         Tab(1).Control(45).Enabled=   0   'False
         Tab(1).Control(46)=   "cmdOtroE(14)"
         Tab(1).Control(46).Enabled=   0   'False
         Tab(1).Control(47)=   "cmdOtroE(13)"
         Tab(1).Control(47).Enabled=   0   'False
         Tab(1).Control(48)=   "cmdOtroE(11)"
         Tab(1).Control(48).Enabled=   0   'False
         Tab(1).Control(49)=   "cmdOtroE(17)"
         Tab(1).Control(49).Enabled=   0   'False
         Tab(1).Control(50)=   "cmdOtroE(12)"
         Tab(1).Control(50).Enabled=   0   'False
         Tab(1).Control(51)=   "cmdOtroN(12)"
         Tab(1).Control(51).Enabled=   0   'False
         Tab(1).ControlCount=   52
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   12
            Left            =   3195
            TabIndex        =   176
            Top             =   1305
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   12
            Left            =   7650
            TabIndex        =   175
            Top             =   1305
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   17
            Left            =   7650
            TabIndex        =   174
            Top             =   4425
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   11
            Left            =   7650
            TabIndex        =   173
            Top             =   675
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   13
            Left            =   7650
            TabIndex        =   172
            Top             =   1920
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   14
            Left            =   7650
            TabIndex        =   171
            Top             =   2550
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   15
            Left            =   7650
            TabIndex        =   170
            Top             =   3180
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   16
            Left            =   7650
            TabIndex        =   169
            Top             =   3810
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   18
            Left            =   7650
            TabIndex        =   168
            Top             =   5055
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   18
            Left            =   3195
            TabIndex        =   167
            Top             =   5055
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   16
            Left            =   3195
            TabIndex        =   166
            Top             =   3810
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   15
            Left            =   3195
            TabIndex        =   165
            Top             =   3180
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   14
            Left            =   3195
            TabIndex        =   164
            Top             =   2550
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   13
            Left            =   3195
            TabIndex        =   163
            Top             =   1920
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   540
            Index           =   11
            Left            =   3195
            TabIndex        =   162
            Top             =   675
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   17
            Left            =   3195
            TabIndex        =   161
            Top             =   4425
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   19
            Left            =   3195
            TabIndex        =   160
            Top             =   5685
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   19
            Left            =   7650
            TabIndex        =   159
            Top             =   5685
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   20
            Left            =   3195
            TabIndex        =   158
            Top             =   6315
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   20
            Left            =   7650
            TabIndex        =   157
            Top             =   6315
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   2
            Left            =   -71805
            TabIndex        =   124
            Top             =   1305
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   2
            Left            =   -67350
            TabIndex        =   123
            Top             =   1305
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   7
            Left            =   -67350
            TabIndex        =   122
            Top             =   4425
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   1
            Left            =   -67350
            TabIndex        =   121
            Top             =   675
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   3
            Left            =   -67350
            TabIndex        =   120
            Top             =   1920
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   4
            Left            =   -67350
            TabIndex        =   119
            Top             =   2550
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   5
            Left            =   -67350
            TabIndex        =   118
            Top             =   3180
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   6
            Left            =   -67350
            TabIndex        =   117
            Top             =   3810
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   8
            Left            =   -67350
            TabIndex        =   116
            Top             =   5055
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   8
            Left            =   -71805
            TabIndex        =   115
            Top             =   5055
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   6
            Left            =   -71805
            TabIndex        =   114
            Top             =   3840
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   5
            Left            =   -71805
            TabIndex        =   113
            Top             =   3180
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   4
            Left            =   -71805
            TabIndex        =   112
            Top             =   2550
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   3
            Left            =   -71805
            TabIndex        =   111
            Top             =   1920
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   540
            Index           =   1
            Left            =   -71805
            TabIndex        =   110
            Top             =   675
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   7
            Left            =   -71805
            TabIndex        =   109
            Top             =   4425
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   9
            Left            =   -71805
            TabIndex        =   108
            Top             =   5685
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   9
            Left            =   -67350
            TabIndex        =   107
            Top             =   5685
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroN 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   10
            Left            =   -71805
            TabIndex        =   106
            Top             =   6315
            Width           =   1275
         End
         Begin VB.CommandButton cmdOtroE 
            Caption         =   "Retiro Moneda Nacional"
            Height          =   555
            Index           =   10
            Left            =   -67350
            TabIndex        =   105
            Top             =   6315
            Width           =   1275
         End
         Begin VB.Label lblMontoE 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
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
            Index           =   1
            Left            =   6975
            TabIndex        =   208
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   12
            Left            =   600
            TabIndex        =   207
            Top             =   1305
            Width           =   990
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   12
            Left            =   4590
            TabIndex        =   206
            Top             =   1305
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   12
            Left            =   6240
            TabIndex        =   205
            Top             =   1305
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   17
            Left            =   6240
            TabIndex        =   204
            Top             =   4425
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   18
            Left            =   6240
            TabIndex        =   203
            Top             =   5055
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   16
            Left            =   6240
            TabIndex        =   202
            Top             =   3810
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   15
            Left            =   6240
            TabIndex        =   201
            Top             =   3180
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   14
            Left            =   6240
            TabIndex        =   200
            Top             =   2550
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   13
            Left            =   6240
            TabIndex        =   199
            Top             =   1920
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   11
            Left            =   6240
            TabIndex        =   198
            Top             =   675
            Width           =   1305
         End
         Begin VB.Label lblMontoN 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
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
            Index           =   1
            Left            =   5175
            TabIndex        =   197
            Top             =   360
            Width           =   540
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   11
            Left            =   4590
            TabIndex        =   196
            Top             =   675
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   13
            Left            =   4590
            TabIndex        =   195
            Top             =   1920
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   14
            Left            =   4590
            TabIndex        =   194
            Top             =   2550
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   15
            Left            =   4590
            TabIndex        =   193
            Top             =   3180
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   16
            Left            =   4590
            TabIndex        =   192
            Top             =   3810
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   18
            Left            =   4590
            TabIndex        =   191
            Top             =   5055
            Width           =   1305
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   11
            Left            =   600
            TabIndex        =   190
            Top             =   675
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   16
            Left            =   600
            TabIndex        =   189
            Top             =   3810
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   13
            Left            =   600
            TabIndex        =   188
            Top             =   1920
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   15
            Left            =   600
            TabIndex        =   187
            Top             =   3180
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   18
            Left            =   600
            TabIndex        =   186
            Top             =   5055
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   14
            Left            =   600
            TabIndex        =   185
            Top             =   2550
            Width           =   990
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   17
            Left            =   4590
            TabIndex        =   184
            Top             =   4425
            Width           =   1305
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   17
            Left            =   600
            TabIndex        =   183
            Top             =   4425
            Width           =   990
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   19
            Left            =   4590
            TabIndex        =   182
            Top             =   5685
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   19
            Left            =   6240
            TabIndex        =   181
            Top             =   5685
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   20
            Left            =   4590
            TabIndex        =   180
            Top             =   6315
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   20
            Left            =   6240
            TabIndex        =   179
            Top             =   6315
            Width           =   1305
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   20
            Left            =   600
            TabIndex        =   178
            Top             =   6315
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   19
            Left            =   600
            TabIndex        =   177
            Top             =   5685
            Width           =   990
         End
         Begin VB.Label lblMontoE 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
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
            Index           =   0
            Left            =   -68025
            TabIndex        =   156
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   2
            Left            =   -74400
            TabIndex        =   155
            Top             =   1305
            Width           =   990
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   2
            Left            =   -70410
            TabIndex        =   154
            Top             =   1305
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   2
            Left            =   -68760
            TabIndex        =   153
            Top             =   1305
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   7
            Left            =   -68760
            TabIndex        =   152
            Top             =   4425
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   8
            Left            =   -68760
            TabIndex        =   151
            Top             =   5055
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   6
            Left            =   -68760
            TabIndex        =   150
            Top             =   3810
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   5
            Left            =   -68760
            TabIndex        =   149
            Top             =   3180
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   4
            Left            =   -68760
            TabIndex        =   148
            Top             =   2550
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   3
            Left            =   -68760
            TabIndex        =   147
            Top             =   1920
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   1
            Left            =   -68760
            TabIndex        =   146
            Top             =   675
            Width           =   1305
         End
         Begin VB.Label lblMontoN 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
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
            Index           =   0
            Left            =   -69825
            TabIndex        =   145
            Top             =   360
            Width           =   540
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   1
            Left            =   -70410
            TabIndex        =   144
            Top             =   675
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   3
            Left            =   -70410
            TabIndex        =   143
            Top             =   1920
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   4
            Left            =   -70410
            TabIndex        =   142
            Top             =   2550
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   5
            Left            =   -70410
            TabIndex        =   141
            Top             =   3180
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   6
            Left            =   -70410
            TabIndex        =   140
            Top             =   3810
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   8
            Left            =   -70410
            TabIndex        =   139
            Top             =   5055
            Width           =   1305
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   1
            Left            =   -74400
            TabIndex        =   138
            Top             =   675
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   6
            Left            =   -74400
            TabIndex        =   137
            Top             =   3810
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Left            =   -74400
            TabIndex        =   136
            Top             =   1920
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   5
            Left            =   -74400
            TabIndex        =   135
            Top             =   3180
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   8
            Left            =   -74400
            TabIndex        =   134
            Top             =   5055
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   4
            Left            =   -74400
            TabIndex        =   133
            Top             =   2550
            Width           =   990
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   7
            Left            =   -70410
            TabIndex        =   132
            Top             =   4425
            Width           =   1305
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   7
            Left            =   -74400
            TabIndex        =   131
            Top             =   4425
            Width           =   990
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   9
            Left            =   -70410
            TabIndex        =   130
            Top             =   5685
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   9
            Left            =   -68760
            TabIndex        =   129
            Top             =   5685
            Width           =   1305
         End
         Begin VB.Label txtOtroN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   10
            Left            =   -70410
            TabIndex        =   128
            Top             =   6315
            Width           =   1305
         End
         Begin VB.Label txtOtroE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   10
            Left            =   -68760
            TabIndex        =   127
            Top             =   6315
            Width           =   1305
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   10
            Left            =   -74400
            TabIndex        =   126
            Top             =   6315
            Width           =   990
         End
         Begin VB.Label lblOtro 
            AutoSize        =   -1  'True
            Caption         =   "Otro Pago :"
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
            Index           =   9
            Left            =   -74400
            TabIndex        =   125
            Top             =   5685
            Width           =   990
         End
      End
      Begin VB.TextBox txtObservacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7395
         Left            =   -74820
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   99
         Top             =   840
         Width           =   9900
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   7875
         TabIndex        =   66
         Top             =   5940
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   3060
         TabIndex        =   65
         Top             =   5940
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   7875
         TabIndex        =   64
         Top             =   4545
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   3060
         TabIndex        =   63
         Top             =   4545
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   5
         Left            =   7875
         TabIndex        =   62
         Top             =   3855
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   5
         Left            =   3060
         TabIndex        =   61
         Top             =   3855
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   7875
         TabIndex        =   60
         Top             =   3240
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   3060
         TabIndex        =   59
         Top             =   3150
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   3
         Left            =   7875
         TabIndex        =   58
         Top             =   2460
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   3
         Left            =   3060
         TabIndex        =   57
         Top             =   2460
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   2
         Left            =   7875
         TabIndex        =   56
         Top             =   1755
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   2
         Left            =   3060
         TabIndex        =   55
         Top             =   1755
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   7875
         TabIndex        =   54
         Top             =   1065
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   3060
         TabIndex        =   53
         Top             =   1065
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   7
         Left            =   7875
         TabIndex        =   52
         Top             =   5235
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   7
         Left            =   3060
         TabIndex        =   51
         Top             =   5235
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Ventas Efectivo MN"
         Height          =   555
         Index           =   0
         Left            =   -72765
         TabIndex        =   19
         Top             =   1845
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   -67890
         TabIndex        =   18
         Top             =   1965
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   10
         Left            =   -72840
         TabIndex        =   17
         Top             =   6480
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Extranjera"
         Height          =   555
         Index           =   11
         Left            =   -67890
         TabIndex        =   16
         Top             =   6600
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   -72765
         TabIndex        =   15
         Top             =   3975
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Tarjetas de Crédito"
         Height          =   555
         Index           =   2
         Left            =   -72765
         TabIndex        =   14
         Top             =   2610
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   -72765
         TabIndex        =   13
         Top             =   3300
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Extranjera"
         Height          =   555
         Index           =   7
         Left            =   -67890
         TabIndex        =   12
         Top             =   3420
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   -72765
         TabIndex        =   11
         Top             =   4680
         Width           =   1260
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   24
         Left            =   -74460
         TabIndex        =   316
         Top             =   5940
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   24
         Left            =   -68775
         TabIndex        =   315
         Top             =   5940
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   24
         Left            =   -70425
         TabIndex        =   314
         Top             =   5940
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   23
         Left            =   -74460
         TabIndex        =   311
         Top             =   5235
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   23
         Left            =   -68775
         TabIndex        =   310
         Top             =   5235
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   23
         Left            =   -70425
         TabIndex        =   309
         Top             =   5235
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   22
         Left            =   -74460
         TabIndex        =   306
         Top             =   4545
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   22
         Left            =   -68775
         TabIndex        =   305
         Top             =   4545
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   22
         Left            =   -70425
         TabIndex        =   304
         Top             =   4545
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   21
         Left            =   -74460
         TabIndex        =   301
         Top             =   3855
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   21
         Left            =   -68775
         TabIndex        =   300
         Top             =   3855
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   21
         Left            =   -70425
         TabIndex        =   299
         Top             =   3855
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   20
         Left            =   -74460
         TabIndex        =   296
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   20
         Left            =   -68775
         TabIndex        =   295
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   20
         Left            =   -70425
         TabIndex        =   294
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   19
         Left            =   -74460
         TabIndex        =   291
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   18
         Left            =   -68775
         TabIndex        =   290
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   18
         Left            =   -70425
         TabIndex        =   289
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   18
         Left            =   -74460
         TabIndex        =   286
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   17
         Left            =   -68775
         TabIndex        =   285
         Top             =   1065
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   17
         Left            =   -70425
         TabIndex        =   284
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propina"
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
         Index           =   4
         Left            =   -68025
         TabIndex        =   281
         Top             =   750
         Width           =   660
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Left            =   -69555
         TabIndex        =   280
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   17
         Left            =   -74460
         TabIndex        =   279
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   19
         Left            =   -68775
         TabIndex        =   278
         Top             =   2460
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   19
         Left            =   -70425
         TabIndex        =   277
         Top             =   2460
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   16
         Left            =   -74460
         TabIndex        =   274
         Top             =   5940
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   15
         Left            =   -68775
         TabIndex        =   273
         Top             =   5235
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   15
         Left            =   -70425
         TabIndex        =   272
         Top             =   5235
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   15
         Left            =   -74460
         TabIndex        =   269
         Top             =   5235
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   14
         Left            =   -68775
         TabIndex        =   268
         Top             =   4545
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   14
         Left            =   -70425
         TabIndex        =   267
         Top             =   4545
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   14
         Left            =   -74460
         TabIndex        =   264
         Top             =   4545
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   16
         Left            =   -68775
         TabIndex        =   263
         Top             =   5940
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   16
         Left            =   -70425
         TabIndex        =   262
         Top             =   5940
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propina"
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
         Index           =   2
         Left            =   -68025
         TabIndex        =   259
         Top             =   750
         Width           =   660
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Index           =   1
         Left            =   -69555
         TabIndex        =   258
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   13
         Left            =   -74460
         TabIndex        =   257
         Top             =   3855
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   12
         Left            =   -68775
         TabIndex        =   256
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   12
         Left            =   -70440
         TabIndex        =   255
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   12
         Left            =   -74460
         TabIndex        =   252
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   11
         Left            =   -68775
         TabIndex        =   251
         Top             =   2460
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   11
         Left            =   -70425
         TabIndex        =   250
         Top             =   2460
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   11
         Left            =   -74460
         TabIndex        =   247
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   13
         Left            =   -68775
         TabIndex        =   246
         Top             =   3855
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   13
         Left            =   -70425
         TabIndex        =   245
         Top             =   3855
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   10
         Left            =   -74460
         TabIndex        =   242
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   10
         Left            =   -68775
         TabIndex        =   241
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   10
         Left            =   -70425
         TabIndex        =   240
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   9
         Left            =   -74460
         TabIndex        =   237
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   9
         Left            =   -68775
         TabIndex        =   236
         Top             =   1065
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   9
         Left            =   -70425
         TabIndex        =   235
         Top             =   1065
         Width           =   1410
      End
      Begin VB.Label lblReciboE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -69465
         TabIndex        =   103
         Top             =   2325
         Width           =   1410
      End
      Begin VB.Label lblReciboN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -71460
         TabIndex        =   102
         Top             =   2325
         Width           =   1770
      End
      Begin VB.Label txtRetiroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   -67890
         TabIndex        =   101
         Top             =   7275
         Width           =   1410
      End
      Begin VB.Label txtRetiroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   -72660
         TabIndex        =   100
         Top             =   7275
         Width           =   1410
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Total Otros Pagos :"
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
         Index           =   21
         Left            =   -74775
         TabIndex        =   98
         Top             =   7845
         Width           =   1665
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   21
         Left            =   -70305
         TabIndex        =   97
         Top             =   7845
         Width           =   1305
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   21
         Left            =   -68655
         TabIndex        =   96
         Top             =   7845
         Width           =   1305
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Index           =   14
         Left            =   5445
         TabIndex        =   95
         Top             =   750
         Width           =   540
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   1
         Left            =   4575
         TabIndex        =   94
         Top             =   1065
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   2
         Left            =   4575
         TabIndex        =   93
         Top             =   1755
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   3
         Left            =   4575
         TabIndex        =   92
         Top             =   2460
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   4
         Left            =   4575
         TabIndex        =   91
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   5
         Left            =   4575
         TabIndex        =   90
         Top             =   3855
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   6
         Left            =   4575
         TabIndex        =   89
         Top             =   4545
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   8
         Left            =   4575
         TabIndex        =   88
         Top             =   5940
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   1
         Left            =   6225
         TabIndex        =   87
         Top             =   1065
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   2
         Left            =   6225
         TabIndex        =   86
         Top             =   1755
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   3
         Left            =   6225
         TabIndex        =   85
         Top             =   2460
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   4
         Left            =   6225
         TabIndex        =   84
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   5
         Left            =   6225
         TabIndex        =   83
         Top             =   3855
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   6
         Left            =   6225
         TabIndex        =   82
         Top             =   4545
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   8
         Left            =   6225
         TabIndex        =   81
         Top             =   5940
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   1
         Left            =   540
         TabIndex        =   80
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   2
         Left            =   540
         TabIndex        =   79
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   6
         Left            =   540
         TabIndex        =   78
         Top             =   4545
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Left            =   540
         TabIndex        =   77
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   5
         Left            =   540
         TabIndex        =   76
         Top             =   3855
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   8
         Left            =   540
         TabIndex        =   75
         Top             =   5940
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   4
         Left            =   540
         TabIndex        =   74
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propina"
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
         Index           =   20
         Left            =   6975
         TabIndex        =   73
         Top             =   750
         Width           =   660
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   7
         Left            =   4575
         TabIndex        =   72
         Top             =   5235
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   7
         Left            =   6225
         TabIndex        =   71
         Top             =   5235
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   7
         Left            =   540
         TabIndex        =   70
         Top             =   5235
         Width           =   735
      End
      Begin VB.Label txtTTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4575
         TabIndex        =   69
         Top             =   6720
         Width           =   1410
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6225
         TabIndex        =   68
         Top             =   6720
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Total Tarjeta :"
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
         Index           =   0
         Left            =   540
         TabIndex        =   67
         Top             =   6720
         Width           =   1230
      End
      Begin VB.Label txtTotalIngresoE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   -69465
         TabIndex        =   50
         Top             =   5700
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Inicio de Caja :"
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
         Index           =   0
         Left            =   -73860
         TabIndex        =   49
         Top             =   1425
         Width           =   1305
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Efectivo en Caja :"
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
         Index           =   8
         Left            =   -74835
         TabIndex        =   48
         Top             =   6105
         Width           =   2040
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Cheques / Depósitos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   11
         Left            =   -74400
         TabIndex        =   47
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total de Tarjetas :"
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
         Index           =   10
         Left            =   -74400
         TabIndex        =   46
         Top             =   2730
         Width           =   1590
      End
      Begin VB.Label txtChequeE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -69465
         TabIndex        =   45
         Top             =   3420
         Width           =   1410
      End
      Begin VB.Label txtTotalTarjetaPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   -69465
         TabIndex        =   44
         Top             =   2730
         Width           =   1410
      End
      Begin VB.Label txtInicioE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -69465
         TabIndex        =   43
         Top             =   1380
         Width           =   1410
      End
      Begin VB.Label txtEfectivoE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -69465
         TabIndex        =   42
         Top             =   1965
         Width           =   1410
      End
      Begin VB.Label txtTotalEfectivoE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -69465
         TabIndex        =   41
         Top             =   6105
         Width           =   1410
      End
      Begin VB.Label txtTotalEfectivoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -71460
         TabIndex        =   40
         Top             =   6105
         Width           =   1770
      End
      Begin VB.Label txtME 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69465
         TabIndex        =   39
         Top             =   885
         Width           =   690
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Final en Caja :"
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
         Index           =   16
         Left            =   -74580
         TabIndex        =   38
         Top             =   6600
         Width           =   1785
      End
      Begin VB.Label txtTotalOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -69465
         TabIndex        =   37
         Top             =   4095
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Otros Tipos Pago :"
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
         Index           =   27
         Left            =   -74880
         TabIndex        =   36
         Top             =   4080
         Width           =   2100
      End
      Begin VB.Label txtSaldoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -71460
         TabIndex        =   35
         Top             =   6600
         Width           =   1770
      End
      Begin VB.Label txtSaldoE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -69465
         TabIndex        =   34
         Top             =   6600
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Puntos :"
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
         Index           =   31
         Left            =   -74040
         TabIndex        =   33
         Top             =   4770
         Width           =   1215
      End
      Begin VB.Label txtPuntoE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -69465
         TabIndex        =   32
         Top             =   4770
         Width           =   1410
      End
      Begin VB.Label txtTotalOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -71460
         TabIndex        =   31
         Top             =   4095
         Width           =   1770
      End
      Begin VB.Label txtTotalTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -71460
         TabIndex        =   30
         Top             =   2730
         Width           =   1770
      End
      Begin VB.Label txtMN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71100
         TabIndex        =   29
         Top             =   885
         Width           =   690
      End
      Begin VB.Label txtChequeN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -71460
         TabIndex        =   28
         Top             =   3420
         Width           =   1770
      End
      Begin VB.Label txtEfectivoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -71460
         TabIndex        =   27
         Top             =   1965
         Width           =   1770
      End
      Begin VB.Label txtInicioN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -71100
         TabIndex        =   26
         Top             =   1380
         Width           =   1410
      End
      Begin VB.Label txtPuntoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   -71460
         TabIndex        =   25
         Top             =   4770
         Width           =   1770
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Ingresos :"
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
         Index           =   23
         Left            =   -74145
         TabIndex        =   24
         Top             =   5700
         Width           =   1350
      End
      Begin VB.Label txtTotalIngresoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   -71460
         TabIndex        =   23
         Top             =   5700
         Width           =   1770
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Efectivo :"
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
         Index           =   32
         Left            =   -74130
         TabIndex        =   22
         Top             =   1965
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "( Cobrado + Recibos - Egresos )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   12
         Left            =   -74895
         TabIndex        =   21
         Top             =   2190
         Width           =   1980
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "( Cobrado + Recibos )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   13
         Left            =   -74160
         TabIndex        =   20
         Top             =   2955
         Width           =   1350
      End
   End
   Begin VB.Label txtUsuario 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   825
      TabIndex        =   3
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label TxtFecha 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8835
      TabIndex        =   2
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Usuario :"
      Height          =   195
      Index           =   19
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   630
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      Height          =   195
      Index           =   18
      Left            =   8205
      TabIndex        =   0
      Top             =   90
      Width           =   540
   End
End
Attribute VB_Name = "frmLiquidacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSumas As Recordset
Dim RsTarjetas As Recordset
Dim RsTarjeta As Recordset
Dim RsOtro As Recordset
Dim RsOtros As Recordset
Dim RsPago As Recordset
Dim RsEgreso As Recordset
Dim RsIngreso As Recordset
Dim oComando As New clsComando
Dim nCambio As Double
Dim nCambio2 As Double
Dim nCambio3 As Double
Dim xFecha As Date
Dim nInicioN As Double
Dim nInicioE As Double
Dim nEfectivoN As Double
Dim nEfectivoE As Double
Dim nEfectivoE2 As Double
Dim nEfectivoE3 As Double
Dim nTotalTarjeta As Double
Dim nTotalTarjetaPropina As Double
Dim nTarjeta(24) As Double
Dim nTarjetaPropina(24) As Double
Dim nChequeN As Double
Dim nChequeE As Double
Dim nTotalOtroN As Double
Dim nTotalOtroE As Double
Dim nOtroN(20) As Double   'era 10
Dim nOtroE(20) As Double
Dim nPuntoN As Double
Dim nPuntoE As Double
Dim nDolar As Double
Dim nDolar2 As Double
Dim nDolar3 As Double
Dim tTarjeta(24) As String
Dim tOtro(20) As String
Dim canttarjetas As Integer


Dim nRetiroN As Double
Dim nRetiroE As Double
'CgMiranda-----------------------
Dim RsExisteMensajes As Recordset
'Fin CGMiranda-----------------------
Dim nReciboIngresoN As Double
Dim nReciboIngresoE As Double
Dim nReciboAnticipoN As Double
Dim nReciboAnticipoE As Double
Dim nReciboEgresoN As Double
Dim nReciboEgresoE As Double
Dim nReciboTarjeta As Double
Dim nCuentaCobrar As Double
Dim nCuentaCorriente As Double
Dim nTotalEfectivoN As Double
Dim nTotalEfectivoE As Double
Dim nTotalEfectivoE2 As Double
Dim nTotalEfectivoE3 As Double
Dim nTotalIngresoN As Double
Dim nTotalIngresoE As Double
Dim nTotalIngresoE2 As Double
Dim nTotalIngresoE3 As Double
Dim nSaldoN As Double
Dim nSaldoE As Double
Dim nEgresoN As Double
Dim nEgresoE As Double
Dim nEgresoE2 As Double
Dim nEgresoE3 As Double
Dim nIngresoN As Double
Dim nIngresoE As Double

Dim nFinalN As Double
Dim nFinalE As Double
Dim nFinalE2 As Double
Dim nFinalE3 As Double

Dim rsDocNoEnv As Recordset

'jsjr 07-07-2025 -----------------------------------
Dim RsPagoNiubiz As Recordset
Dim RsOperationTerminal As Recordset
Dim RsOperationCaja As Recordset
'------------------------------
Dim RsPagoIzipay As Recordset
'------------------------------

Dim i As Integer
Dim FEenvio As Boolean

Private Sub cmdDescargo_Click()
    Dim RsDescargoAutomatico As New Recordset
    Set RsDescargoAutomatico = Lib.OpenRecordset("select lDescargoAutomatico from tparametro", Cn)
    If IIf(IsNull(RsDescargoAutomatico!lDescargoAutomatico), False, RsDescargoAutomatico!lDescargoAutomatico) Then
     MsgBox "No se puede usar el descargo manual porque el descargo automático está activo.", vbInformation, sMensaje
     Exit Sub
    End If
    
  frmDescargo.Show vbModal
End Sub

Private Sub cmdMovimiento_Click(Index As Integer)
  Select Case Index
         Case Is = 0 ' Efectivo MN
              sTipo = ""
              frmNumPad.Show vbModal
              nEfectivoN = IIf(wEnter = True, sDescrip, nEfectivoN)
              txtEfectivoN.Caption = Format(nEfectivoN, "###,###,###,##0.00")
              CalcularTotales
  
         Case Is = 1 ' Efectivo ME
              sTipo = ""
              frmNumPad.Show vbModal
              nEfectivoE = IIf(wEnter = True, sDescrip, nEfectivoE)
              txtEfectivoE.Caption = Format(nEfectivoE, "###,###,###,##0.00")
              CalcularTotales
                               
         Case Is = 2 ' Tarjetas de Credito
              SSTab.Tab = 1
                                             
         Case Is = 4 ' Otros Pagos
              SSTab.Tab = 2
                                                                                          
         Case Is = 6 ' Cheque MN
              sTipo = ""
              frmNumPad.Show vbModal
              nChequeN = IIf(wEnter = True, sDescrip, nChequeN)
              txtChequeN.Caption = Format(nChequeN, "###,###,###,##0.00")
              CalcularTotales

         Case Is = 7 ' Cheque MN
              sTipo = ""
              frmNumPad.Show vbModal
              nChequeE = IIf(wEnter = True, sDescrip, nChequeE)
              txtChequeE.Caption = Format(nChequeE, "###,###,###,##0.00")
              CalcularTotales
              
         Case Is = 8 ' puntos MN
              sTipo = ""
              frmNumPad.Show vbModal
              nPuntoN = IIf(wEnter = True, sDescrip, nPuntoN)
              txtPuntoN.Caption = Format(nPuntoN, "###,###,###,##0.00")
              CalcularTotales

         Case Is = 10 ' Retiro MN
              sTipo = ""
              frmNumPad.Show vbModal
              nRetiroN = IIf(wEnter = True, sDescrip, nRetiroN)
              
              If nRetiroN > nTotalEfectivoN Then
                 nRetiroN = 0
                 MsgBox "No se puede retirar mas del Efectivo"
              End If
              txtRetiroN.Caption = Format(nRetiroN, "###,###,###,##0.00")
              CalcularTotales

         Case Is = 11 ' Retiro MN
              sTipo = ""
              frmNumPad.Show vbModal
              nRetiroE = IIf(wEnter = True, sDescrip, nRetiroE)
              
              If nRetiroE > nTotalEfectivoE Then
                 nRetiroE = 0
                 MsgBox "No se puede retirar mas del Efectivo"
              End If
              txtRetiroE.Caption = Format(nRetiroE, "###,###,###,##0.00")
              CalcularTotales
  End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
  Dim nCorrela As String
  Select Case Index
            Case Is = 0 ' Cierre
              If lObligaCierre Then
                 If (Supervisor("11") = False) Then
                    MsgBox "Clave No Permitida", vbExclamation, sMensaje
                    Exit Sub
                 End If
              End If
              
              If lActivaConsultaDescargo Then
                    If MsgBox("Debe Descargar Ventas antes de cerrar el turno. ¿ Realizó este proceso  ?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
                        If MsgBox("¿Desea realizar el Descargo de ventas?", vbYesNo + vbQuestion, sMensaje) = vbYes Then
                            frmDescargo.Show vbModal
                        Else
                        Exit Sub
                        End If
                    End If
              End If
                    
              If MsgBox("Seguro de Cerrar el Turno " & sTurno & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                 Exit Sub
              End If
              
              'jsjr 07-07-2025 -----------------------------------------------------------------------------------
              
               If Calcular("select COUNT(*) as codigo from TCAJA where tCaja = '" & sCaja & "' and lActivaIntegracionNiubiz = 1", Cn) > 0 Then
               
                    Set RsOperationCaja = Lib.OpenRecordset("select IdTerminal from DPAGOPOST where LTRIM(RTRIM(Codigo)) != '' and tTurno='" & sTurno & "' and Integracion = 'NIUBIZ' group by IdTerminal", Cn)
                    
                    Do While Not RsOperationCaja.EOF
                    
                       Dim TotalCobrados As Integer
                       Dim TotalAnulados As Integer
                       
                       TotalCobrados = Calcular("select sum(nMonto) as codigo from DPAGOPOST where  tTurno = '" & sTurno & "' and IdTerminal = '" & RsOperationCaja!IdTerminal & "' and  Integracion = 'NIUBIZ' and tEstadoDocumento = '02'", Cn)
                       TotalAnulados = Calcular("select sum(nMonto) as codigo from DPAGOPOST where  tTurno = '" & sTurno & "' and IdTerminal = '" & RsOperationCaja!IdTerminal & "' and  Integracion = 'NIUBIZ' and tEstadoDocumento = '04'", Cn)
                       
                      
                      If Calcular("select count(*) as codigo from MCIERREPOST where Integracion = 'NIUBIZ' AND tTurno = '" & sTurno & "' and IdTerminal = '" & RsOperationCaja!IdTerminal & "'", Cn) = 0 Then
                      
                        Dim RsTerminal As Recordset
                        Dim NombreTerminal As String
                        NombreTerminal = "Niubiz"
                        
                        
                        Set RsTerminal = Lib.OpenRecordset("select tDescripcion from TTERMINAL where tIntegracion = '01' and Codigo = '" & RsOperationCaja!IdTerminal & "'", Cn)
                        If RsTerminal.RecordCount = 1 Then
                            NombreTerminal = RsTerminal!tDescripcion
                        End If
                        
                        
                        If MsgBox("Se han detectado movimientos asociados a la terminal '" & NombreTerminal & "'" & vbCrLf & _
                                  "¿Desea enviar la intencion de cierre?", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
                            
                            
                            Isql = "INSERT INTO MCIERREPOST (IdTerminal,Integracion,tTurno, lEnvio, nMontoIN, nMontoDE, tLote,tUsuario, fRegistro)" & _
                              " VALUES ( '" & RsOperationCaja!IdTerminal & "','NIUBIZ','" & sTurno & "',0 ," & TotalCobrados & "," & TotalAnulados & ",'" & sCaja & "','" & sUsuario & "',GETDATE())"
                      
                            Cn.Execute Isql
                            
                            Set RsOperationTerminal = Lib.OpenRecordset("execute SP_EJECUTA_MOTOR_INALAMBRICO '" & "CIERRE," & sTurno & ",'", Cn)
                      
                      
                            If RsOperationTerminal!Valor = "1" Then
                                MsgBox "Se ha enviado una solicitud de cierre al terminal (NIUBIZ). Por favor, verifique el dispositivo.", vbExclamation, sMensaje
                            Else
                            
                                Isql = "DELETE FROM MCIERREPOST WHERE tTurno='" & sTurno & "' and Integracion='NIUBIZ' and IdTerminal='" & RsOperationCaja!IdTerminal & "'"
                                Cn.Execute Isql
                                
                                MsgBox "Error al enviar el cierre al terminal (NIUBIZ), Por favor, verifique su configuracion.", vbCritical, sMensaje
                                Exit Sub
                            End If
                               
                        Else
                        
                            Isql = "INSERT INTO MCIERREPOST (IdTerminal,Integracion,tTurno, lEnvio, nMontoIN, nMontoDE, tLote,tUsuario, fRegistro)" & _
                              " VALUES ( '" & RsOperationCaja!IdTerminal & "','NIUBIZ','" & sTurno & "',0," & TotalCobrados & ", " & TotalAnulados & ",'','" & sUsuario & "',GETDATE())"
                      
                            Cn.Execute Isql
                      
                        End If
                        
                      End If
                      
                      RsOperationCaja.MoveNext
                    Loop
                    
              
                End If
              '----------------------------------------
              
              'REPORTE IZIPAY---CSHM 10/10/2025
              If Calcular("select COUNT(*) as codigo from TCAJA where tCaja = '" & sCaja & "' and lActivaIntegracionIzipay = 1", Cn) > 0 Then
               
                    Set RsPagoIzipay = Lib.OpenRecordset("select sum(nMonto) AS TOTAL, count(*) as Cantidad from DPAGOPOST where tTurno = '" & sTurno & "' and Integracion = 'IZIPAY'", Cn)
                    
                    If IIf(IsNull(RsPagoIzipay!Cantidad), 0, RsPagoIzipay!Cantidad) > 0 Then
                    
                      Set RsOperationCaja = Lib.OpenRecordset("select top 1 tTerminal from TCAJATERMINAL where tCaja = '" & sCaja & "' and lActivo = 1", Cn)
                       
                      If RsOperationCaja.RecordCount = 0 Then
                          MsgBox "Hay operaciones registrados en tu terminal, pero no se encontró un enlace activo en la caja actual para efectuar el cierre.", vbExclamation, sMensaje
                          Exit Sub
                      End If
                      
                        Isql = "INSERT INTO MCIERREPOST (IdTerminal,Integracion,tTurno, lEnvio, nMontoIN, nMontoDE, tLote,tUsuario, fRegistro)" & _
                              " VALUES ( '" & RsOperationCaja!tTerminal & "','IZIPAY','" & sTurno & "',0," & IIf(IsNull(RsPagoIzipay!TOTAL), 0, RsPagoIzipay!TOTAL) & ", " & "0.0" & ",'" & sCaja & "','" & sUsuario & "',GETDATE())"

                      Cn.Execute Isql
                      
                      Dim sComando As String
                                Dim sResultado As String
                                Dim oShell As Object
                                Dim oExec As Object
                                Dim sLinea As String
                                Dim sPathExe As String
                               
                               sPathExe = "C:\INFOMATICA\MotorPagoInalambrico\ServicioPagoInalambrico.exe"
                                sComando = "CIERRE," & sTurno & "," & ""
                                
                                On Error Resume Next
                                Set oShell = CreateObject("WScript.Shell")
                                Set oExec = oShell.Exec("""" & sPathExe & """ " & sComando)
                                
                                ' Leer salida del EXE (si devuelve texto)
                                Do Until oExec.StdOut.AtEndOfStream
                                    sLinea = oExec.StdOut.ReadLine
                                    sResultado = sResultado & sLinea
                                Loop
                                
                                If Err.Number <> 0 Then
                                    MsgBox "Error al ejecutar el servicio del PINPAD: " & Err.Description, vbCritical, sMensaje
                                    Err.Clear
                                End If
                                
                                ' --- Validar resultado ---
                                If InStr(1, sResultado, "1") > 0 Then
                                    MsgBox "Respuesta exitosa del terminal (Izipay).", vbInformation, sMensaje
                                    
                                    Dim cmd As Object
                                    Set cmd = CreateObject("ADODB.Command")
                                
                                    With cmd
                                        .ActiveConnection = Cn
                                        .CommandType = adCmdStoredProc
                                        .CommandText = "SP_InsertPagoIntegracion"
                                
                                        .Parameters.Append .CreateParameter("@documento", adVarChar, adParamInput, 50, sTurno)
                                        .Parameters.Append .CreateParameter("@estado", adBoolean, adParamInput, , False)
                                        .Parameters.Append .CreateParameter("@fecha", adDate, adParamInput, , Now)
                                        .Parameters.Append .CreateParameter("@tipo", adInteger, adParamInput, , 6)
                                
                                        .Execute
                                    End With
                                
                                    Set cmd = Nothing
                                    
                                Else
                                    Isql = "DELETE FROM MCIERREPOST WHERE tTurno='" & sTurno & "'"
                                    Cn.Execute Isql
                                    MsgBox "No se pudo enviar la solicitud de cobro al terminal (Izipay). Verifique configuración.", vbCritical, sMensaje
                                    Exit Sub
                                End If
                                
                                ' Limpieza
                                Set oExec = Nothing
                                Set oShell = Nothing
                      
                      
                      Isql = "INSERT INTO MCIERREPOST (IdTerminal,Integracion,tTurno, lEnvio, nMontoIN, nMontoDE, tLote,tUsuario, fRegistro)" & _
                              " VALUES ( '" & RsOperationCaja!tTerminal & "','IZIPAY','" & sTurno & "',0," & IIf(IsNull(RsPagoIzipay!TOTAL), 0, RsPagoIzipay!TOTAL) & ", " & "0.0" & ",'','" & sUsuario & "',GETDATE())"

                      Cn.Execute Isql
                      
                    End If
                End If
              '-------------
              

               'CGMiranda ---------------------------------------------------------------------------------------
               Dim sql As String
                sql = "usp_listarmensajes"
                Set oComando = New clsComando
                If Not oComando.CreateCmdSp(sql, Cn) Then
                    Set oComando = Nothing
                    Exit Sub
                End If
                oComando.CreateParameter "@fechaini", adDBDate, adParamInput, 8, Null
                oComando.CreateParameter "@fechafin", adDBDate, adParamInput, 15, Null
                oComando.CreateParameter "@tcaja", adVarChar, adParamInput, 3, sCaja
                If Not oComando.GetParamOK Then
                    Set oComando = Nothing
                    Exit Sub
                End If
                Set RsExisteMensajes = oComando.GetSP()
                If Not (RsExisteMensajes.EOF Or RsExisteMensajes.BOF) Then
                   
                    If MsgBox("Existen mensajes en los puntos de Adición. Desea desactivarlos ?", vbQuestion + vbOKCancel, sMensaje) = vbOK Then
                              sql = "USP_CERRAR_MENSAJES_CIERRETURNO"
                                If Not oComando.CreateCmdSp(sql, Cn) Then
                                    Set oComando = Nothing
                                    Exit Sub
                                End If
                                oComando.CreateParameter "@usuario", adVarChar, adParamInput, 15, sUsuario
                                oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                                If Not oComando.GetParamOK Then
                                    Set oComando = Nothing
                                    Exit Sub
                                End If
                                If Not oComando.ExecSP Then
                                    Set oComando = Nothing
                                    Exit Sub
                                End If
                    End If
                End If
                'fIN cgMIRANDA--------------------------------------------------------------------------------------
                
                For i = 1 To 24
                    If i > 8 Then
                        nTarjeta(8) = nTarjeta(8) + nTarjeta(i)
                    End If
                Next i
                
              Isql = "Update MTURNO Set " & _
                     "lCierre = 1, " & _
                     "fFinal = getdate(), " & _
                     "nMontoEN =" & nEfectivoN & ", " & _
                     "nMontoEE =" & nEfectivoE & ", " & _
                     "nMontoEE2 =" & nEfectivoE2 & ", " & _
                     "nMontoEE3 =" & nEfectivoE3 & ", " & _
                     "nMontoCN =" & nChequeN & ", " & _
                     "nMontoCE =" & nChequeE & ", " & _
                     "nMontoPN =" & nPuntoN & ", " & _
                     "nMontoPE =" & nPuntoE & ", " & _
                     "nMontoFN =" & nTotalEfectivoN - nRetiroN & ", " & _
                     "nMontoFE =" & nTotalEfectivoE - nRetiroE & ", " & _
                     "nMontoFE2 =" & nTotalEfectivoE2 & ", " & _
                     "nMontoFE3 =" & nTotalEfectivoE3 & ", " & _
                     "nTarjeta1 =" & nTarjeta(1) & ", nTarjeta2 =" & nTarjeta(2) & ", nTarjeta3 =" & nTarjeta(3) & ", nTarjeta4 =" & nTarjeta(4) & ", nTarjeta5 =" & nTarjeta(5) & ", nTarjeta6 =" & nTarjeta(6) & ", nTarjeta7 =" & nTarjeta(7) & ", nTarjeta8 =" & nTarjeta(8) & ", " & _
                     "nPropina1 =" & nTarjetaPropina(1) & ", nPropina2 =" & nTarjetaPropina(2) & ", nPropina3 =" & nTarjetaPropina(3) & ", nPropina4 =" & nTarjetaPropina(4) & ", nPropina5 =" & nTarjetaPropina(5) & ", nPropina6 =" & nTarjetaPropina(6) & ", nPropina7 =" & nTarjetaPropina(7) & ", nPropina8 =" & nTarjetaPropina(8) & ", " & _
                     " nOtroN1 =" & nOtroN(1) & ", nOtroN2 =" & nOtroN(2) & ", nOtroN3 =" & nOtroN(3) & ", nOtroN4 =" & nOtroN(4) & ", nOtroN5 =" & nOtroN(5) & ", nOtroN6 =" & nOtroN(6) & ", nOtroN7 =" & nOtroN(7) & ", nOtroN8 =" & nOtroN(8) & ", nOtroN9 =" & nOtroN(9) & ", nOtroN10 =" & nOtroN(10) & ",nOtroN11 =" & nOtroN(11) & ",nOtroN12 =" & nOtroN(12) & ",nOtroN13 =" & nOtroN(13) & ",nOtroN14 =" & nOtroN(14) & ",nOtroN15 =" & nOtroN(15) & ",nOtroN16 =" & nOtroN(16) & ",nOtroN17 =" & nOtroN(17) & ",nOtroN18 =" & nOtroN(18) & ",nOtroN19 =" & nOtroN(19) & ",nOtroN20 =" & nOtroN(20) & ", " & _
                     " nOtroE1 =" & nOtroE(1) & ", nOtroE2 =" & nOtroE(2) & ", nOtroE3 =" & nOtroE(3) & ", nOtroE4 =" & nOtroE(4) & ", nOtroE5 =" & nOtroE(5) & ", nOtroE6 =" & nOtroE(6) & ", nOtroE7 =" & nOtroE(7) & ", nOtroE8 =" & nOtroE(8) & ", nOtroE9 =" & nOtroE(9) & ", nOtroE10 =" & nOtroE(10) & ",nOtroE11 =" & nOtroE(11) & ",nOtroE12 =" & nOtroE(12) & ",nOtroE13 =" & nOtroE(13) & ",nOtroE14 =" & nOtroE(14) & ",nOtroE15 =" & nOtroE(15) & ",nOtroE16 =" & nOtroE(16) & ",nOtroE17 =" & nOtroE(17) & ",nOtroE18 =" & nOtroE(18) & ",nOtroE19 =" & nOtroE(19) & ",nOtroE20 =" & nOtroE(20) & ", " & _
                     "tObservacion='" & txtObservacion.Text & "', nDiferencia=0, lReplica=1 " & _
                     "where tTurno = '" & sTurno & "'"
                     'ojo
              Cn.Execute Isql

              wInicio = False
              MsgBox "Turno " & sTurno & " cerrado satisfactoriamente", vbInformation, sMensaje
              
              If pais = "003" Then
                cmdOpcion(0).Enabled = False
                cmdOpcion(7).Enabled = False
                cmdDescargo.Enabled = False
                SSTab.Enabled = False
              Else
                Unload Me
              End If
              'DIA CONTABLE
                If lDiaContable = False Then
                        frmDiaContable.obtieneModoIngreso "Cerrar"
                         frmDiaContable.Show vbModal
                End If
             'DIA CONTABLE
             
             
             'ENVIA LOS RESUMENES DE BOL Y FAC - GCAA 23012023
                If (lFEBiz) And (lFEBizOffLine) Then
                    INSERTA_FE_INFOREST "", 3, FechaServidor()
                    INSERTA_FE_INFOREST "", 3, FechaServidor() - 1
                    INSERTA_FE_INFOREST "", 3, FechaServidor() - 2
                    INSERTA_FE_INFOREST "", 3, FechaServidor() - 3
                    
                    INSERTA_FE_INFOREST "", 4, FechaServidor()
                    INSERTA_FE_INFOREST "", 4, FechaServidor() - 1
                    INSERTA_FE_INFOREST "", 4, FechaServidor() - 2
                    INSERTA_FE_INFOREST "", 4, FechaServidor() - 3
                End If
                
                If pais = "000" Then
                    If lFacturacionE = True And lFEBiz = False And lFEOfisis = False And lFEpape = False And lFECarbajal = False And lFESpring = False And lFEGesa = False Then
                        INSERTAFE Format(FechaServidor() - 1, "yyyymmdd"), "", 3, ""
                        INSERTAFE Format(FechaServidor() - 2, "yyyymmdd"), "", 3, ""
                        INSERTAFE Format(FechaServidor() - 3, "yyyymmdd"), "", 3, ""
                        INSERTAFE Format(FechaServidor() - 4, "yyyymmdd"), "", 3, ""
                        INSERTAFE Format(FechaServidor() - 5, "yyyymmdd"), "", 3, ""
                    End If
                End If

         Case Is = 1 ' Cancelar
              Unload Me
         
         Case Is = 2 ' Impresion
              Cn.Execute "UPDATE MTURNO SET TOBSERVACION='" & txtObservacion.Text & "' where tturno='" & sTurno & "'"
              ImprimeLiquidacion
                      
         Case Is = 3
            If (Supervisor("11") = False) Then
                MsgBox "Clave No Permitida", vbExclamation, sMensaje
                Exit Sub
             End If
         
              If MsgBox("Seguro de realizar el Cierre Z ?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                 Exit Sub
              End If
              
              If pais = "003" Then
                If ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 0, Nothing) = 0 Then
                    MsgBox "No se ralizo el cierre Z", vbInformation, "Impresora fiscal"
                    Exit Sub
                Else
                    MsgBox "El cierre Z, se realizo satisfactoriamente!!", vbInformation, "Impresora fiscal"
                End If
              End If

         Case Is = 4 ' Imprime Liquidacion
               Cn.Execute "UPDATE MTURNO SET TOBSERVACION='" & txtObservacion.Text & "' where tturno='" & sTurno & "'"
               With frmRepLiquidacionTicket
                    .cmdBusca(0).Enabled = False
                    .txtTurno = sTurno
                    .chkUsuario.Enabled = False
                    .chkTurno.Enabled = False
                    .chkNoCortesia.Enabled = False
                    .chkCortesia.Enabled = False
                    .Show vbModal
               End With
              
         Case Is = 5 ' Imprime Paloteo
               Cn.Execute "UPDATE MTURNO SET TOBSERVACION='" & txtObservacion.Text & "' where tturno='" & sTurno & "'"
               With frmRepPaloteoTicket
                    .chkTurno.value = 0
                    .cmdBusca(1).Enabled = False
                    .txtTurno = sTurno
                    .sTurno = sTurno
                    .chkTurno.Enabled = False
                    .Show vbModal
               End With
         Case Is = 6
         
              If pais = "003" Then
                If ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 1, Nothing) = 0 Then
                    MsgBox "Percance al realizar cierre X, favor de revisar si se imprimio el cierre, caso contrario revisar fuente de papel o conexion!!!", vbInformation, "Impresora fiscal"
                    'Exit Sub
                End If
              End If
         Case Is = 7 ' Imprime Paloteo
              SSTab.Tab = 3
         
  End Select
End Sub

Private Sub cmdOtroE_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nOtroE(Index) = IIf(wEnter = True, sDescrip, nOtroE(Index))
    txtOtroE(Index).Caption = Format(nOtroE(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub cmdOtroN_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nOtroN(Index) = IIf(wEnter = True, sDescrip, nOtroN(Index))
    txtOtroN(Index).Caption = Format(nOtroN(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub cmdPape_Click(Index As Integer)
On Error GoTo fin
    Select Case Index
        Case Is = 0
            Set grdGrilla.DataSource = Nothing
            Set rsDocNoEnv = Lib.OpenRecordset("exec usp_ListDocumentosFE '" & sCaja & "','20180101','20180101',1 ", Cn)
            If rsDocNoEnv.RecordCount > 0 Then
                Set grdGrilla.DataSource = rsDocNoEnv
            Else
                MsgBox "Todos los Documentos Fueron Enviados!!!", vbInformation
                Set grdGrilla.DataSource = Nothing
                Me.frmDoc.Visible = False
            End If
        Case 1
            If lFEpape Then
                Screen.MousePointer = vbHourglass
                If rsDocNoEnv.RecordCount > 0 Then
                    rsDocNoEnv.MoveFirst
                    Do While Not rsDocNoEnv.EOF
                        If Not FacturarTCPIP(2, rsDocNoEnv!Documento, 0) Then
                        End If
                        Sleep 2000
                        If Not FacturarTCPIP(3, rsDocNoEnv!Documento, 0) Then
                        End If
                        Sleep 1000
                        rsDocNoEnv.MoveNext
                    Loop
                    cmdPape_Click (0)
                Else
                    Me.frmDoc.Visible = False
                End If
                Screen.MousePointer = vbDefault
            ElseIf lFEBiz Then
                Screen.MousePointer = vbHourglass
                Me.pgbEnvio.value = 0
                Me.pgbEnvio.Min = 0
                Me.pgbEnvio.Max = rsDocNoEnv.RecordCount
                Me.pgbEnvio.Visible = True
                Me.lblProgreso.Visible = True
                If rsDocNoEnv.RecordCount > 0 Then
                    rsDocNoEnv.MoveFirst
                    Me.lblProgreso.Caption = "Enviando Documento...: "
                    Do While Not rsDocNoEnv.EOF
                        DoEvents
                        Dim sd As String
                        Me.pgbEnvio.value = Me.pgbEnvio.value + 1
                        If rsDocNoEnv!Doc = "D" Then
                            If Not INSERTA_FE_INFOREST(rsDocNoEnv!Documento, 1, DateTime.Date) Then
                            End If
                        ElseIf rsDocNoEnv!Doc = "N" Then
                            If Not INSERTA_FE_INFOREST(rsDocNoEnv!Documento, 2, DateTime.Date) Then
                            End If
                        End If
                        rsDocNoEnv.MoveNext
                        Me.lblProgreso.Caption = "Enviando Documento.........." & Me.pgbEnvio.value + 1 & " DE " & Me.pgbEnvio.Max
                        Sleep 2000
                    Loop
                    cmdPape_Click (0)
                Else
                    Me.frmDoc.Visible = False
                End If
                Me.lblProgreso.Visible = False
                Me.pgbEnvio.Visible = False
                Screen.MousePointer = vbDefault
            Else 'INFOFACT
                Screen.MousePointer = vbHourglass
                Me.pgbEnvio.value = 0
                Me.pgbEnvio.Min = 0
                Me.pgbEnvio.Max = rsDocNoEnv.RecordCount
                Me.pgbEnvio.Visible = True
                Me.lblProgreso.Visible = True
                If rsDocNoEnv.RecordCount > 0 Then
                    rsDocNoEnv.MoveFirst
                    Me.lblProgreso.Caption = "Enviando Documento...: "
                    Do While Not rsDocNoEnv.EOF
                        DoEvents
                        'Dim sd As String
                        Me.pgbEnvio.value = Me.pgbEnvio.value + 1
                        If rsDocNoEnv!Doc = "D" Then
                            If Not INSERTAFE(rsDocNoEnv!Documento, "", 1, "") Then
                            End If
                        ElseIf rsDocNoEnv!Doc = "N" Then
                            If Not INSERTAFE(rsDocNoEnv!Documento, "", 2, "") Then
                            End If
                        End If
                        rsDocNoEnv.MoveNext
                        Me.lblProgreso.Caption = "Enviando Documento.........." & Me.pgbEnvio.value + 1 & " DE " & Me.pgbEnvio.Max
                        Sleep 500
                    Loop
                    cmdPape_Click (0)
                Else
                    Me.frmDoc.Visible = False
                End If
                Me.lblProgreso.Visible = False
                Me.pgbEnvio.Visible = False
                Screen.MousePointer = vbDefault
            End If
        Case Is = 2
            cmdPape_Click (1)
            Me.frmDoc.Visible = False
    End Select
    Exit Sub
fin:
    Screen.MousePointer = vbDefault
    MsgBox "Se genero un Inconveniente, favor de Refrescar los valores!!" & vbNewLine + Error, vbInformation
End Sub

Private Sub cmdPropina_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nTarjetaPropina(Index) = IIf(wEnter = True, sDescrip, nTarjetaPropina(Index))
    txtPropina(Index).Caption = Format(nTarjetaPropina(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub cmdTarjeta_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nTarjeta(Index) = IIf(wEnter = True, sDescrip, nTarjeta(Index))
    txtTarjeta(Index).Caption = Format(nTarjeta(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub Form_Activate()
    If pais = "000" And (lFEpape Or lFEBiz) And FEenvio = True Then
        Me.frmDoc.Visible = True
        cmdPape_Click (0)
        FEenvio = False
    Else
        If lFacturacionE = True And lFEBiz = False And lFEOfisis = False And lFEpape = False And lFECarbajal = False And lFESpring = False And lFEGesa = False And FEenvio = True Then
            Me.frmDoc.Visible = True
            cmdPape_Click (0)
            FEenvio = False
        Else
            Me.frmDoc.Visible = False
        End If
    End If
    If lMCPV Then
        If Calcular("select count(tDocumento) as codigo from MDOCUMENTO where tEstadoDocumento ='01' and tUsuario ='" & sUsuario & "'", Cn) > 0 Then
            MsgBox "Tienes Documentos por Cancelar", vbCritical, sMensaje
            Unload Me
        End If
    Else
        If Calcular("select count(tDocumento) as codigo from MDOCUMENTO where tEstadoDocumento ='01' and tCaja ='" & sCaja & "'", Cn) > 0 Then
            MsgBox "Tienes Documentos por Cancelar", vbCritical, sMensaje
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
   Centrar Me
   Screen.MousePointer = vbHourglass
   
   frmDoc.Top = 0
   frmDoc.Left = 0
   
   
   If lMCPV Then
        If Calcular("select count(tCodigoPedido) as codigo from vPedidoDetalle where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tUsuario = '" & sUsuario & "')", Cn) > 0 Then
           MsgBox "Tienes Pedidos Abiertos", vbExclamation, sMensaje
        End If
                   
        If Calcular("select count(tCodigoPedido) as codigo from vPedidoDetalle where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tUsuario <> '" & sUsuario & "')", Cn) > 0 Then
           MsgBox "Existen Pedidos Abiertos de otros Usuarios", vbExclamation, sMensaje
        End If
   Else
        If Calcular("select count(tCodigoPedido) as codigo from vPedidoDetalle where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tCaja = '" & sCaja & "')", Cn) > 0 Then
           MsgBox "Tienes Pedidos Abiertos", vbExclamation, sMensaje
        End If
                   
        If Calcular("select count(tCodigoPedido) as codigo from vPedidoDetalle where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO')", Cn) > 0 Then
           MsgBox "Tienes Pedidos Abiertos del Punto de Ventas de Mozos", vbExclamation, sMensaje
        End If
   End If
   txtUsuario.Caption = sUsuario
   txtFecha.Caption = FechaServidor()
   Me.Caption = "Cierre Turno " & sTurno
   SSTab.Tab = 0
   
   If lCierre Then
      cmdOpcion(4).Enabled = False
      cmdOpcion(5).Enabled = False
      lblReciboN.Visible = False
      lblReciboE.Visible = False
      lblReciboE2.Visible = False
      lblReciboE3.Visible = False
   End If
        
   'Descripciones MN y ME
   txtMN.Caption = sMonN
   lblMontoN(0).Caption = sMonN
   lblMontoN(1).Caption = sMonN
   cmdMovimiento(0).Caption = "Ventas Efectivo " & sMonN
   cmdMovimiento(4).Caption = "Otros Tipos Pago "
   cmdMovimiento(6).Caption = "Cheques " & sMonN
   cmdMovimiento(8).Caption = "Puntos " & sMonN
   cmdMovimiento(10).Caption = "Retiro Efectivo " & sMonN
   
   If sMonE <> "" And sMonN <> sMonE Then
      txtME.Caption = sMonE
      lblMontoE(0).Caption = sMonE
      lblMontoE(1).Caption = sMonE
      cmdMovimiento(1).Caption = "Ventas Efectivo " & sMonE
      cmdMovimiento(7).Caption = "Cheques " & sMonE
      cmdMovimiento(11).Caption = "Retiro Efectivo " & sMonE
   Else
      txtME.Visible = False
      lblMontoE(0).Visible = False
      lblMontoE(1).Visible = False
      txtInicioE.Visible = False
      txtEfectivoE.Visible = False
      txtEfectivoE.Visible = False
      txtSaldoE.Visible = False
      txtTotalOtroE.Visible = False
      txtChequeE.Visible = False
      txtPuntoE.Visible = False
      txtTotalIngresoE.Visible = False
      txtTotalEfectivoE.Visible = False
      cmdMovimiento(1).Visible = False
      cmdMovimiento(7).Visible = False
      cmdMovimiento(11).Visible = False
      
      cmdOtroE(1).Visible = False
      cmdOtroE(2).Visible = False
      cmdOtroE(3).Visible = False
      cmdOtroE(4).Visible = False
      cmdOtroE(5).Visible = False
      cmdOtroE(6).Visible = False
      cmdOtroE(7).Visible = False
      cmdOtroE(8).Visible = False
      cmdOtroE(9).Visible = False
      cmdOtroE(10).Visible = False
      
      'LG
      cmdOtroE(11).Visible = False
      cmdOtroE(12).Visible = False
      cmdOtroE(13).Visible = False
      cmdOtroE(14).Visible = False
      cmdOtroE(15).Visible = False
      cmdOtroE(16).Visible = False
      cmdOtroE(17).Visible = False
      cmdOtroE(18).Visible = False
      cmdOtroE(19).Visible = False
      cmdOtroE(20).Visible = False
      
      txtOtroE(1).Visible = False
      txtOtroE(2).Visible = False
      txtOtroE(3).Visible = False
      txtOtroE(4).Visible = False
      txtOtroE(5).Visible = False
      txtOtroE(6).Visible = False
      txtOtroE(7).Visible = False
      txtOtroE(8).Visible = False
      txtOtroE(9).Visible = False
      txtOtroE(10).Visible = False
      'LG
      txtOtroE(11).Visible = False
      txtOtroE(12).Visible = False
      txtOtroE(13).Visible = False
      txtOtroE(14).Visible = False
      txtOtroE(15).Visible = False
      txtOtroE(16).Visible = False
      txtOtroE(17).Visible = False
      txtOtroE(18).Visible = False
      txtOtroE(19).Visible = False
      txtOtroE(20).Visible = False
   End If
                        
   If pais = "003" Then
    txtME2.Caption = sMonE2
    txtME3.Caption = sMonE3
    frameOME.BorderStyle = 0
    cmdOpcion(0).Caption = "Cierre X"
    cmdOpcion(3).Visible = True
    cmdOpcion(6).Visible = True
   Else
    frameOME.Visible = False
    'txtME2.Visible = False
   End If
   'Fecha del Turno
   xFecha = Calcular("select fInicial as codigo from MTURNO where tTurno='" & sTurno & "'", Cn)
      
   'Tipo de Cambio del turno
   nCambio = Calcular("select nVenta as Codigo from TTIPOCAMBIO where fFecha='" & Format(xFecha, "yyyy/mm/dd") & "'", Cn)
   nCambio2 = Calcular("select isnull(nVenta2,0) as Codigo from TTIPOCAMBIO where fFecha='" & Format(xFecha, "yyyy/mm/dd") & "'", Cn)
   nCambio3 = Calcular("select isnull(nVenta3,0) as Codigo from TTIPOCAMBIO where fFecha='" & Format(xFecha, "yyyy/mm/dd") & "'", Cn)
   
   txtObservacion.Text = Calcular("select isnull(tObservacion,'') codigo from MTURNO where tturno='" & sTurno & "'", Cn)
   
   'Saldo Inicial
   Set RsSumas = Lib.OpenRecordset("select * from MTURNO where tTurno ='" & sTurno & "'", Cn)
   If RsSumas.RecordCount = 0 Then
      nInicioN = 0
      nInicioE = 0
   Else
      nInicioN = RsSumas!nMontoIN
      nInicioE = RsSumas!nMontoIE
   End If
   
   LlenaTarjeta
   LlenaOtro
   Asignar
   
   
   nTotalIngresoN = nEfectivoN + nTotalTarjeta + nChequeN + nTotalOtroN + nPuntoN + nEgresoN
   nTotalIngresoE = nEfectivoE + nChequeE + nTotalOtroE + nEgresoE
   nTotalIngresoE2 = nEfectivoE2 + nEgresoE2
   nTotalIngresoE3 = nEfectivoE3 + nEgresoE3
   nTotalEfectivoN = nEfectivoN
   
   nTotalEfectivoE = nEfectivoE
   nTotalEfectivoE2 = nEfectivoE2
   nTotalEfectivoE3 = nEfectivoE3
   
   nFinalN = nTotalEfectivoN - nSaldoN
   nFinalE = nTotalEfectivoE - nSaldoE
   nFinalE2 = nTotalEfectivoE2 ' - nSaldoE
   nFinalE3 = nTotalEfectivoE3 ' - nSaldoE
   
   'Llena Pantalla
   lblReciboN.Caption = Format(nEgresoN, "###,###,##0.00")
   lblReciboE.Caption = Format(nEgresoE, "###,###,##0.00")
   lblReciboE2.Caption = Format(nEgresoE2, "###,###,##0.00")
   lblReciboE3.Caption = Format(nEgresoE3, "###,###,##0.00")
   
   txtInicioN.Caption = Format(nInicioN, "###,###,##0.00")
   txtInicioE.Caption = Format(nInicioE, "###,###,##0.00")
   txtEfectivoN.Caption = Format(nEfectivoN, "###,###,##0.00")
   txtEfectivoE.Caption = Format(nEfectivoE, "###,###,##0.00")
   txtEfectivoE2.Caption = Format(nEfectivoE2, "###,###,##0.00")
   txtEfectivoE3.Caption = Format(nEfectivoE3, "###,###,##0.00")
   txtTotalTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
   txtTotalTarjetaPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
   txtChequeN.Caption = Format(nChequeN, "###,###,##0.00")
   txtChequeE.Caption = Format(nChequeE, "###,###,##0.00")
   txtTotalOtroN.Caption = Format(nTotalOtroN, "###,###,##0.00")
   txtTotalOtroE.Caption = Format(nTotalOtroE, "###,###,##0.00")
   txtPuntoN.Caption = Format(nPuntoN, "###,###,##0.00")
   txtPuntoE.Caption = Format(nPuntoE, "###,###,##0.00")
   
   txtTotalIngresoN.Caption = Format(nTotalIngresoN, "###,###,##0.00")
   txtTotalIngresoE.Caption = Format(nTotalIngresoE, "###,###,##0.00")
   txtTotalIngresoE2.Caption = Format(nTotalIngresoE2, "###,###,##0.00")
   txtTotalIngresoE3.Caption = Format(nTotalIngresoE3, "###,###,##0.00")
   txtTotalEfectivoN.Caption = Format(nTotalEfectivoN, "###,###,##0.00")
   txtTotalEfectivoE.Caption = Format(nTotalEfectivoE, "###,###,##0.00")
   txtTotalEfectivoE2.Caption = Format(nTotalEfectivoE2, "###,###,##0.00")
   txtTotalEfectivoE3.Caption = Format(nTotalEfectivoE3, "###,###,##0.00")
   txtSaldoN.Caption = Format(nFinalN, "###,###,##0.00")
   txtSaldoE.Caption = Format(nFinalE, "###,###,##0.00")
   txtSaldoE2.Caption = Format(nFinalE2, "###,###,##0.00")
   txtSaldoE3.Caption = Format(nFinalE3, "###,###,##0.00")
   
   txtRetiroN.Caption = Format(0, "###,###,##0.00")
   txtRetiroE.Caption = Format(0, "###,###,##0.00")
    
    If lFEBiz Then
        Call ConfGrilla(6, grdGrilla, "Prefijo", 2, "Pref", 600, 2, 0, "", _
             "Documento", 2, "Documento", 1400, 0, 0, "", _
             "Cliente", 2, "Cliente", 4000, 0, 0, "", _
             "Monto", 2, "Monto", 700, 0, 0, "", _
             "Usuario", 2, "Usuario", 1500, 0, 0, "", _
             "Turno", 2, "Turno", 1000, 0, 0, "")
   Else
        Call ConfGrilla(6, grdGrilla, "Prefijo", 2, "Pref", 600, 2, 0, "", _
             "Documento", 2, "Documento", 1400, 0, 0, "", _
             "Cliente", 2, "Cliente", 4000, 0, 0, "", _
             "Monto", 2, "Monto", 700, 0, 0, "", _
             "Usuario", 2, "Usuario", 1500, 0, 0, "", _
             "Turno", 2, "Turno", 1000, 0, 0, "")
   End If
   FEenvio = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmLiquidacionDetalle = Nothing
End Sub

Public Sub LlenaTarjeta()
   'GCAA 23/03/2020
   Set RsTarjetas = Lib.OpenRecordset("select tCodigoTarjeta, tDetallado, tResumido from tTarjetaCredito where lActivo =1", Cn)
   'canttarjetas = RsTarjetas.RecordCount
   If RsTarjetas.RecordCount = 0 Then
    MsgBox "No hay tarjetas configuradas, favor de revisar las configuraciones!!!", vbInformation
    Exit Sub
   End If
   
   For i = 1 To 24
       RsTarjetas.MoveFirst
       
       If i > 9 Then
        RsTarjetas.Find "tcodigoTarjeta ='" & Trim(str(i)) & " '"
       Else
        RsTarjetas.Find "tcodigoTarjeta = '0" & Trim(str(i)) & "'"
       End If
       
       If RsTarjetas.EOF Then
         If i <= 24 Then
          lblTarjeta(i).Visible = False
          cmdTarjeta(i).Visible = False
          cmdPropina(i).Visible = False
          txtTarjeta(i).Visible = False
          txtPropina(i).Visible = False
          End If
       Else
'            If i > 7 Then
'               lblTarjeta(8).Caption = "OTRAS TARJETAS"
'               cmdTarjeta(8).Caption = "OTRAS TARJETAS"
'               cmdPropina(8).Caption = "OTRAS TARJETAS"
'               tTarjeta(8) = "zzz"
'               If Not lCierre Then
'                  cmdTarjeta(8).Enabled = False
'                  cmdPropina(8).Enabled = False
'               End If
'            Else
               lblTarjeta(i).Caption = RsTarjetas!tDetallado
               cmdTarjeta(i).Caption = RsTarjetas!tResumido
               cmdPropina(i).Caption = RsTarjetas!tResumido
               tTarjeta(i) = RsTarjetas!tCodigoTarjeta
               If Not lCierre Then
                  cmdTarjeta(i).Enabled = False
                  cmdPropina(i).Enabled = False
               End If
           ' End If
       End If
   Next i
   
End Sub

Public Sub CalcularTotales()
   'Calcula Tarjeta
   nTotalTarjeta = 0
   nTotalTarjetaPropina = 0
   nTotalOtroN = 0
   nTotalOtroE = 0
   For i = 1 To 24
       nTotalTarjeta = nTotalTarjeta + nTarjeta(i)
       nTotalTarjetaPropina = nTotalTarjetaPropina + nTarjetaPropina(i)
   Next i
   
   For i = 1 To 20
       nTotalOtroN = nTotalOtroN + nOtroN(i)
       nTotalOtroE = nTotalOtroE + nOtroE(i)
   Next i
   
   txtTTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
   txtTPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
   txtTotalTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
   txtTotalTarjetaPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
   
   txtTotalOtroN.Caption = Format(nTotalOtroN, "###,###,##0.00")
   txtTotalOtroE.Caption = Format(nTotalOtroE, "###,###,##0.00")
'   txtOtroN(9).Caption = Format(nTotalOtroN, "###,###,##0.00")
'   txtOtroE(9).Caption = Format(nTotalOtroE, "###,###,##0.00")
      
   'Total de ingreso
   nTotalIngresoN = nEfectivoN + nTotalTarjeta + nChequeN + nTotalOtroN + nPuntoN + nEgresoN
   nTotalIngresoE = nEfectivoE + nChequeE + nTotalOtroE + nEgresoE
   
   'Total de Efectivo
   nTotalEfectivoN = nEfectivoN
   nTotalEfectivoE = nEfectivoE
   nTotalEfectivoE2 = nEfectivoE2
   nTotalEfectivoE3 = nEfectivoE3
   
   nFinalN = nTotalEfectivoN - nRetiroN
   nFinalE = nTotalEfectivoE - nRetiroE
   nFinalE2 = nTotalEfectivoE2 '- nRetiroE
   nFinalE3 = nTotalEfectivoE3 '- nRetiroE

   'Saldo en Efectivo
   txtTotalIngresoN.Caption = Format(nTotalIngresoN, "###,###,##0.00")
   txtTotalIngresoE.Caption = Format(nTotalIngresoE, "###,###,##0.00")
   txtTotalEfectivoN.Caption = Format(nTotalEfectivoN, "###,###,##0.00")
   txtTotalEfectivoE.Caption = Format(nTotalEfectivoE, "###,###,##0.00")
   txtSaldoN.Caption = Format(nFinalN, "###,###,##0.00")
   txtSaldoE.Caption = Format(nFinalE, "###,###,##0.00")
   txtSaldoE2.Caption = Format(nFinalE2, "###,###,##0.00")
   txtSaldoE3.Caption = Format(nFinalE3, "###,###,##0.00")
End Sub

Public Sub ImprimeLiquidacion()
   Dim sTitulo1 As String
   Dim sTitulo2 As String
   Dim sTitulo3 As String
   Dim sLinea As String
   Screen.MousePointer = vbHourglass
   'Configura la impresora la impresion Font
   Imprimir (sPreCuenta)
   Printer.FontName = sFont
   Printer.FontBold = False
   
   'Cabecera
   ImprimeXCentro "Liquidación de Cajero", 40
   ImprimeXCentro sRazonSocial, 40
   Printer.Print ""
      
   sTitulo1 = ""
   sTitulo2 = ""
   sTitulo1 = "Turno   : " & sTurno
   sTitulo2 = "Fecha   : " & Format(Now, "dd MMMM yyyy HH:mm") & " Hrs "
   
   Printer.Print sTitulo1
   Printer.Print sTitulo2
   Printer.Print ""
   ImprimeXCentro "(Según Cajero)", 40
   
   Printer.Print String(40, "-")
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "CONCEPTO            " & "Mon MN" & "        " & "Mon ME"
   Else
      Printer.Print "CONCEPTO            " & "Mon MN"
   End If
   
   sLinea = ""
   sLinea = String(40, "-")
   '-------------------------------------------------------------------------------------------------------------------------------------------------
   'Fondo de Caja-------------------------------------------------------------------------------------------------------------------------------------
   Printer.Print
   If sMonE <> "" And sMonN <> sMonE Then
      sLinea = sLinea & vbNewLine & "Fondo de Caja   : " & Right(String(11, " ") & Format(nInicioN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nInicioE, "###,##0.00"), 11)
   Else
      sLinea = sLinea & vbNewLine & "Fondo de Caja   : " & Right(String(11, " ") & Format(nInicioN, "###,##0.00"), 11)
   End If
      
   'Efectivo
   sLinea = sLinea & vbNewLine & ""
   If pais = "003" Then
        sLinea = sLinea & vbNewLine & "  Efectivo(" & Right(String(4, " ") & sMonN, 4) & "): " & Right(String(11, " ") & Format(nEfectivoN, "###,##0.00"), 11) & Right(String(11, " ") & "", 11)
        sLinea = sLinea & vbNewLine & "  Efectivo(" & Right(String(4, " ") & sMonE, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE, "###,##0.00"), 11)
        sLinea = sLinea & vbNewLine & "  Efectivo(" & Right(String(4, " ") & sMonE2, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE2, "###,##0.00"), 11)
        sLinea = sLinea & vbNewLine & "  Efectivo(" & Right(String(4, " ") & sMonE3, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE3, "###,##0.00"), 11)
   Else
    If sMonE <> "" And sMonN <> sMonE Then
       sLinea = sLinea & vbNewLine & "       Efectivo : " & Right(String(11, " ") & Format(nEfectivoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nEfectivoE, "###,##0.00"), 11)
    Else
       sLinea = sLinea & vbNewLine & "       Efectivo : " & Right(String(11, " ") & Format(nEfectivoN, "###,##0.00"), 11)
    End If
   End If
   'Total de Tarjetas
   'GCAA
   sLinea = sLinea & vbNewLine & ""
   sLinea = sLinea & vbNewLine & "Tarjetas Credito: " & Right(String(11, " ") & Format(nTotalTarjeta, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalTarjetaPropina, "###,##0.00"), 11)
   For i = 1 To 24
       If lblTarjeta(i).Visible = True Then
          sLinea = sLinea & vbNewLine & " - " & Mid(cmdTarjeta(i).Caption & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nTarjeta(i), "###,##0.00"), 11) & Right(String(11, " ") & Format(nTarjetaPropina(i), "###,##0.00"), 11)
       End If
   Next i

   'Total de Cheques
   sLinea = sLinea & vbNewLine & " "
   If sMonE <> "" And sMonN <> sMonE Then
      sLinea = sLinea & vbNewLine & "Cheque/Deposito : " & Right(String(11, " ") & Format(nChequeN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nChequeE, "###,##0.00"), 11)
   Else
      sLinea = sLinea & vbNewLine & "Cheque/Deposito : " & Right(String(11, " ") & Format(nChequeN, "###,##0.00"), 11)
   End If

   'Otros Tipos de Pago
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & "Otros Tipos Pago: " & Right(String(11, " ") & Format(nTotalOtroN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalOtroE, "###,##0.00"), 11)
   For i = 1 To 20
       If sMonE <> "" And sMonN <> sMonE Then
          If lblOtro(i).Visible = True Then
             sLinea = sLinea & vbNewLine & " - " & Mid(lblOtro(i).Caption & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nOtroN(i), "###,##0.00"), 11) & Right(String(11, " ") & Format(nOtroE(i), "###,##0.00"), 11)
          End If
       Else
          If lblOtro(i).Visible = True Then
             sLinea = sLinea & vbNewLine & " - " & Mid(lblOtro(i).Caption & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nOtroN(i), "###,##0.00"), 11)
          End If
       End If
   Next i
 
   'Puntos
   sLinea = sLinea & vbNewLine & " "
   If sMonE <> "" And sMonN <> sMonE Then
      sLinea = sLinea & vbNewLine & "Puntos          : " & Right(String(11, " ") & Format(nPuntoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nPuntoE, "###,##0.00"), 11)
   Else
      sLinea = sLinea & vbNewLine & "Puntos          : " & Right(String(11, " ") & Format(nPuntoN, "###,##0.00"), 11)
   End If

   'Total Ingreso en Caja
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & String(40, "-")
   sLinea = sLinea & vbNewLine & " "
   If pais = "003" Then
    sLinea = sLinea & vbNewLine & "T. Ingreso(" & Right(String(4, " ") & sMonN, 4) & "): " & Right(String(11, " ") & Format(nTotalIngresoN, "###,##0.00"), 11) & Right(String(11, " ") & "", 11)
    sLinea = sLinea & vbNewLine & "T. Ingreso(" & Right(String(4, " ") & sMonE, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nTotalIngresoE, "###,##0.00"), 11)
    sLinea = sLinea & vbNewLine & "T. Ingreso(" & Right(String(4, " ") & sMonE2, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE2, "###,##0.00"), 11)
    sLinea = sLinea & vbNewLine & "T. Ingreso(" & Right(String(4, " ") & sMonE3, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE3, "###,##0.00"), 11)
   Else
    If sMonE <> "" And sMonN <> sMonE Then
       sLinea = sLinea & vbNewLine & "Total Ingreso   : " & Right(String(11, " ") & Format(nTotalIngresoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalIngresoE, "###,##0.00"), 11)
    Else
       sLinea = sLinea & vbNewLine & "Total Ingreso   : " & Right(String(11, " ") & Format(nTotalIngresoN, "###,##0.00"), 11)
    End If
   End If

   'Total Efectivo en Caja
   sLinea = sLinea & vbNewLine & " "
   If pais = "003" Then
    sLinea = sLinea & vbNewLine & "T. Efectivo(" & Right(String(3, " ") & sMonN, 3) & "): " & Right(String(11, " ") & Format(nTotalEfectivoN, "###,##0.00"), 11) & Right(String(11, " ") & "", 11)
    sLinea = sLinea & vbNewLine & "T. Efectivo(" & Right(String(3, " ") & sMonE, 3) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nTotalEfectivoE, "###,##0.00"), 11)
    sLinea = sLinea & vbNewLine & "T. Efectivo(" & Right(String(3, " ") & sMonE2, 3) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE2, "###,##0.00"), 11)
    sLinea = sLinea & vbNewLine & "T. Efectivo(" & Right(String(3, " ") & sMonE3, 3) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE3, "###,##0.00"), 11)
   Else
    If sMonE <> "" And sMonN <> sMonE Then
       sLinea = sLinea & vbNewLine & "Total Efectivo  : " & Right(String(11, " ") & Format(nTotalEfectivoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalEfectivoE, "###,##0.00"), 11)
    Else
       sLinea = sLinea & vbNewLine & "Total Efectivo  : " & Right(String(11, " ") & Format(nTotalEfectivoN, "###,##0.00"), 11)
    End If
   End If

   'Retiro de Caja
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      sLinea = sLinea & vbNewLine & "Retiro Efectivo : " & Right(String(11, " ") & Format(nRetiroN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nRetiroE, "###,##0.00"), 11)
   Else
      sLinea = sLinea & vbNewLine & "Retiro Efectivo : " & Right(String(11, " ") & Format(nRetiroN, "###,##0.00"), 11)
   End If

   'Saldo de Caja
   sLinea = sLinea & vbNewLine & " "
   If pais = "003" Then
    sLinea = sLinea & vbNewLine & "Saldo Caja(" & Right(String(4, " ") & sMonN, 4) & "): " & Right(String(11, " ") & Format(nTotalEfectivoN - nRetiroN, "###,##0.00"), 11) & Right(String(11, " ") & "", 11)
    sLinea = sLinea & vbNewLine & "Saldo Caja(" & Right(String(4, " ") & sMonE, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nTotalEfectivoE - nRetiroE, "###,##0.00"), 11)
    sLinea = sLinea & vbNewLine & "Saldo Caja(" & Right(String(4, " ") & sMonE2, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE2, "###,##0.00"), 11)
    sLinea = sLinea & vbNewLine & "Saldo Caja(" & Right(String(4, " ") & sMonE3, 4) & "): " & Right(String(11, " ") & "", 11) & Right(String(11, " ") & Format(nEfectivoE3, "###,##0.00"), 11)
   Else
    If sMonE <> "" And sMonN <> sMonE Then
       sLinea = sLinea & vbNewLine & "Saldo en Caja   : " & Right(String(11, " ") & Format(nTotalEfectivoN - nRetiroN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalEfectivoE - nRetiroE, "###,##0.00"), 11)
    Else
       sLinea = sLinea & vbNewLine & "Saldo en Caja   : " & Right(String(11, " ") & Format(nTotalEfectivoN - nRetiroN, "###,##0.00"), 11)
    End If
   End If
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & String(40, "-")
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & String(25, " ") & "---------------"
   sLinea = sLinea & vbNewLine & String(25, " ") & String((15 - Len(RTrim(sUsuario))) / 2, " ") & sUsuario
   sLinea = sLinea & vbNewLine & " "
   sLinea = sLinea & vbNewLine & String(40, "-")
   
   If pais = "003" Then
    sLinea = Replace(Replace(Replace(sLinea, ",", "|"), ".", ","), "|", ".")
   End If
   
   Printer.Print sLinea
   '--------------------------------------------------------------------------------------------------------------------------------------
   '---------------------------------------------------------------------------------------------------------------------------------------
   Dim sDoc As String
   Dim sTipoDoc As String
   Dim nCantidadDoc As Integer
   
   'For i = 1 To Calcular("select count(tCodigo) as codigo From TTABLA Where TTABLA = 'TIPODOCUMENTO'", Cn)
   For i = 1 To Calcular("select count(tCodigoTIPODOCUMENTO) as codigo From TTIPODOCUMENTO ", Cn)
       sTipoDoc = Mid("00", 1, 2 - Len(Trim(str(i)))) & Trim(str(i))
      ' sDoc = Calcular("Select tDetallado As Codigo from TTABLA Where TTABLA = 'TIPODOCUMENTO' AND tCODIGO = '" & sTipoDoc & "'", Cn)
       sDoc = Calcular("Select TDESCRIPCION As Codigo from TTIPODOCUMENTO Where   tCODIGOTIPODOCUMENTO = '" & sTipoDoc & "'", Cn)
       nCantidadDoc = Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn)
       If sDoc <> "0" And nCantidadDoc > 0 Then
          Printer.Print " "
          Printer.Print sDoc
          Printer.Print "Del     : " & Format(Calcular("select min(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Al      : " & Format(Calcular("select max(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Emitido : " & Format(Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tEstadoDocumento<>'04' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Anulado : " & Format(Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tEstadoDocumento='04' and tTurno='" & sTurno & "'", Cn), "##,##0")
       End If
   Next i
   Printer.Print ""
   If Len(Trim(txtObservacion.Text)) > 0 Then
      Printer.Print String(40, "-")
      ImprimeXLinea txtObservacion.Text, 40, 0
   End If
   Printer.Print String(40, "-")
   Printer.Print "Caja    : " & sCaja
   Printer.Print " "
   Printer.EndDoc
   Screen.MousePointer = vbDefault
End Sub

Public Sub LlenaOtro()
   Set RsOtros = Lib.OpenRecordset("select * from vTipoCancelacion where lActivo=1", Cn)
   If RsOtros.RecordCount = 0 Then
      cmdMovimiento(4).Enabled = False
      Exit Sub
   End If
   RsOtros.MoveFirst
   
   For i = 1 To 20
       If RsOtros.EOF Then
          lblOtro(i).Visible = False
          cmdOtroN(i).Visible = False
          cmdOtroE(i).Visible = False
          txtOtroN(i).Visible = False
          txtOtroE(i).Visible = False
       Else
          lblOtro(i).Caption = RsOtros!descripcion
          cmdOtroN(i).Caption = RsOtros!tResumido
          cmdOtroE(i).Caption = RsOtros!tResumido
          tOtro(i) = RsOtros!codigo
          If Not lCierre Then
             cmdOtroN(i).Enabled = False
             cmdOtroE(i).Enabled = False
          End If
          RsOtros.MoveNext
       End If
   Next i
      
End Sub

Public Sub Asignar()
   nDolar = 0
   nDolar2 = 0
   nDolar3 = 0
   nEfectivoN = 0
   nEfectivoE = 0
   nEfectivoE2 = 0
   nEfectivoE3 = 0
   nTotalTarjeta = 0
   nTotalTarjetaPropina = 0
   nChequeN = 0
   nChequeE = 0
   nTotalOtroN = 0
   nTotalOtroE = 0
   nPuntoN = 0
   nPuntoE = 0
   
   Dim nIngresoTarjeta As Double
   Dim RsIngresoTarjeta As Recordset
   Dim nIngresoTarjetaDolares As Double
   
   
   If Not lCierre Then
      'Recibos de Egreso
      nEgresoN = 0
      nEgresoE = 0
      Isql = "select tMoneda, sum(nMonto) as nMonto From megreso where tEstadoDocumento='01' and tTurno='" & sTurno & "' Group by tMoneda"
      Set RsEgreso = Lib.OpenRecordset(Isql, Cn)
      If RsEgreso.RecordCount > 0 Then
         Do While Not RsEgreso.EOF
            If RsEgreso!tMONEDA = "01" Then
               nEgresoN = IIf(IsNull(RsEgreso!nMonto), 0, RsEgreso!nMonto)
            ElseIf RsEgreso!tMONEDA = "02" Then
               nEgresoE = IIf(IsNull(RsEgreso!nMonto), 0, RsEgreso!nMonto)
            ElseIf RsEgreso!tMONEDA = "03" Then
               nEgresoE2 = IIf(IsNull(RsEgreso!nMonto), 0, RsEgreso!nMonto)
            ElseIf RsEgreso!tMONEDA = "04" Then
               nEgresoE3 = IIf(IsNull(RsEgreso!nMonto), 0, RsEgreso!nMonto)
            Else
               nEgresoE = IIf(IsNull(RsEgreso!nMonto), 0, RsEgreso!nMonto)
            End If
            RsEgreso.MoveNext
         Loop
      End If

      'Recibos de Ingreso
      nIngresoN = 0
      nIngresoE = 0
      Isql = "select tMoneda, sum(nMonto) as nMonto From mIngreso where tEstadoDocumento='01' and tTurno='" & sTurno & "' and tTipoPago='01' Group by tMoneda "
            
      Set RsIngreso = Lib.OpenRecordset(Isql, Cn)
      If RsIngreso.RecordCount > 0 Then
         Do While Not RsIngreso.EOF
            If RsIngreso!tMONEDA = "01" Then
               nIngresoN = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
            Else
               nIngresoE = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
            End If
            RsIngreso.MoveNext
         Loop
      End If
      
      
      'Isql = "select tMoneda, sum(nMonto) as nMonto From mIngreso where tEstadoDocumento='01' and tTurno='" & sTurno & "' and tTipoPago='02' Group by tMoneda "
      'Set RsIngresoTarjeta = Lib.OpenRecordset(Isql, Cn)
      'If RsIngresoTarjeta.RecordCount > 0 Then
      '   Do While Not RsIngresoTarjeta.EOF
      '      If RsIngresoTarjeta!tMoneda = "01" Then
      '         nIngresoTarjeta = IIf(IsNull(RsIngresoTarjeta!nMonto), 0, RsIngresoTarjeta!nMonto)
      '      Else
      '         nIngresoTarjetaDolares = IIf(IsNull(RsIngresoTarjeta!nMonto), 0, RsIngresoTarjeta!nMonto)
      '      End If
      '      RsIngresoTarjeta.MoveNext
      '   Loop
      'End If
      
   If lNcOfisis Then

       Isql = " select tTipoPago, tMoneda, isnull(tTarjeta,'')tTarjeta, isnull(tOtroTipoPago,'')tOtroTipoPago," & _
                "sum(case when ISNULL( DBO.MNOTACREDITO.nVenta,0)>0 and tMoneda='01' then 0 else  case when ISNULL( DBO.MNOTACREDITO.nVenta,0)>0 and tMoneda='02'  then 0 else nMOnto end  end )  as nMonto," & _
                "sum(nPropina) as nPropina, sum(nPropina) as nPropina, sum(case when ISNULL( DBO.MNOTACREDITO.nVenta,0)>0 and tMoneda='01' then 0 else  case when ISNULL( DBO.MNOTACREDITO.nVenta,0)>0 and tMoneda='02'then 0 else nDolar end  end )  as nDolar " & _
                "from dpagodocumento LEFT OUTER JOIN MNOTACREDITO ON  DBO.DPAGODOCUMENTO.tDocumento=DBO.MNOTACREDITO.tDocumento AND DBO.MNOTACREDITO.tTurno='" & sTurno & "' and dbo.MNOTACREDITO.tEstadoDocumento in ('05','02') " & _
                "and DBO.MNOTACREDITO.nVenta=(select dbo.MDOCUMENTO.nVenta from MDOCUMENTO where dbo.MDOCUMENTO.tDocumento=  DBO.DPAGODOCUMENTO.tDocumento) " & _
                " where dpagodocumento.tTurno='" & sTurno & "'   group by tTipoPago, tMoneda, isnull(tTarjeta,''), isnull(tOtroTipoPago,'')"
                
    Else
      Isql = "select tTipoPago, tMoneda, isnull(tTarjeta,'')tTarjeta, isnull(tOtroTipoPago,'')tOtroTipoPago, sum(nMonto) as nMonto, sum(nPropina) as nPropina, sum(nDolar) as nDolar from dpagodocumento " & _
             "where tTurno='" & sTurno & "' group by tTipoPago, tMoneda, isnull(tTarjeta,''), isnull(tOtroTipoPago,'')"
             
   
    End If
             '& _
'             "UNION " & _
'             "select '01', tMoneda, '', '', sum(nMonto) as nMonto, 0 as nPropina, 0 as nDolar from mIngreso " & _
'             "where tEstadoDocumento='01' and tTurno='" & sTurno & "' and tTipoPago='01' Group by tTarjeta, tMoneda " & _
'             "UNION " & _
'             "select '02', tMoneda, tTarjeta, '', sum(nMonto) as nMonto, 0 as nPropina, 0 as nDolar from mIngreso " & _
'             "where tEstadoDocumento='01' and tTurno='" & sTurno & "' and tTipoPago='02' Group by tTarjeta, tMoneda "
' estas lineas estan de mas.. x eso estan como comentarios.
             
      Set RsPago = Lib.OpenRecordset(Isql, Cn)
      If RsPago.RecordCount > 0 Then
         cmdMovimiento(0).Enabled = False
         cmdMovimiento(1).Enabled = False
         cmdMovimiento(6).Enabled = False
         cmdMovimiento(7).Enabled = False
         cmdMovimiento(8).Enabled = False
         RsPago.MoveFirst
         Do While Not RsPago.EOF
            Select Case RsPago!tTipoPago
                   Case Is = "01"  'Efectivo
                        If RsPago!tMONEDA = "01" Then
                           nEfectivoN = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        ElseIf RsPago!tMONEDA = "02" Then
                           nEfectivoE = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                           nDolar = IIf(IsNull(RsPago!nDolar), 0, RsPago!nDolar)
                        ElseIf RsPago!tMONEDA = "03" Then
                           nEfectivoE2 = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                           nDolar2 = IIf(IsNull(RsPago!nDolar), 0, RsPago!nDolar)
                        ElseIf RsPago!tMONEDA = "04" Then
                           nEfectivoE3 = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                           nDolar3 = IIf(IsNull(RsPago!nDolar), 0, RsPago!nDolar)
                        Else
                           nEfectivoE = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                           nDolar = IIf(IsNull(RsPago!nDolar), 0, RsPago!nDolar)
                        End If
                        
                   Case Is = "02"  'TC
                   'gcaa
                        For i = 1 To 24 'canttarjetas
'                        If i > 7 Then
'                               nTarjeta(8) = nTarjeta(8) + IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
'                               nTarjetaPropina(8) = nTarjetaPropina(8) + IIf(IsNull(RsPago!nPropina), 0, RsPago!nPropina)
'                               txtTarjeta(8).Caption = Format(nTarjeta(8), "###,###,##0.00")
'                               txtPropina(8).Caption = Format(nTarjetaPropina(8), "###,###,##0.00")
'                               nTotalTarjeta = nTotalTarjeta + IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
'                               nTotalTarjetaPropina = nTotalTarjetaPropina + IIf(IsNull(RsPago!nPropina), 0, RsPago!nPropina)
'                               txtTTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
'                               txtTPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
'                                Exit For
'                        Else
                            If tTarjeta(i) = IIf(IsNull(RsPago!tTarjeta), "", RsPago!tTarjeta) Then
                               nTarjeta(i) = nTarjeta(i) + IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                               nTarjetaPropina(i) = nTarjetaPropina(i) + IIf(IsNull(RsPago!nPropina), 0, RsPago!nPropina)
                               txtTarjeta(i).Caption = Format(nTarjeta(i), "###,###,##0.00")
                               txtPropina(i).Caption = Format(nTarjetaPropina(i), "###,###,##0.00")
                               nTotalTarjeta = nTotalTarjeta + IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                               nTotalTarjetaPropina = nTotalTarjetaPropina + IIf(IsNull(RsPago!nPropina), 0, RsPago!nPropina)
                               txtTTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
                               txtTPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
                               Exit For
                            End If
'                         End If
                            
                        Next i
                   
                   Case Is = "03"  'Cheque
                        If RsPago!tMONEDA = "01" Then
                           nChequeN = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        Else
                           nChequeE = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        End If
                                   
                   Case Is = "04"  'Otros Pago
                        For i = 1 To 20
                            If tOtro(i) = IIf(IsNull(RsPago!tOtroTipoPago), "", RsPago!tOtroTipoPago) Then
                               If RsPago!tMONEDA = "01" Then
                                  nOtroN(i) = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                                  nTotalOtroN = nTotalOtroN + nOtroN(i)
                               Else
                                  nOtroE(i) = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                                  nTotalOtroE = nTotalOtroE + nOtroE(i)
                               End If
                               txtOtroN(i).Caption = Format(nOtroN(i), "###,###,##0.00")
                               txtOtroE(i).Caption = Format(nOtroE(i), "###,###,##0.00")
                               txtOtroN(21).Caption = Format(nTotalOtroN, "###,###,##0.00")
                               txtOtroE(21).Caption = Format(nTotalOtroE, "###,###,##0.00")
                               Exit For
                            End If
                        Next i
                        
                   Case Is = "05"  'Puntos
                        If RsPago!tMONEDA = "01" Then
                           nPuntoN = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        Else
                           nPuntoE = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        End If
                   
            End Select
            RsPago.MoveNext
         Loop
         
               
        'JSJR 21/01/2025
        'Isql = "select tTipoPago, tMoneda, isnull(tTarjeta, '') tTarjeta, " & _
           "isnull(tOtroTipoPago, '') tOtroTipoPago, sum(nMonto) as nMonto, " & _
           "sum(nPropina) as nPropina, sum(nDolar) as nDolar " & _
           "from ( " & _
           "    select tTipoPago, tMoneda, isnull(tTarjeta, '') tTarjeta, " & _
           "           isnull(tOtroTipoPago, '') tOtroTipoPago, nMonto, nPropina, nDolar " & _
           "    from dpagodocumento " & _
           "    where tTurno = '" & sTurno & "' " & _
           "    union all " & _
           "    select tTipoPago, tMoneda, tTarjeta, '' as tOtroTipoPago, nMonto, " & _
           "           0 as nPropina, 0 as nDolar " & _
           "    from MINGRESO " & _
           "    where tTurno = '" & sTurno & "' and tEstadoDocumento = '01' and tTipoPago = '02' " & _
           ") as R " & _
           "group by R.tTipoPago, R.tMoneda, isnull(R.tTarjeta, ''), isnull(R.tOtroTipoPago, '')"
           
           
            'JSJR 21/01/2025
            Isql = "select tTipoPago, tMoneda, tTarjeta, '' as tOtroTipoPago, nMonto,0 as nPropina, 0 as nDolar from MINGRESO " & _
                "where tTurno = '" & sTurno & "' and tEstadoDocumento = '01' and tTipoPago = '02'"
         

            Dim RsRecibo As Recordset
            Set RsRecibo = Lib.OpenRecordset(Isql, Cn)
            
            Do While Not RsRecibo.EOF
                Select Case RsRecibo!tTipoPago
                 Case Is = "01"  'Efectivo
                    If RsRecibo!tMONEDA = "01" Then
                        nEfectivoN = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                    ElseIf RsRecibo!tMONEDA = "02" Then
                        nEfectivoE = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                        nDolar = IIf(IsNull(RsRecibo!nDolar), 0, RsRecibo!nDolar)
                    ElseIf RsRecibo!tMONEDA = "03" Then
                        nEfectivoE2 = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                        nDolar2 = IIf(IsNull(RsRecibo!nDolar), 0, RsRecibo!nDolar)
                    ElseIf RsRecibo!tMONEDA = "04" Then
                        nEfectivoE3 = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                        nDolar3 = IIf(IsNull(RsRecibo!nDolar), 0, RsRecibo!nDolar)
                    Else
                        nEfectivoE = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                        nDolar = IIf(IsNull(RsRecibo!nDolar), 0, RsRecibo!nDolar)
                    End If
            
                 Case Is = "02"  'TC
            
                    For i = 1 To 24
                        If tTarjeta(i) = IIf(IsNull(RsRecibo!tTarjeta), "", RsRecibo!tTarjeta) Then
                            nTarjeta(i) = nTarjeta(i) + IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                            nTarjetaPropina(i) = nTarjetaPropina(i) + IIf(IsNull(RsRecibo!nPropina), 0, RsRecibo!nPropina)
                            txtTarjeta(i).Caption = Format(nTarjeta(i), "###,###,##0.00")
                            txtPropina(i).Caption = Format(nTarjetaPropina(i), "###,###,##0.00")
                            nTotalTarjeta = nTotalTarjeta + IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                            nTotalTarjetaPropina = nTotalTarjetaPropina + IIf(IsNull(RsRecibo!nPropina), 0, RsRecibo!nPropina)
                            txtTTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
                            txtTPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
                         Exit For
                        End If
            
            
                    Next i
            
                 Case Is = "03"  'Cheque
                    If RsRecibo!tMONEDA = "01" Then
                        nChequeN = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                    Else
                        nChequeE = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                    End If
            
                 Case Is = "04"  'Otros Pago
                    For i = 1 To 20
                        If tOtro(i) = IIf(IsNull(RsRecibo!tOtroTipoPago), "", RsRecibo!tOtroTipoPago) Then
                            If RsRecibo!tMONEDA = "01" Then
                                nOtroN(i) = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                                nTotalOtroN = nTotalOtroN + nOtroN(i)
                            Else
                                nOtroE(i) = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                                nTotalOtroE = nTotalOtroE + nOtroE(i)
                            End If
                            txtOtroN(i).Caption = Format(nOtroN(i), "###,###,##0.00")
                            txtOtroE(i).Caption = Format(nOtroE(i), "###,###,##0.00")
                            txtOtroN(21).Caption = Format(nTotalOtroN, "###,###,##0.00")
                            txtOtroE(21).Caption = Format(nTotalOtroE, "###,###,##0.00")
                         Exit For
                        End If
                    Next i
            
                 Case Is = "05"  'Puntos
                    If RsRecibo!tMONEDA = "01" Then
                        nPuntoN = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                    Else
                        nPuntoE = IIf(IsNull(RsRecibo!nMonto), 0, RsRecibo!nMonto)
                    End If
            
                End Select
                RsRecibo.MoveNext
            Loop
            
            
         nEfectivoN = nEfectivoN + nIngresoN - nEgresoN - ((nDolar - nEfectivoE) * nCambio) - ((nDolar2 - nEfectivoE2) * nCambio2) - ((nDolar3 - nEfectivoE3) * nCambio3)
         nEfectivoE = nDolar + nIngresoE - nEgresoE
         nEfectivoE2 = nDolar2 - nEgresoE2 '+ nIngresoE - nEgresoE
         nEfectivoE3 = nDolar3 - nEgresoE3
         nTotalTarjeta = nTotalTarjeta + nIngresoTarjeta
      End If
   End If
End Sub


