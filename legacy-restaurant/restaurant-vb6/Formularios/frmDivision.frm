VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDivision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "División de Pedidos"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmDivision.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   Begin VB.Frame fraPedido 
      Caption         =   " Pedido Origen "
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
      Height          =   3195
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   2805
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mesa :"
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
         Index           =   5
         Left            =   465
         TabIndex        =   45
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Orden :"
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
         Index           =   4
         Left            =   450
         TabIndex        =   44
         Top             =   900
         Width           =   435
      End
      Begin VB.Label txtMesa 
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
         Left            =   915
         TabIndex        =   43
         Top             =   1170
         Width           =   1800
      End
      Begin VB.Label txtFecha 
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
         Left            =   915
         TabIndex        =   42
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label txtMozo 
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
         Left            =   915
         TabIndex        =   41
         Top             =   1725
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mozo :"
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
         Index           =   8
         Left            =   450
         TabIndex        =   40
         Top             =   1770
         Width           =   435
      End
      Begin VB.Label txtMotorizado 
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
         Left            =   915
         TabIndex        =   39
         Top             =   2010
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observac. :"
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
         Index           =   6
         Left            =   195
         TabIndex        =   38
         Top             =   2310
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Motorizado :"
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
         Index           =   7
         Left            =   90
         TabIndex        =   37
         Top             =   2055
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
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
         Index           =   1
         Left            =   435
         TabIndex        =   36
         Top             =   315
         Width           =   450
      End
      Begin VB.Label txtObservacion 
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
         Height          =   390
         Left            =   915
         TabIndex        =   35
         Top             =   2325
         Width           =   1800
      End
      Begin VB.Label txtDescuento 
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   915
         TabIndex        =   34
         Top             =   1440
         Width           =   1530
      End
      Begin VB.Label txtOrden 
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
         Left            =   915
         TabIndex        =   33
         Top             =   855
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Index           =   2
         Left            =   135
         TabIndex        =   32
         Top             =   1485
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido :"
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
         Index           =   0
         Left            =   60
         TabIndex        =   31
         Top             =   600
         Width           =   825
      End
      Begin VB.Label txtTipoPedido 
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
         Left            =   915
         TabIndex        =   30
         Top             =   555
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   26
         Left            =   2535
         TabIndex        =   29
         Top             =   1485
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Jala Pedidos (Mesas)"
      Height          =   645
      Index           =   5
      Left            =   9315
      TabIndex        =   23
      Top             =   1657
      Width           =   1260
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
      Height          =   645
      Index           =   0
      Left            =   10620
      Picture         =   "frmDivision.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2415
      Width           =   1260
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
      Height          =   645
      Index           =   1
      Left            =   9315
      Picture         =   "frmDivision.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2415
      Width           =   1260
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cantidad Pedidos"
      Height          =   645
      Index           =   2
      Left            =   9315
      TabIndex        =   20
      Top             =   765
      Width           =   1260
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   735
      Index           =   11
      Left            =   8505
      Picture         =   "frmDivision.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   45
      Width           =   735
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   735
      Index           =   8
      Left            =   8505
      Picture         =   "frmDivision.frx":0F10
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2340
      Width           =   735
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   735
      Index           =   9
      Left            =   8505
      Picture         =   "frmDivision.frx":17DA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1575
      Width           =   735
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   735
      Index           =   10
      Left            =   8505
      Picture         =   "frmDivision.frx":20A4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   810
      Width           =   735
   End
   Begin VB.Frame fraDetalle 
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
      Height          =   5355
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   11850
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   7
         Left            =   6660
         Picture         =   "frmDivision.frx":296E
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   6
         Left            =   7935
         Picture         =   "frmDivision.frx":3238
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   5
         Left            =   9210
         Picture         =   "frmDivision.frx":3B02
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   4
         Left            =   10485
         Picture         =   "frmDivision.frx":43CC
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Compartir"
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
         Index           =   1
         Left            =   5265
         Picture         =   "frmDivision.frx":4C96
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4005
         Width           =   1260
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Disgregar"
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
         Index           =   0
         Left            =   5265
         TabIndex        =   18
         Top             =   3285
         Width           =   1260
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   3915
         Picture         =   "frmDivision.frx":4FA0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   2640
         Picture         =   "frmDivision.frx":586A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   1365
         Picture         =   "frmDivision.frx":6134
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   90
         Picture         =   "frmDivision.frx":69FE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4695
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Todos"
         Height          =   645
         Index           =   0
         Left            =   5265
         Picture         =   "frmDivision.frx":72C8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2340
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Quitar Uno"
         Height          =   645
         Index           =   2
         Left            =   5265
         Picture         =   "frmDivision.frx":7852
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1620
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Todos"
         Height          =   645
         Index           =   1
         Left            =   5265
         Picture         =   "frmDivision.frx":7DDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   1260
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Pasar Uno"
         Height          =   645
         Index           =   3
         Left            =   5265
         Picture         =   "frmDivision.frx":8366
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   900
         Width           =   1260
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle2 
         Height          =   4080
         Left            =   6615
         TabIndex        =   15
         Top             =   180
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   7197
         _LayoutType     =   4
         _RowHeight      =   25
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
         Splits(0).ScrollBars=   2
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
         Caption         =   "Productos a Facturar"
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
      Begin TrueOleDBGrid80.TDBGrid grdDetalle1 
         Height          =   4080
         Left            =   90
         TabIndex        =   16
         Top             =   180
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   7197
         _LayoutType     =   4
         _RowHeight      =   25
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
         Splits(0).ScrollBars=   2
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
         Caption         =   "Productos por Facturar"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Items :"
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
         Left            =   7200
         TabIndex        =   53
         Top             =   4395
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Items :"
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
         Left            =   540
         TabIndex        =   52
         Top             =   4395
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
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
         Left            =   9810
         TabIndex        =   51
         Top             =   4388
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
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
         Index           =   3
         Left            =   3240
         TabIndex        =   50
         Top             =   4395
         Width           =   495
      End
      Begin VB.Label txtItem2 
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
         Left            =   8235
         TabIndex        =   49
         Top             =   4350
         Width           =   765
      End
      Begin VB.Label txtMonto2 
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
         Left            =   10395
         TabIndex        =   48
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Label txtItem 
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
         Left            =   1575
         TabIndex        =   47
         Top             =   4350
         Width           =   765
      End
      Begin VB.Label txtMonto 
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
         Left            =   3870
         TabIndex        =   46
         Top             =   4350
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Genera Ped"
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
      Index           =   3
      Left            =   9315
      Picture         =   "frmDivision.frx":88F0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   1260
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Elimina Ped"
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
      Index           =   4
      Left            =   10620
      Picture         =   "frmDivision.frx":89F2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   1260
   End
   Begin TrueOleDBGrid80.TDBGrid grdPedido 
      Height          =   3195
      Left            =   2835
      TabIndex        =   17
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   5636
      _LayoutType     =   4
      _RowHeight      =   25
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
      Splits(0).ScrollBars=   2
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
      Caption         =   "Cabecera de Pedidos"
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
End
Attribute VB_Name = "frmDivision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDetalle1 As Recordset
Dim RsDetalle2 As Recordset
Dim RsPedido As Recordset
Dim sDetalle As String
Dim sCabecera As String
Dim sObservacion As String
Dim nContador As Integer
Dim sCorrelativo As String
Dim xPedido As String
Dim nPos As Variant
Dim lCompartir As Boolean

'disgregar
Dim sProductoPropiedad As String
Dim xDisgregar As Boolean
Dim itemDisgregado As String

'25/06
Dim sUsuarioAutoriza As String

Dim nPedidosC As Integer

Dim sDetalleCombo As String
Dim xp As Integer

Private Sub cmdDetalle_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    Select Case Index
    
           Case Is = 0  ' Disgregar
                If RsDetalle1.RecordCount = 0 Then
                   Exit Sub
                End If
            
                    If RsDetalle1!nCantidad > 1 Then
                               Dim xCantidad As Double
                               Dim xItem As String
                               Dim yItem As String
                               
                               Dim lPropiedad As Boolean
                               lPropiedad = False
                               xDisgregar = True
                               xCantidad = RsDetalle1!nCantidad - 1
                               
                               xItem = RsDetalle1!tItem
                               
                               If Calcular("select count(*) as codigo from tproductopropiedad where tcodigopedido='" & xPedido & "' and titem='" & xItem & "'", Cn) > 0 Then
                                    lPropiedad = True
                               End If
                               
                               If lDisgrega Then
                                    sTipo = ""
                                    frmNumPad.Show vbModal
                                    If wEnter And val(sDescrip) > 0 And val(sDescrip) < RsDetalle1!nCantidad Then
                                       'Primera Division
                                       xCantidad = val(sDescrip)
                                       Cn.Execute "update " & sDetalle & " set nCantidad= " & xCantidad & ", nImpuesto1 = nImpuesto1/nCantidad*" & xCantidad & ", nImpuesto2 = nImpuesto2/nCantidad*" & xCantidad & ", nImpuesto3 = nImpuesto3/nCantidad*" & xCantidad & ", nVenta = nVenta/nCantidad*" & xCantidad & " where tCodigoPedido='" & xPedido & "' and tItem='" & RsDetalle1!tItem & "'"
                                                                                                                                  
                                       'Segunda Division
                                       yItem = Calcular("select max(titem) as codigo from " & sDetalle, Cn)
                                       yItem = Mid("000", 1, 3 - Len(LTrim(str(val(yItem) + 1)))) & LTrim(str(val(yItem) + 1))
                                       xCantidad = RsDetalle1!nCantidad - xCantidad
                                       
                                       Cn.Execute "Insert into " & sDetalle & _
                                                  " select tCodigoPedido, '" & yItem & "' as tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                                                  "tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, " & _
                                                  "nPrecioOficial, " & xCantidad & ", nPrecioImpuesto1 * " & xCantidad & ",  nPrecioImpuesto2 * " & xCantidad & ", nPrecioImpuesto3 * " & xCantidad & ", nPrecioVenta * " & xCantidad & ", tObservacion, tCortesia, lImprime, tEstadoItem, tArea, " & _
                                                  "lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD, tUsuarioD, " & _
                                                  "nInsumo , nGasto, nManoObra, nOrden, lCorte, tPosicion, fEnvio, nEnvio, tUnidadNegocio, tOferta, tAutorizaOferta,tsubalmacen,fdiacontable, lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio, Estado,tCajaD,tUsuarioAutorizaEnvio,tPerfilUsuarioAutoriza " & _
                                                  "from " & sDetalle & " where tCodigoPedido='" & xPedido & "' and tItem='" & xItem & "'"
                                                  
                                        'CESAR COMBO
                                        Cn.Execute "Insert Into " & sDetalleCombo & _
                                                      " select tCodigoPedido, '" & yItem & "' as tItem, tItemCombo, tProducto, tProductoCombo, " & xCantidad & ", tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, " & _
                                                      "tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio, Estado From " & sDetalleCombo & " where tCodigoPedido='" & xPedido & "' and tItem='" & xItem & "'"
                                        '---------------------------
                                                  
                                       If lPropiedad = True Then
                                          Cn.Execute "insert into " & sProductoPropiedad & " select tcodigopedido , '" & yItem & "', tcodigopropiedad,tproducto,tenlace, ninsumo,ngasto,nmanoobra,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario " & _
                                                     "From tproductopropiedad where tcodigopedido ='" & xPedido & "' and titem='" & xItem & "'"
                                       End If
                                                  
                                    Else
                                       MsgBox "Cantidad no aceptada", vbCritical
                                    End If
                               
                               Else
                                    
                                  Dim xcantidadA As Double
                                  xcantidadA = Calcular("select nCantidad as codigo from Dpedido where tCodigoPedido ='" & xPedido & "' and tItem ='" & RsDetalle1!tItem & "'", Cn)
                                     
                                  Cn.Execute "update " & sDetalle & " set nCantidad=1, nImpuesto1=nImpuesto1/nCantidad, nImpuesto2=nImpuesto2/nCantidad, nImpuesto3=nImpuesto3/nCantidad, nVenta=nVenta/nCantidad where tCodigoPedido='" & xPedido & "' and tItem='" & RsDetalle1!tItem & "'"
                                  
                                  'Cn.Execute "UPDATE " & sDetalleCombo & " set nCantidad = 1 Where tCodigoPedido='" & xPedido & "' and tItem='" & RsDetalle1!tItem & "'"
                                  '------------------
                                  
                                  Do While True
                                     yItem = Calcular("select max(titem) as codigo from " & sDetalle, Cn)
                                     yItem = Mid("000", 1, 3 - Len(LTrim(str(val(yItem) + 1)))) & LTrim(str(val(yItem) + 1))
                                     If xCantidad > 0 Then
                                        
                                        If xCantidad > 1 Then
                                        
                                           Cn.Execute "Insert into " & sDetalle & _
                                                      " select tCodigoPedido, '" & yItem & "' as tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                                                      "tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, " & _
                                                      "nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, " & _
                                                      "lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD, tUsuarioD, " & _
                                                      "nInsumo , nGasto, nManoObra, nOrden, lCorte, tPosicion, fEnvio, nEnvio, tUnidadNegocio, tOferta, tAutorizaOferta,tsubalmacen,fDiacontable,  lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio, Estado,tcajaD,tUsuarioAutorizaEnvio,tPerfilUsuarioAutoriza " & _
                                                      "from " & sDetalle & " where tCodigoPedido='" & xPedido & "' and tItem='" & xItem & "'"
                                                      
                                           Cn.Execute "Insert Into " & sDetalleCombo & _
                                                      " select tCodigoPedido, '" & yItem & "' as tItem, tItemCombo, tProducto, tProductoCombo, nCantidad/" & xcantidadA & ", tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto/" & xcantidadA & ",nImpuesto1/" & xcantidadA & ", nImpuesto2/" & xcantidadA & ",  nImpuesto3/" & xcantidadA & ",   nVenta/" & xcantidadA & ", nInsumo/" & xcantidadA & ", nGasto/" & xcantidadA & ", nManoObra/" & xcantidadA & ", lImprimeArea, lImprime, nOrden, " & _
                                                      "tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio, Estado From " & sDetalleCombo & " where tCodigoPedido='" & xPedido & "' and tItem='" & xItem & "'"
                                                      
                                        Else
                                        
                                           Cn.Execute "Insert into " & sDetalle & _
                                                      " select tCodigoPedido, '" & yItem & "' as tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                                                      "tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, " & _
                                                      "nPrecioOficial, " & xCantidad & ", nImpuesto1 * " & xCantidad & ",  nImpuesto2 * " & xCantidad & ", nImpuesto3 * " & xCantidad & ", nVenta * " & xCantidad & ", tObservacion, tCortesia, lImprime, tEstadoItem, tArea, " & _
                                                      "lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD, tUsuarioD, " & _
                                                      "nInsumo , nGasto, nManoObra, nOrden, lCorte, tPosicion, fEnvio, nEnvio, tUnidadNegocio, tOferta, tAutorizaOferta,tsubalmacen, fdiacontable,  lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio, Estado,tCajaD,tUsuarioAutorizaEnvio,tPerfilUsuarioAutoriza " & _
                                                      "from " & sDetalle & " where tCodigoPedido='" & xPedido & "' and tItem='" & xItem & "'"
                                                      
                                            Cn.Execute "Insert Into " & sDetalleCombo & _
                                                       " select tCodigoPedido, '" & yItem & "' as tItem, tItemCombo, tProducto, tProductoCombo, nCantidad/" & xcantidadA & ", tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto/" & xcantidadA & ",nImpuesto1/" & xcantidadA & ", nImpuesto2/" & xcantidadA & ",  nImpuesto3/" & xcantidadA & ",   nVenta/" & xcantidadA & ", nInsumo/" & xcantidadA & ", nGasto/" & xcantidadA & ", nManoObra/" & xcantidadA & ", lImprimeArea, lImprime, nOrden, " & _
                                                       "tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio,Estado From " & sDetalleCombo & " where tCodigoPedido='" & xPedido & "' and tItem='" & xItem & "'"
                                                      
                                        End If
                                        
                                        If lPropiedad = True Then
                                             Cn.Execute "insert into " & sProductoPropiedad & " select tcodigopedido , '" & yItem & "', tcodigopropiedad,tproducto,tenlace, ninsumo,ngasto,nmanoobra, ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario " & _
                                                        "From tproductopropiedad where tcodigopedido ='" & xPedido & "' and titem='" & xItem & "'"
                                        End If
                                        xCantidad = xCantidad - 1
                                      Else
                                        Exit Do
                                      End If
                                  Loop
                                  
                                  Cn.Execute "UPDATE " & sDetalleCombo & " set nCantidad = nCantidad/ " & xcantidadA & ", nPrecioNeto = nPrecioNeto /" & xcantidadA & ", nImpuesto1 = nImpuesto1/" & xcantidadA & ", nImpuesto2 = nImpuesto2/" & xcantidadA & ", nImpuesto3 = nImpuesto3/" & xcantidadA & ",nVenta=nVenta/" & xcantidadA & ", nInsumo = nInsumo/" & xcantidadA & ", nGasto = nGasto/" & xcantidadA & ", nManoObra = nManoObra/" & xcantidadA & " Where tCodigoPedido='" & xPedido & "' and tItem='" & RsDetalle1!tItem & "'"
                                 
                              End If
                                
                              RsDetalle1.Requery
                    End If
                    
           Case Is = 1  ' Compartir
                 If lCompartir = False Then

                        nPedidosC = RsPedido.RecordCount + 1
                        If RsPedido.RecordCount = 0 Then
                           Exit Sub
                        End If
                        If MsgBox("Deseas compartir los " & RsDetalle1.RecordCount & " item(s), en los " & RsPedido.RecordCount + 1 & " pedidos ?", vbInformation + vbYesNo, sMensaje) = vbYes Then
                           Isql = "insert into " & sDetalle & " " & _
                                  "select " & sCabecera & ".tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, " & _
                                  "nPrecioOficial, nCantidad/" & RsPedido.RecordCount + 1 & ", nImpuesto1/" & RsPedido.RecordCount + 1 & ", nImpuesto2/" & RsPedido.RecordCount + 1 & ", nImpuesto3/" & RsPedido.RecordCount + 1 & ", nVenta/" & RsPedido.RecordCount + 1 & ", " & sDetalle & ".tObservacion, tCortesia, lImprime, tEstadoItem, tArea, " & _
                                  "lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD, tUsuarioD, nInsumo , nGasto, nManoObra, nOrden, lCorte, tPosicion, fEnvio, nEnvio, tUnidadNegocio, tOferta, tAutorizaOferta,tsubalmacen,fdiacontable,  lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio, 'X' as Estado,tCajaD,tUsuarioAutorizaEnvio,tPerfilUsuarioAutoriza " & _
                                  "From " & sCabecera & ", " & sDetalle & " where Estado=' '"
                                  
                           Cn.Execute Isql
                           Isql = "update " & sDetalle & " " & _
                                   "set ncantidad= ncantidad / " & RsPedido.RecordCount + 1 & ",nimpuesto1= nimpuesto1 / " & RsPedido.RecordCount + 1 & ",nimpuesto2= nimpuesto2 / " & RsPedido.RecordCount + 1 & " , nimpuesto3 = nimpuesto3/" & RsPedido.RecordCount + 1 & " , nventa=nventa/" & RsPedido.RecordCount + 1 & "     " & _
                                   "where tcodigopedido='" & xPedido & "'"
                           Cn.Execute Isql
                           
                           RsPedido.Requery
                           RsDetalle2.Requery
                           RsDetalle1.Requery
                           lCompartir = True

                        End If
                    Else
                            MsgBox "Ya compartió el pedido original"
                    End If

    End Select
End Sub

Private Sub cmdMovimiento_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    If RsPedido.RecordCount = 0 Then
       Exit Sub
    End If
    Dim xItem As String
    
    Screen.MousePointer = vbHourglass
    Select Case Index
           Case Is = 1  ' Pasar Todas
                 Cn.Execute "update " & sDetalle & " set Estado ='X', tCodigoPedido='" & RsPedido!tCodigoPedido & "' where tCodigoPedido='" & xPedido & "'"
           
           Case Is = 0  ' Quitar Todas
                If RsDetalle2.RecordCount <> 0 Then
                    Cn.Execute "update " & sDetalle & " set Estado =' ', tCodigoPedido='" & xPedido & "' where tCodigoPedido='" & RsDetalle2!tCodigoPedido & "'"
                End If
           Case Is = 3  ' Pasar una
                If RsDetalle1.RecordCount <> 0 Then
                   Cn.Execute "update " & sDetalle & " set Estado ='X', tCodigoPedido='" & RsPedido!tCodigoPedido & "' where tItem = '" & RsDetalle1!tItem & "' and tCodigoPedido = '" & RsDetalle1!tCodigoPedido & "'"
                End If
           
           Case Is = 2  ' Quitar una
                If RsDetalle2.RecordCount <> 0 Then
                   Cn.Execute "update " & sDetalle & " set Estado =' ', tCodigoPedido='" & xPedido & "' where tItem = '" & RsDetalle2!tItem & "' and tCodigoPedido = '" & RsDetalle2!tCodigoPedido & "'"
                End If
    End Select
    nPos = RsPedido.Bookmark
    RsPedido.Requery
    RsPedido.Bookmark = nPos
    RsDetalle1.Requery
    RsDetalle2.Requery
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    Select Case Index
           Case Is = 0  ' Ultimo Detalle1
                MoverPuntero Ultimo, grdDetalle1
           Case Is = 1  ' Siguiente Detalle1
                MoverPuntero siguiente, grdDetalle1
           Case Is = 2  ' Previo Detalle1
                MoverPuntero previo, grdDetalle1
           Case Is = 3  ' Primero Detalle1
                MoverPuntero Primero, grdDetalle1
           Case Is = 4  ' Ultimo Detalle2
                MoverPuntero Ultimo, grdDetalle2
           Case Is = 5  ' Siguiente Detalle2
                MoverPuntero siguiente, grdDetalle2
           Case Is = 6  ' Previo Detalle2
                MoverPuntero previo, grdDetalle2
           Case Is = 7  ' Primero Detalle2
                MoverPuntero Primero, grdDetalle2
           Case Is = 8 ' Ultimo Documento
                MoverPuntero Ultimo, grdPedido
           Case Is = 9 ' Siguiente Documento
                MoverPuntero siguiente, grdPedido
           Case Is = 10  ' Previo Documento
                MoverPuntero previo, grdPedido
           Case Is = 11  ' Primero Documento
                MoverPuntero Primero, grdPedido
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case Index
          Case Is = 0 ' Aceptar
    On Error GoTo ERRORDIVISION

               Dim xmaxped As Double
                '// validacion 05/05/2018 ELDC
               If Calcular("select testadopedido as Codigo from mpedido where tcodigopedido='" & xPedido & "'", Cn) <> "01" Then
                    MsgBox "El pedido " & xPedido & " No puede ser Dividido, Estado del pedido Diferente de Emitido, Refrescar los pedidos. ", vbInformation, sMensaje
                    Exit Sub
               End If
                '// Fin de Validacion
               If RsPedido.RecordCount > 0 Then
                     RsPedido.MoveFirst
                    Do While Not RsPedido.EOF
                        xmaxped = (Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & LTrim(RsPedido!tCodigoPedido) & "'", Cn))
                        If xmaxped > 0 Then
                            Dim nVentaTemp As Double
                            nVentaTemp = Calcular("select sum(nVenta) as codigo from " & sDetalle & " where tcodigopedido='" & LTrim(RsPedido!tCodigoPedido) & "'", Cn)
                            If CDbl(Calcular("select sum(nventa) as codigo from dpedido where tcodigopedido='" & LTrim(RsPedido!tCodigoPedido) & "'", Cn)) + nVentaTemp > xmaxped Then
                                MsgBox ("El pedido: " & LTrim(RsPedido!tCodigoPedido) & " tiene Como Monto Maximo pedido: " & xmaxped & " ,Verificar!!!"), vbInformation
                                Exit Sub
                            End If
                        End If
                        RsPedido.MoveNext
                    Loop
                    
                  RsPedido.MoveFirst
               End If
               
               Cn.BeginTrans
               Dim sCorrela As String
               Dim nItem As Integer
               Dim xItem As String
          
               
               
               Dim iTipoPedido As String
               iTipoPedido = Calcular("Select tTipoPedido As Codigo from MPEDIDO where tCodigoPedido='" & xPedido & "'", Cn)
               
               Cn.Execute "delete DPEDIDO where tCodigoPedido='" & xPedido & "'"
               
               Cn.Execute "insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                          "tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, " & _
                          "tEstadoItem, tArea, lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD , tUsuarioD, nInsumo, nGasto, nManoObra, nOrden,tsubalmacen,fdiacontable, fEnvio, nEnvio, lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio,tcajaD, tUsuarioAutorizaEnvio, tPerfilUsuarioAutoriza) " & _
                          "select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                          "tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, " & _
                          "tEstadoItem, tArea, lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD , tUsuarioD, nInsumo, nGasto, nManoObra, nOrden,tsubalmacen,fdiacontable, fEnvio, nEnvio, lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio,tcajaD, tUsuarioAutorizaEnvio, tPerfilUsuarioAutoriza " & _
                          "From " & sDetalle & " " & _
                          "where Estado=' '"

                    If xDisgregar = True Then
                          Cn.Execute "insert into tproductopropiedad " & _
                                      " select tcodigopedido , titem, tcodigopropiedad,tproducto,tenlace, ninsumo,ngasto,1,nmanoobra,ncantidad, ninsumounitario,ngastounitario,nmanoobraunitario " & _
                                      " From " & sProductoPropiedad & " where tcodigopedido ='" & xPedido & "'"
                    End If
                    
                ' aqui GCAA 12102020
                    
                    'ACTUALIZA  tAplicaExoneracion, tAplicaInafectacion
         
                    Dim oComando  As clsComando
                    Set oComando = New clsComando
                    If Not oComando.CreateCmdSp("spUpdate_DPEDIDO_Ina", Cn) Then
                       Set oComando = Nothing
                       Exit Sub
                    End If
                    
                    oComando.CreateParameter "@pedido", adVarChar, adParamInput, 10, xPedido
                    
                    If Not oComando.GetParamOK Then
                       Set oComando = Nothing
                       Exit Sub
                    End If
                
                    If Not oComando.ExecSP Then
                       Set oComando = Nothing
                       Exit Sub
                     End If

                    
                    
                    
                'COMBO CESAR
                Cn.Execute "delete CPEDIDO where tCodigoPedido='" & xPedido & "'"
                
                Isql = "insert into CPEDIDO (tCodigoPedido, tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, " & _
                           "tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio) " & _
                           "select tCodigoPedido, tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, " & _
                           "tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio " & _
                           "From " & sDetalleCombo & " " & _
                           "where Estado=' '"
                Cn.Execute Isql
                '-------------------------------------------------------

               If Not RsPedido.EOF Then
                        If lCompartir = True Then
                            Isql = "update cpedido set ncantidad=ncantidad/" & nPedidosC & ",nprecioneto=nprecioneto/" & nPedidosC & ",nimpuesto1=nimpuesto1/" & nPedidosC & ", nimpuesto2=nimpuesto2/" & nPedidosC & ",nimpuesto3=nimpuesto3/" & nPedidosC & ",nventa=nventa/" & nPedidosC & ", ninsumo=ninsumo/" & nPedidosC & ",ngasto=ngasto/" & nPedidosC & ",nmanoobra=nmanoobra/" & nPedidosC & " Where tCodigoPedido = '" & xPedido & "'"
                            Cn.Execute Isql
        
                        End If
                
                        Dim sAnio As String
                        sAnio = Calcular("SELECT YEAR(GETDATE()) AS CODIGO", Cn)
                
                        RsPedido.MoveFirst
                  
                        Do While Not RsPedido.EOF
                        
                            sCorrelativo = LTrim(RsPedido!tCodigoPedido)
                            If Len(sCorrelativo) = 3 And RsPedido!nItem > 0 Then
                            sCorrela = Calcular("select max(tCodigoPedido) as Codigo from mpedido where substring(tCodigoPedido,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                        
                                        If IsNull(sCorrela) Or Mid(sCorrela, 1, 2) <> Mid(Trim(str(sAnio)), 3, 2) Then
                                           sCorrela = Mid(Trim(str(sAnio)), 3, 2) & "00000001"
                                        Else
                                           sCorrela = Mid(Trim(str(sAnio)), 3, 2) & Lib.Correlativo(Mid(sCorrela, 3, 8), 8)
                                        End If
                                                                                      
                                        'Insertar Cabecera
                                        'pariente2013
                                        Cn.Execute "update " & sCabecera & " set tCodigoPedido='" & sCorrela & "' where tcodigoPedido='" & sCorrelativo & "'"
                                        Cn.Execute "Insert into MPEDIDO (tCodigoPedido, nCorrelativo, tClienteDelivery, tClienteCtaCte, fFecha, tMoneda, nMonto, tEstadoPedido, tTipoAtencion, tTipoPedido, lPrioridad, tAnulacionPedido, tMesa, nMesa, tMozo, tMotorizado, tCaja, tSalon, tTurno, fProgramacion, " & _
                                                                        "nTiempo, tObservacion, tUsuario, fRegistro, nAdulto, nNino, tMotivoAnulacion, tusuarioAnulado, fRegAnulado, tObservacionAnulado, tTurnoAnulado, tClienteCorp, tTienda, fRegCuenta, nPrecuenta, tCajaAnterior, tTurnoAnterior, " & _
                                                                        "tComanda, tPuntoVenta, tHabitacion, tReserva, tPasajero, tCompania, tContacto, tFichaPasajero, tTipoComanda, nDescuento, tDescuento, tObservacionDescuento, tUsuarioDescuento, tEmpacador, fEmpacador, fAsignacion, fSalida, fLlegada,fdiacontable,tcodigoinvitado,tcodigopariente, lAtendidoC, fAtendidoC, tUsuarioAtendio,tentregara, ttiporecepcion, nMinutosAntesEnvio)  " & _
                                                                        "select '" & sCorrela & "', nCorrelativo, tClienteDelivery, tClienteCtaCte, getdate(), tMoneda, nMonto, tEstadoPedido, tTipoAtencion, " & _
                                                                        "tTipoPedido, lPrioridad, tAnulacionPedido, '' as tMesa, 0, tMozo, tMotorizado, tCaja, tSalon, tTurno, fProgramacion, nTiempo, " & _
                                                                        "'" & RsPedido!tObservacion & "', tUsuario, getdate(), 0, 0, tMotivoAnulacion, tUsuarioAnulado, fRegAnulado, tObservacionAnulado, tTurnoAnulado, " & _
                                                                        "tClienteCorp, tTienda, fRegCuenta, 0, tCajaAnterior, tTurnoAnterior, '', tPuntoVenta, tHabitacion, tReserva, tPasajero, tCompania, tContacto, tFichaPasajero, tTipoComanda, " & _
                                                                        "nDescuento , tDescuento, tObservacionDescuento, tUsuarioDescuento, " & _
                                                                        "tEmpacador, fEmpacador, fAsignacion, fSalida, fLlegada,'" & Format(obtieneDiaContable, "yyyyMMdd") & "',isnull(tcodigoinvitado,'') tcodigoinvitado,isnull(tcodigopariente,'') tcodigopariente, lAtendidoC, fAtendidoC, tUsuarioAtendio,tentregara,0, nMinutosAntesEnvio " & _
                                                                        "from MPEDIDO where tCodigoPedido='" & xPedido & "'"
                                                                        
                                        If lcover Then
                                          Cn.Execute "update Mpedido set nAdulto=1 where tcodigoPedido='" & sCorrela & "'"
                                        End If
                                                                        
                                        'pariente2013
                                        Cn.Execute "update " & sDetalle & " set tCodigoPedido='" & sCorrela & "' where tcodigoPedido='" & sCorrelativo & "'"
                                        sCorrelativo = sCorrela
                            End If
                            
                            nItem = val(Calcular("select max(tItem) as Codigo from DPEDIDO where tCodigoPedido='" & sCorrelativo & "'", Cn)) + 1
                                                             
                            Cn.Execute "Insert into DPEDIDO (tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                                            "tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, " & _
                                            "tEstadoItem, tArea, lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD , tUsuarioD, nInsumo, nGasto, nManoObra, nOrden,lCorte, tPosicion, fEnvio, nEnvio, tUnidadNegocio, tOferta, tAutorizaOferta,tsubalmacen,fdiacontable, lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio,tcajaD, tUsuarioAutorizaEnvio, tPerfilUsuarioAutoriza) " & _
                                       "select tCodigoPedido, substring('000',1,3 - len(ltrim(cast(cast(tItem as Integer) + " & nItem & " as char) ))) + cast(cast(tItem as Integer) + " & nItem & " as char), " & _
                                            "tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                                            "tMoneda, nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, nDescuento, " & _
                                            "nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, " & _
                                            "lCombinacion, nCombinacion, lImprimeArea, tFacturado, tDocumento, lTransferido, tComanda, fRegistro, tMozoD, tUsuarioD, " & _
                                            "nInsumo , nGasto, nManoObra, nOrden, lCorte, tPosicion, fEnvio, nEnvio, tUnidadNegocio, tOferta, tAutorizaOferta,tsubalmacen,'" & Format(obtieneDiaContable, "yyyyMMdd") & "', lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio,tcajaD, tUsuarioAutorizaEnvio, tPerfilUsuarioAutoriza " & _
                                            "from " & sDetalle & " where tCodigoPedido='" & sCorrelativo & "'"
                                            
                             ' aqui GCAA 12102020
                             
                              'ACTUALIZA  tAplicaExoneracion, tAplicaInafectacion
         
                               ' Dim oComando  As clsComando
                                Set oComando = New clsComando
                                If Not oComando.CreateCmdSp("spUpdate_DPEDIDO_Ina", Cn) Then
                                   Set oComando = Nothing
                                   Exit Sub
                                End If
                                
                                oComando.CreateParameter "@pedido", adVarChar, adParamInput, 10, sCorrelativo
                                
                                If Not oComando.GetParamOK Then
                                   Set oComando = Nothing
                                   Exit Sub
                                End If
                            
                                If Not oComando.ExecSP Then
                                   Set oComando = Nothing
                                   Exit Sub
                                 End If

                            'MESA247
                            If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sCorrelativo & "'", Cn) <> "" Then
                                Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sCorrelativo & "'"
                            End If
                                            
                            'CESAR COMBO
                            Isql = "Insert into CPEDIDO (tCodigoPedido,tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, " & _
                                       "tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio) " & _
                                       "select tCodigoPedido, substring('000',1,3 - len(ltrim(cast(cast(tItem as Integer) + " & nItem & " as char) ))) + cast(cast(tItem as Integer) + " & nItem & " as char), " & _
                                       "tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, " & _
                                       "tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio " & _
                                        "from " & sDetalleCombo & " where tCodigoPedido='" & sCorrelativo & "'"
                            Cn.Execute Isql
                            '---------------------------------------------------------------------------------
                     
                                            
                            'Transferencia
                            '25/06
                            If sUsuarioAutoriza = "" Then
                               sUsuarioAutoriza = sUsuario
                            End If
                     
'                            Dim MaxIdLog As Integer
'                            MaxIdLog = Calcular("select isnull( MAX(ID),0) + 1 as codigo FROM LOG_INFOREST", Cn)
                         
                            '25/06
                            Cn.Execute "Insert into TPEDIDO (tPedidoIni, tItemIni, tPedidoFin, tItemFin, tProducto, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, fRegistro, tUsuario,fDiaContable) " & _
                             "select '" & xPedido & "', tItem, '" & sCorrelativo & "', substring('000',1,3 - len(ltrim(cast(cast(tItem as Integer) + " & nItem & " as char) ))) + cast(cast(tItem as Integer) + " & nItem & " as char), tCodigoProducto, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, getdate(), '" & sUsuarioAutoriza & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' from " & sDetalle & " where tCodigoPedido='" & sCorrelativo & "'"
                             
'                            Cn.Execute "insert into LOG_INFOREST " & _
'                            " select 'PUNTO VENTA' ,  'TRANSFERENCIA' , '" & xPedido & "', '', '','Se ha realizado la transferencia desde el Pedido '+'" & xPedido & "'+' (Prod.'+tCodigoProducto+' - Cant. '+ cast(nCantidad as nvarchar)+' [ Imp1:'+ cast(nImpuesto1 as nvarchar) +', Imp2:'+ cast(nImpuesto2 as nvarchar)+', Imp3:'+ cast(nImpuesto3 as nvarchar)+', nVenta:'+ cast(nVenta as nvarchar)+' ]. Hacia el pedido '+'" & sCorrelativo & "', 'Transferencia del pedido '+'" & xPedido & "'+' hacia el pedido '+'" & sCorrelativo & "' , 'Transferencia del pedido', '" & sUsuario & "', cast(getdate() as datetime), 1, " & MaxIdLog & " from " & sDetalle & " where tCodigoPedido='" & sCorrelativo & "'"
                                'LOG INFOREST
                                
'                                Dim Detalletransferencia As Long
'                                Isql = "Select tCodigoProducto from " & sDetalle & " where tCodigoPedido='" & sCorrelativo & "'"
'                                Detalletransferencia = Cn.Execute Isql
                                'Call Log_Inforest("PUNTO VENTA", "TRANSFERENCIA", sPedido, Error, "", "", "Transferencia del producto ", "Modificación del tipo pedido (" & sPedido & ") ", sUsuario)
                   
                              'PARTICION DE PROPIEDADES
                               '
                               If lCompartir = False Then
                                     Cn.Execute "update TPRODUCTOPROPIEDAD set tcodigopedido = '" & sCorrelativo & "' , titem=substring('000',1,3 - len(ltrim(cast(cast(tItem as Integer) + " & nItem & " as char) ))) + cast(cast(tItem as Integer) + " & nItem & " as char) where tcodigopedido='" & xPedido & "' and tItem in(select tItem from " & sDetalle & " where tcodigopedido='" & sCorrelativo & "')"
                               Else
                                       Dim cPropiedades As Integer
                                       Dim rsPedidosItem As New ADODB.Recordset
                                       Dim rsPI As New ADODB.Recordset
                                       Dim P As Long
                                       
                                       'verifica si hay propiedades en el pedido original
                                       cPropiedades = Calcular("select count(*) as codigo from tproductopropiedad where tcodigopedido='" & xPedido & "' ", Cn)
                                       If (cPropiedades > 0) Then
            
                                          ' Set rsPedidosItem = Lib.OpenRecordset("select tpedidoini,titemini,tpedidofin,max(titemfin) titemfin,tproducto ,max(ncorrelativo) from tpedido where tpedidoini='" & xPedido & "' and tpedidofin='" & sCorrelativo & "' group by tpedidoini,titemini,tpedidofin,tproducto", Cn)
                                          Set rsPedidosItem = Lib.OpenRecordset("SELECT  TP.tPedidoIni, TP.tItemIni,TP.tPedidoFin,MAX(TP.tItemFin) TITEMFIN, dbo.DPEDIDO.tCodigoProducto  , Max (TP.nCorrelativo) FROM    dbo.DPEDIDO INNER JOIN ( select tpedidoini,titemini,tpedidofin,max(titemfin) AS titemfin ,max(ncorrelativo)  AS NCORRELATIVO from tpedido where tpedidoini='" & xPedido & "' and tpedidofin='" & sCorrelativo & "'   group by tpedidoini,titemini,tpedidofin ) TP ON dbo.DPEDIDO.tCodigoPedido = TP.TPedidoFin AND dbo.DPEDIDO.tItem = TP.tItemFin WHERE   (TP.tPedidoFin = '" & sCorrelativo & "' ) AND (TP.tPedidoIni = '" & xPedido & "') GROUP BY   TP.tPedidoIni,            TP.tItemIni,            TP.tPedidoFin,dbo.DPEDIDO.tCodigoProducto", Cn)
                                         
                                           If (rsPedidosItem.RecordCount > 0) Then
                                               rsPedidosItem.MoveFirst
                                                   For P = 0 To rsPedidosItem.RecordCount - 1
                                                                   Isql = " insert into tproductopropiedad (tcodigopedido, titem,tcodigopropiedad,tproducto,tenlace,ninsumo,ngasto,nmanoobra,lreplica,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario) " & _
                                                                            " select '" & sCorrelativo & "','" & rsPedidosItem.Fields(3) & "',tcodigopropiedad,'" & rsPedidosItem.Fields(4) & "',tenlace,ninsumo,ngasto,nmanoobra,lreplica,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from tproductopropiedad " & _
                                                                            " where tcodigopedido='" & xPedido & "' and tproducto='" & rsPedidosItem.Fields(4) & "' and titem='" & rsPedidosItem.Fields(1) & "'"
                                                                ' inserta en tproductopropiedad los item del pedido original que tienen propiedades
                                                                 Cn.Execute Isql
                                                      rsPedidosItem.MoveNext
                                                   Next P
                                           End If
                                       End If
                                End If
                    
                                If lCompartir = False Then
                                   Isql = "update CPEDIDO set tcodigopedido = '" & sCorrelativo & "', tItem=substring('000',1,3 - len(ltrim(cast(cast(tItem as Integer) + " & nItem & " as char) ))) + cast(cast(tItem as Integer) + " & nItem & " as char) where tcodigopedido='" & xPedido & "' and tItem in(select tItem from " & sDetalle & " where tcodigopedido='" & sCorrelativo & "')"
                                   Cn.Execute Isql
                                Else
                                
                                              'verifica si hay propiedades en combos en el pedido original
                                        cPropiedades = Calcular("select count(*) as codigo from tcombopropiedad where tcodigopedido='" & xPedido & "'", Cn)
                                        If (cPropiedades > 0) Then
                                          '  Set rsPedidosItem = Lib.OpenRecordset(" select tpedidoini,titemini,tpedidofin,titemfin,tproducto from tpedido where tpedidoini='" & xPedido & "' and tpedidofin='" & sCorrelativo & "'", Cn)
                                            Set rsPedidosItem = Lib.OpenRecordset("select tpedidoini,titemini,tpedidofin,max(titemfin) titemfin,tproducto ,max(ncorrelativo) from tpedido where tpedidoini='" & xPedido & "' and tpedidofin='" & sCorrelativo & "' group by tpedidoini,titemini,tpedidofin,tproducto", Cn)
                                            If (rsPedidosItem.RecordCount > 0) Then
                                                rsPedidosItem.MoveFirst
                                                    For P = 0 To rsPedidosItem.RecordCount - 1
                                                                    Isql = " insert into TCOMBOPROPIEDAD (TCODIGOPEDIDO,TITEM,TITEMCOMBO,TCODIGOPROPIEDAD,TPRODUCTO,TENLACE,NINSUMO,NGASTO,NMANOOBRA,LREPLICA,NCANTIDAD,NINSUMOUNITARIO,NGASTOUNITARIO,NMANOOBRAUNITARIO) " & _
                                                                             " select '" & sCorrelativo & "','" & rsPedidosItem.Fields(3) & "',TITEMCOMBO,TCODIGOPROPIEDAD,TPRODUCTO,tenlace,ninsumo,ngasto,nmanoobra,lreplica,ncantidad,ninsumounitario,ngastounitario,nmanoobraunitario from TCOMBOPROPIEDAD " & _
                                                                             " where tcodigopedido='" & xPedido & "'   and titem='" & rsPedidosItem.Fields(1) & "'"
                                                                 ' inserta en tproductopropiedad los item del pedido original que tienen propiedades
                                                                  Cn.Execute Isql
                                      
                                                       rsPedidosItem.MoveNext
                                                    Next P
                                            End If
                                        End If
                                End If
                    
                    
                                If lCompartir = False Then
                                    Cn.Execute "update TCOMBOPROPIEDAD set tcodigopedido = '" & sCorrelativo & "', tItem=substring('000',1,3 - len(ltrim(cast(cast(tItem as Integer) + " & nItem & " as char) ))) + cast(cast(tItem as Integer) + " & nItem & " as char) where tcodigopedido='" & xPedido & "' and tItem in(select tItem from " & sDetalle & " where tcodigopedido='" & sCorrelativo & "')"
                                Else
                                      ' jala cpedido
                                      Dim rsItemCombo As New ADODB.Recordset
                                     ' Isql = "SELECT     dbo.TPEDIDO.tPedidoIni, dbo.TPEDIDO.tItemIni, dbo.TPEDIDO.tPedidoFin, dbo.TPEDIDO.tItemFin, dbo.DPEDIDO.tCodigoProducto FROM         dbo.DPEDIDO INNER JOIN  dbo.TPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.TPEDIDO.tPedidoFin AND dbo.DPEDIDO.tItem = dbo.TPEDIDO.tItemFin WHERE     (dbo.DPEDIDO.lCombinacion = 1) AND (dbo.TPEDIDO.tPedidoFin = '" & sCorrelativo & "') AND (dbo.TPEDIDO.tPedidoIni = '" & xPedido & "') "
                                    ' Isql = "SELECT     dbo.TPEDIDO.tPedidoIni, dbo.TPEDIDO.tItemIni, dbo.TPEDIDO.tPedidoFin, MAX(dbo.TPEDIDO.tItemFin) TITEMFIN, dbo.DPEDIDO.tCodigoProducto ,MAX(DBO.TPEDIDO.NCORRELATIVO) FROM         dbo.DPEDIDO INNER JOIN  dbo.TPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.TPEDIDO.tPedidoFin AND dbo.DPEDIDO.tItem = dbo.TPEDIDO.tItemFin WHERE     (dbo.DPEDIDO.lCombinacion = 1) AND (dbo.TPEDIDO.tPedidoFin = '" & sCorrelativo & "') AND (dbo.TPEDIDO.tPedidoIni = '" & xPedido & "') GROUP BY   dbo.TPEDIDO.tPedidoIni,            dbo.TPEDIDO.tItemIni,            dbo.TPEDIDO.tPedidoFin,dbo.DPEDIDO.tCodigoProducto "
                                     Isql = " SELECT  TP.tPedidoIni, TP.tItemIni,TP.tPedidoFin, MAX(TP.tItemFin) TITEMFIN, dbo.DPEDIDO.tCodigoProducto  , Max (TP.nCorrelativo) FROM    dbo.DPEDIDO INNER JOIN (   select tpedidoini,titemini,tpedidofin,max(titemfin) AS titemfin ,max(ncorrelativo)  AS NCORRELATIVO    from tpedido where tpedidoini='" & xPedido & "' and tpedidofin='" & sCorrelativo & "'    group by tpedidoini,titemini,tpedidofin) TP  ON dbo.DPEDIDO.tCodigoPedido = TP.TPedidoFin AND dbo.DPEDIDO.tItem = TP.tItemFin WHERE   (dbo.DPEDIDO.lCombinacion = 1) AND (TP.tPedidoFin = '" & sCorrelativo & "') AND (TP.tPedidoIni = '" & xPedido & "') GROUP BY   TP.tPedidoIni,            TP.tItemIni,            TP.tPedidoFin,dbo.DPEDIDO.tCodigoProducto "
                                      Set rsItemCombo = Lib.OpenRecordset(Isql, Cn)
                                      
                                      If (rsItemCombo.RecordCount > 0) Then
                                        
                                          rsItemCombo.MoveFirst
                                          For P = 0 To rsItemCombo.RecordCount - 1
                                                  Isql = " insert into cpedido (tCodigoPedido,tItem,tItemCombo,tProducto,tProductoCombo,nCantidad,tCodigoGrupo,tCodigoSubGrupo,nPrecioNeto,nImpuesto1,nImpuesto2,nImpuesto3,nVenta,nInsumo,nGasto,nManoObra,lImprimeArea,lImprime,nOrden,tObservacion,lCorte) " & _
                                                         " select '" & sCorrelativo & "', '" & rsItemCombo.Fields(3) & "' ,titemcombo, tproducto,tproductocombo,ncantidad,tcodigogrupo,tcodigosubgrupo,nprecioneto,nimpuesto1,        nImpuesto2 , nImpuesto3, nVenta, nInsumo, nGasto, nmanoobra, limprimearea, limprime, nOrden, tobservacion, lCorte " & _
                                                         " From cpedido " & _
                                                         " where tcodigopedido='" & xPedido & "' and titem='" & rsItemCombo.Fields(1) & "'"
                                                  Cn.Execute Isql
                                                  If lCompartir = False Then
                                                      Isql = "delete from cpedido where tcodigopedido='" & rsItemCombo.Fields(0) & "' and titem='" & rsItemCombo.Fields(1) & "' and tcodigopedido+titem not in(select tcodigopedido+titem from dpedido where tcodigopedido='" & rsItemCombo.Fields(0) + rsItemCombo.Fields(1) & "') "
                                                      Cn.Execute Isql
                                                  End If
                                              rsItemCombo.MoveNext
                                          Next P
                                      End If
                                End If
                                                                
                            If lClub = True And pais = "002" Then
                                
                                Dim iRsDPEDIDO As Recordset
                                Isql = "Select * from DPEDIDO Where tCodigoPedido='" & sCorrelativo & "'"
                                Set iRsDPEDIDO = Lib.OpenRecordset(Isql, Cn)
    
                                For xp = 0 To iRsDPEDIDO.RecordCount - 1
                                    Dim iCodigoProducto As String
                                    Dim iItem As String
                                    Dim iCantidad As Double
                                    
                                    iItem = iRsDPEDIDO!tItem
                                    iCodigoProducto = iRsDPEDIDO!tCodigoProducto
                                    iCantidad = iRsDPEDIDO!nCantidad
    
                                    Cn.Execute "USP_CALCULA_PRECIO '" & sCorrelativo & "','" & iItem & "','" & iCodigoProducto & "','" & iTipoPedido & "'," & iCantidad & ""
    
                                    iRsDPEDIDO.MoveNext
                                Next xp
                                
                                Dim iDescuento As Double
                                iDescuento = Calcular("select ndescuento as codigo from MPEDIDO where tcodigopedido = '" & sCorrelativo & "'", Cn)
                   
    
                                If iDescuento > 0 Then
                                    CalculaDescuento sCorrelativo
                                End If
                                
                            End If
                                     
                     
                            RsPedido.MoveNext
                    Loop
               End If
               wEnter = True
                Cn.CommitTrans
               Unload Me
               Exit Sub
ERRORDIVISION:
    Cn.RollbackTrans
    MsgBox Error, vbCritical
Unload Me
    Screen.MousePointer = vbDefault
    
    Exit Sub
          Case Is = 1 ' Cancelar
               Unload Me
               
          Case Is = 2 ' Cantidad
               Dim i As Integer
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And val(sDescrip) > 0 Then
                   
                  If nContador > val(sDescrip) Then
                     If val(Calcular("select count(tCodigoPedido) as Codigo from " & sDetalle & " where ESTADO='X' and cast(tcodigopedido as integer) > " & val(sDescrip), Cn)) > 0 Then
                        MsgBox "Los Pedidos tienen items, no se puede eliminar", vbCritical, sMensaje
                        Exit Sub
                     End If
                  End If
                              
                  Cn.Execute "delete from " & sCabecera
                  nContador = 0
                  For i = 1 To val(sDescrip)
                      nContador = nContador + 1
                      sCorrelativo = Mid("000", 1, 3 - Len(LTrim(str(nContador)))) & LTrim(str(nContador))
                      If txtMesa.Caption <> "" Then
                         sObservacion = txtMesa.Caption & " - " & sCorrelativo
                      ElseIf txtObservacion.Caption <> "" Then
                         sObservacion = LTrim(sObservacion) & " - " & sCorrelativo
                      Else
                         sObservacion = LTrim(txtOrden.Caption) & " - " & sCorrelativo
                      End If
                      Cn.Execute "insert into " & sCabecera & " (tCodigoPedido, nCorrelativo, tMesa, tObservacion) values ('" & sCorrelativo & "', '" & txtOrden.Caption & "','','" & sObservacion & "') "
                  Next i
                  RsPedido.Requery
                  
               End If
               
          Case Is = 3 ' Aumentar
               nContador = nContador + 1
               sCorrelativo = Mid("000", 1, 3 - Len(LTrim(str(nContador)))) & LTrim(str(nContador))
               If txtMesa.Caption <> "" Then
                  sObservacion = txtMesa.Caption & " - " & sCorrelativo
               ElseIf txtObservacion.Caption <> "" Then
                  sObservacion = LTrim(txtObservacion) & " - " & sCorrelativo
               Else
                  sObservacion = LTrim(txtOrden.Caption) & " - " & sCorrelativo
               End If
               Cn.Execute "insert into " & sCabecera & " (tCodigoPedido, nCorrelativo, tMesa, tObservacion) values ('" & sCorrelativo & "', '" & txtOrden.Caption & "','','" & sObservacion & "') "
               RsPedido.Requery
               
          Case Is = 4 ' Disminuir
               sCorrelativo = Mid("000", 1, 3 - Len(LTrim(str(nContador)))) & LTrim(str(nContador))
               
               If val(Calcular("select count(tCodigoPedido) as Codigo from " & sDetalle & " where Estado='X' and tCodigoPedido='" & sCorrelativo & "'", Cn)) > 0 Then
                  MsgBox "El Pedido tiene items, no se puede eliminar", vbCritical, sMensaje
                  Exit Sub
               End If
               
               Cn.Execute "delete from " & sCabecera & " where tCodigoPedido='" & sCorrelativo & "'"
               RsPedido.Requery
               nContador = IIf(nContador = 0, 0, nContador - 1)
          
          Case Is = 5 ' Agregar Pedido
               sTemp = ""
               If sModulo = "INFOREST" Then
                  If lMCPV Then
                     Isql = "select *, 'Importar' as Caso from vPedidoGrilla " & _
                            "where Codigo <> '" & xPedido & "' and tUsuario='" & sUsuario & "'  and tEstadoPedido='01'  and Codigo not in (select tCodigoPedido from " & sCabecera & ") "
                  Else
                     Isql = "select *, 'Importar' as Caso from vPedidoGrilla " & _
                            "where Codigo <> '" & xPedido & "' and tCaja='" & sCaja & "'  and tEstadoPedido='01'  and Codigo not in (select tCodigoPedido from " & sCabecera & ") "
                  End If
               ElseIf sModulo = "PVCORP" Then
                  Isql = "select *, 'Importar' as Caso from vPedidoGrilla " & _
                         "where Codigo <> '" & xPedido & "' and tCaja='" & sCaja & "'  and tEstadoPedido='01' and tTipoPedido='04' and Codigo not in (select tCodigoPedido from " & sCabecera & ") "
               Else
                  If lSomelier Then
                     Isql = "select *, 'Importar' as Caso from vPedidoGrilla " & _
                            "where Codigo <> '" & xPedido & "' and tTurno='" & sTurno & "' and tEstadoPedido='01'  and Codigo not in (select tCodigoPedido from " & sCabecera & ") "
                  Else
                     Isql = "select *, 'Importar' as Caso from vPedidoGrilla " & _
                            "where Codigo <> '" & xPedido & "' and tTurno='" & sTurno & "' and tUsuario='" & Right(sUsuario, 15) & "' and tEstadoPedido='01'  and Codigo not in (select tCodigoPedido from " & sCabecera & ") "
                  End If
               End If
               
               If sModulo = "PVCORP" Then
                  Call ConfGrilla(7, frmBusquedaRapida.grdGrilla, "Función", 2, "Caso", 1000, 0, 0, "", _
                                                                  "Caja", 2, "tCaja", 600, 2, 0, "", _
                                                                  "Mesa", 2, "Mesa", 1200, 0, 0, "", _
                                                                  "Guia", 2, "tObservacion", 2000, 0, 0, "", _
                                                                  "Pedido", 2, "Descripcion", 1100, 0, 0, "", _
                                                                  "Monto", 2, "Suma", 1100, 1, 0, "###,##0.00", _
                                                                  "Usuario", 2, "tUsuario", 1200, 0, 0, "")
                  frmBusquedaRapida.nPredeterm = 3
               Else
                  Call ConfGrilla(7, frmBusquedaRapida.grdGrilla, "Función", 2, "Caso", 1000, 0, 0, "", _
                                                                  "Caja", 2, "tCaja", 600, 2, 0, "", _
                                                                  "Mesa", 2, "Mesa", 1200, 0, 0, "", _
                                                                  "Observacion", 2, "tObservacion", 2000, 0, 0, "", _
                                                                  "Pedido", 2, "Descripcion", 1100, 0, 0, "", _
                                                                  "Monto", 2, "Suma", 1100, 1, 0, "###,##0.00", _
                                                                  "Usuario", 2, "tUsuario", 1200, 0, 0, "")
                  frmBusquedaRapida.nPredeterm = 2
               
               End If
               frmBusquedaRapida.Show vbModal
               
               If wEnter Then
                  Dim RsTemp As Recordset
                  Set RsTemp = Lib.OpenRecordset("SELECT nCorrelativo, dbo.MPEDIDO.tMesa, TMESA.tDetallado AS Mesa, dbo.MPEDIDO.tObservacion FROM MPEDIDO LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa where tCodigoPedido='" & sCodigo & "'", Cn)
                  
                  Dim xTipoPedido As String
                  xTipoPedido = Calcular("Select tTipoPedido as codigo From MPEDIDO where tCodigoPedido='" & xPedido & "'", Cn)
                  
                  Dim yTipoPedido As String
                  yTipoPedido = Calcular("Select tTipoPedido as codigo From MPEDIDO where tCodigoPedido='" & sCodigo & "'", Cn)
                                    
                  If xTipoPedido = yTipoPedido Then
                     Cn.Execute "insert into " & sCabecera & " (tCodigoPedido, nCorrelativo, tMesa, tObservacion) values ('" & sCodigo & "', '" & RsTemp!nCorrelativo & "','" & IIf(IsNull(RsTemp!tMesa), "", RsTemp!tMesa) & "','" & IIf(IsNull(RsTemp!tObservacion), "", RsTemp!tObservacion) & "') "
                  Else
                       If MsgBox("El pedido seleccionado pertenece a un canal de venta diferente al pedido de origen esto generaria un descuadre en la liquidación de cajero, Desea realizar la transferencia ? ", vbInformation + vbYesNo, sMensaje) = vbYes Then
                          Cn.Execute "insert into " & sCabecera & " (tCodigoPedido, nCorrelativo, tMesa, tObservacion) values ('" & sCodigo & "', '" & RsTemp!nCorrelativo & "','" & IIf(IsNull(RsTemp!tMesa), "", RsTemp!tMesa) & "','" & IIf(IsNull(RsTemp!tObservacion), "", RsTemp!tObservacion) & "') "
                       Else
                          Exit Sub
                       End If
                  End If
                                                                     
                  Set RsTemp = Nothing
                  RsPedido.Requery
                  wEnter = False
               End If
   End Select

End Sub

Private Sub Form_Load()
 
    'Pedido
   sCabecera = dbTemporal(sCaja, 4, "tCodigoPedido", "nVarchar(10)", _
                                    "nCorrelativo", "int", _
                                    "tMesa", "nVarChar(3)", _
                                    "tObservacion", "nVarChar(250)")
        
   'Detalle
   sDetalle = dbTemporal(sCaja, 58, "tCodigoPedido", "nVarChar(10)", _
                                     "tItem", "nVarChar(3)", _
                                     "tTipoPedido", "nVarChar(2)", _
                                     "tCodigoProducto", "nVarChar(7)", _
                                     "tCodigoGrupo", "nVarChar(2)", _
                                     "tCodigoSubGrupo", "nVarChar(4)", _
                                     "tMoneda", "nVarChar(3)", _
                                     "nPrecioNeto", "Float", "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", "nPrecioVenta", "Float", _
                                     "nRecargo", "Float", "nDescuento", "Float", "nPrecioOficial", "Float", _
                                     "nCantidad", "Float", _
                                     "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                     "nVenta", "Float", _
                                     "tObservacion", "nVarChar(255)", _
                                     "tCortesia", "nVarChar(4)", _
                                     "lImprime", "Bit", _
                                     "tEstadoItem", "nVarChar(3)", _
                                     "tArea", "nVarChar(3)", "lCombinacion", "Bit", "nCombinacion", "Smallint", "lImprimeArea", "Bit", _
                                     "tFacturado", "nVarChar(1)", "tDocumento", "nVarChar(15)", "lTransferido", "Bit", "tComanda", "nVarChar(10)", "fRegistro", "smalldatetime", "tMozoD", "nVarChar(4)", "tUsuarioD", "nVarChar(15)", _
                                     "nInsumo", "Float", "nGasto", "Float", "nManoObra", "Float", "nOrden", "Int", "lCorte", "bit", "tPosicion", "nVarChar(2)", "fEnvio", "smalldatetime", "nEnvio", "int", "tUnidadNegocio", "nVarchar(2)", "tOferta", "nVarChar(5)", "tAutorizaOferta", "nVarChar(15)", "tsubalmacen", "nvarchar(3)", "fDiaContable", "smalldatetime", _
                                     "lAtendidoC", "Bit", "fAtendidoC", "datetime", "tUsuarioAtendio", "nVarChar(15)", "lCantadoC", "Bit", "fCantadoC", "datetime", "lTipoEnvio", "Bit", "Estado", "nVarChar(1)", "tCajaD", "nVarchar(3)", "tUsuarioAutorizaEnvio", "nVarchar(15)", "tPerfilUsuarioAutoriza", "nVarchar(2)")
                                     
  'COMBO CESAR
   sDetalleCombo = dbTemporal(sCaja, 25, "tCodigoPedido", "nVarChar(10)", _
                                     "tItem", "nVarChar(3)", _
                                     "tItemCombo", "nVarChar(3)", _
                                     "tProducto", "nVarChar(7)", _
                                     "tProductoCombo", "nVarChar(7)", _
                                     "nCantidad", "Float", _
                                     "tCodigoGrupo", "nVarChar(2)", _
                                     "tCodigoSubGrupo", "nVarChar(4)", _
                                     "nPrecioNeto", "Float", _
                                     "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                     "nVenta", "Float", _
                                     "nInsumo", "Float", "nGasto", "Float", "nManoObra", "Float", _
                                     "lImprimeArea", "Bit", _
                                     "lImprime", "Bit", _
                                     "nOrden", "Int", "tObservacion", "nVarChar(255)", "lCorte", "bit", _
                                     "lAtendidoC", "Bit", "fAtendidoC", "datetime", "tUsuarioAtendio", "nVarChar(15)", "Estado", "nVarChar(1)")
                                     
    
    ' disgregar pedido
    sProductoPropiedad = dbTemporal(sCaja, 12, "tCodigoPedido", "nvarchar (20)", "tItem", "nvarchar(6)", "tCodigoPropiedad", "nvarchar(8)", "tProducto", "nvarchar(14)", "tEnlace", "nvarchar(14)", "nInsumo", "float", "nGasto", "float", "nManoObra", "float", "ncantidad", "float", "nInsumoUnitario", "float", "nGastoUnitario", "float", "nManoObraUnitario", "float")
   
    Cn.Execute "Insert into " & sDetalle & " select tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, tMoneda, nPrecioNeto,nPrecioImpuesto1, nPrecioImpuesto2,  nPrecioImpuesto3,  nPrecioVenta,    nRecargo,  nDescuento,  nPrecioOficial, nCantidad,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, tObservacion, tCortesia, lImprime, tEstadoItem, tArea, lCombinacion,  nCombinacion, lImprimeArea,   tFacturado,  tDocumento, lTransferido,  tComanda, fRegistro, tMozoD,tUsuarioD, nInsumo, nGasto, nManoObra,  nOrden,lCorte,tPosicion, fEnvio,nEnvio, tUnidadNegocio, tOferta, tAutorizaOferta,  tsubalmacen,fDiaContable, lAtendidoC, fAtendidoC, tUsuarioAtendio, lCantadoC, fCantadoC, lTipoEnvio, ' ' as Estado,tcajad, tUsuarioAutorizaEnvio, tPerfilUsuarioAutoriza From DPEDIDO where tCodigoPedido='" & sPedido & "'"
   'Cn.Execute "Insert into " & sDetalle & " select *, ' ' as Estado From DPEDIDO where tCodigoPedido='" & sPedido & "'"
    itemDisgregado = ""
    
   'COMBO CESAR
   Cn.Execute "Insert into " & sDetalleCombo & " select tCodigoPedido, tItem, tItemCombo, tProducto, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto,nImpuesto1, nImpuesto2,  nImpuesto3,   nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden, tObservacion, lCorte, lAtendidoC, fAtendidoC, tUsuarioAtendio, ' ' As Estado From CPEDIDO where tCodigoPedido='" & sPedido & "'"
    
   'Configura Cabecera
   Isql = "SELECT " & sCabecera & ".tCodigoPedido, " & sCabecera & ".nCorrelativo, dbo.TMESA.tDetallado AS Mesa, " & sCabecera & ".tObservacion, COUNT(dbo." & sDetalle & ".tCodigoProducto) AS nItem, isnull(SUM(dbo." & sDetalle & ".nVenta),0) AS nVenta " & _
          "FROM " & sCabecera & " LEFT OUTER JOIN dbo.TMESA ON " & sCabecera & ".tMesa = dbo.TMESA.tCodigoMesa LEFT OUTER JOIN dbo." & sDetalle & " ON " & sCabecera & ".tCodigoPedido = dbo." & sDetalle & ".tCodigoPedido " & _
          "GROUP BY " & sCabecera & ".tCodigoPedido, " & sCabecera & ".nCorrelativo, dbo.TMESA.tDetallado, " & sCabecera & ".tObservacion"
   
   Set RsPedido = Lib.OpenRecordset(Isql, Cn)
   Call ConfGrilla(5, grdPedido, "Orden", 2, "nCorrelativo", 600, 1, 0, "", _
                                 "Mesa", 2, "Mesa", 1290, 0, 0, "", _
                                 "Observación", 2, "tObservacion", 1500, 0, 0, "", _
                                 "Items", 2, "nItem", 600, 1, 0, "##0", _
                                 "Total", 2, "nVenta", 1000, 1, 0, "#,###,##0.00")
                                      
   Set grdPedido.DataSource = RsPedido
      
   'Configura Detalle
   Isql = "SELECT [" & sDetalle & "].*, TPRODUCTO.tResumido AS Producto, [nVenta]*[nCantidad] AS Subtotal " & _
          "FROM TPRODUCTO INNER JOIN [" & sDetalle & "] ON TPRODUCTO.tCodigoProducto = [" & sDetalle & "].tCodigoProducto where Estado=' '"
          
   Set RsDetalle1 = Lib.OpenRecordset(Isql, Cn)
   Call ConfGrilla(4, grdDetalle1, "Producto", 2, "Producto", 1980, 0, 0, "", _
                                   "Precio", 2, "nPrecioVenta", 900, 1, 0, "###,###,##0.00", _
                                   "Cant", 2, "nCantidad", 600, 1, 0, "#,##0.00", _
                                   "SubTotal", 2, "nVenta", 1000, 1, 0, "###,###,##0.00")
                                                                      
   Set grdDetalle1.DataSource = RsDetalle1
   
   'Detalle 2
   Isql = "SELECT [" & sDetalle & "].*, TPRODUCTO.tResumido AS Producto, [nVenta]*[nCantidad] AS Subtotal " & _
          "FROM TPRODUCTO INNER JOIN [" & sDetalle & "] ON TPRODUCTO.tCodigoProducto = [" & sDetalle & "].tCodigoProducto where Estado = 'X'"
   Set RsDetalle2 = Lib.OpenRecordset(Isql, Cn)
   Call ConfGrilla(4, grdDetalle2, "Producto", 2, "Producto", 1980, 0, 0, "", _
                                   "Precio", 2, "nPrecioVenta", 900, 1, 0, "###,###,##0.00", _
                                   "Cant", 2, "nCantidad", 600, 1, 0, "#,##0.00", _
                                   "SubTotal", 2, "nVenta", 1000, 1, 0, "###,###,##0.00")
   Set grdDetalle2.DataSource = RsDetalle2
   
   xPedido = sPedido
   grdPedido.Caption = "Pedido Origen " & sPedido
   grdDetalle1.Caption = "Detalle del Pedido " & sPedido
   nContador = 0
               
   CalculaMonto
   Screen.MousePointer = vbDefault
   lCompartir = False
   
   
   xDisgregar = False
   itemDisgregado = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sCabecera
   Cn.Execute "drop table " & sDetalle
   Set RsPedido = Nothing
   Set RsDetalle1 = Nothing
   Set RsDetalle2 = Nothing
   Set frmDivision = Nothing
End Sub

Public Sub CalculaMonto()
    txtMonto.Caption = Format(Calcular("SELECT Sum(nVenta) AS Codigo From " & sDetalle & " where Estado = ' '", Cn), "###,###,###,##0.00")
    txtItem.Caption = RsDetalle1.RecordCount
End Sub

Private Sub grdPedido_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   If RsPedido.RecordCount > 0 Then
      grdDetalle2.Caption = "Detalle del Pedido " & grdPedido.Columns(2).Text
      RsDetalle2.Filter = "tCodigoPedido='" & RsPedido!tCodigoPedido & "'"
      txtMonto2.Caption = Format(RsPedido!nVenta, "###,##0.00")
      txtItem2.Caption = Format(RsPedido!nItem, "###,##0")
      CalculaMonto
   End If
End Sub

'25/06
Public Sub usuarioautoriza(X As String)
    sUsuarioAutoriza = X
End Sub

'diaContable
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
End Function
'diaContable



Private Function CalculaDescuento(xcodigopedido As String) As Boolean

    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    Dim sCodigoDescuento As String
    Dim RsDetalle As ADODB.Recordset
    Dim xDescuento As Double
    Dim nPVenta As Double
    Dim Acumulado As Double
    
    Dim nImpuesto1 As Double
    Dim nImpuesto2 As Double
    Dim nImpuesto3 As Double
    
    Dim nPBase As Double
    Dim nRecargo As Double
    Dim sCortesia As String
    
    sCortesia = ""
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM dbo.DPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoPedido='" & xcodigopedido & "' and lDescuento=1", Cn)
    sCodigoDescuento = Calcular("select tDescuento as codigo from MPEDIDO  where tcodigopedido= '" & xcodigopedido & "' ", Cn)
    xDescuento = Calcular("select nDescuento as codigo from MPEDIDO  where tcodigopedido= '" & xcodigopedido & "' ", Cn)
    
    Isql = "select * from vPedidoDetalle where tcodigopedido= '" & xcodigopedido & "'"
    Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
    
    If RsDetalle.RecordCount <> 0 Then

         RsDetalle.MoveFirst
        
         Do While Not RsDetalle.EOF

           nPVenta = 0
           
           lAcumulable = True
           
           nOferta = 0
           Acumulado = 0
                      
           If RsDetalle!lDescuento And lAcumulable = True Then
              If Calcular("select lRatio as Codigo FROM vMotivoDescuento where Codigo='" & sCodigoDescuento & "'", Cn) Then
                 nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((RsDetalle!nPrecioOficial - nOferta) * xDescuento / 100)
              Else
                 Dim xPorc As Double
                 xPorc = (RsDetalle!nPrecioOficial - nOferta) * RsDetalle!nCantidad * 100 / nSuma
                 nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((xPorc * xDescuento / 100) / RsDetalle!nCantidad)
              End If
              
               Select Case pais ' ok
                 Case "001" 'Bolivia
                              Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                              Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                              Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                              Acumulado = (Acumulado / 100)
                             
                              nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                              nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                              nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                              nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
               
                 Case Else 'Peru, Ecuador
                              Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                              Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                              Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                              Acumulado = 1 + (Acumulado / 100)
                             
                              nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                              nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                              nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                              nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
               End Select
               
              Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                     "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                     "nRecargo = " & nRecargo & ", " & _
                     "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                     "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                     "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                     "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                     "nPrecioVenta = " & nPVenta & ", " & _
                     "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                     "nCantidad = " & RsDetalle!nCantidad & ", " & _
                     "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                     "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                     "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                     "tCortesia = '" & sCortesia & "' " & _
                     "where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido ='" & xcodigopedido & "'"
                     Cn.Execute Isql
        End If
    RsDetalle.MoveNext
    Loop
End If

End Function

