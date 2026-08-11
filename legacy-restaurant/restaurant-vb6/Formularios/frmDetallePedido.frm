VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDetallePedido 
   BorderStyle     =   0  'None
   Caption         =   "Detalle de pedido"
   ClientHeight    =   16200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16200
   ScaleWidth      =   19680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrConsultaCombo 
      Interval        =   100
      Left            =   720
      Top             =   11520
   End
   Begin VB.Timer tmrConsulta 
      Interval        =   100
      Left            =   120
      Top             =   11520
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   14895
      Begin VB.PictureBox pictxCliente 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4080
         ScaleHeight     =   375
         ScaleWidth      =   2655
         TabIndex        =   43
         Top             =   240
         Width           =   2655
      End
      Begin VB.PictureBox pictxObs 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   11880
         ScaleHeight     =   375
         ScaleWidth      =   2775
         TabIndex        =   42
         Top             =   240
         Width           =   2775
      End
      Begin VB.PictureBox pictxFecha 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7680
         ScaleHeight     =   375
         ScaleWidth      =   2415
         TabIndex        =   41
         Top             =   240
         Width           =   2415
      End
      Begin VB.VScrollBar vscrPedidos 
         Height          =   3330
         Left            =   14160
         TabIndex        =   40
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox pictxNumTotal 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   12840
         ScaleHeight     =   495
         ScaleWidth      =   1935
         TabIndex        =   18
         Top             =   4560
         Width           =   1935
      End
      Begin VB.PictureBox pictxTotal 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   12000
         ScaleHeight     =   495
         ScaleWidth      =   855
         TabIndex        =   17
         Top             =   4560
         Width           =   855
      End
      Begin VB.PictureBox pictxNumDescuento 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   9960
         ScaleHeight     =   495
         ScaleWidth      =   1935
         TabIndex        =   16
         Top             =   4560
         Width           =   1935
      End
      Begin VB.PictureBox pictxDescuento 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   8520
         ScaleHeight     =   495
         ScaleWidth      =   1455
         TabIndex        =   15
         Top             =   4560
         Width           =   1455
      End
      Begin VB.PictureBox pictxNumIcbper 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   6480
         ScaleHeight     =   495
         ScaleWidth      =   1935
         TabIndex        =   14
         Top             =   4560
         Width           =   1935
      End
      Begin VB.PictureBox pictxIcbper 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5280
         ScaleHeight     =   495
         ScaleWidth      =   1215
         TabIndex        =   13
         Top             =   4560
         Width           =   1215
      End
      Begin VB.PictureBox pictxNumProductos 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2175
         TabIndex        =   12
         Top             =   4560
         Width           =   2175
      End
      Begin TrueOleDBGrid80.TDBGrid grdPedidos 
         Height          =   3330
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Width           =   13920
         _ExtentX        =   24553
         _ExtentY        =   5874
         _LayoutType     =   4
         _RowHeight      =   26
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   0
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   16777215
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MarqueeUnique   =   0   'False
         RowDividerStyle =   7
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   15263202
         RowSubDividerColor=   16777215
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   16777215
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.locked=0,.bold=0"
         _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
         _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Arial"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
         _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Arial"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&,.bold=0"
         _StyleDefs(19)  =   ":id=6,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(20)  =   ":id=6,.fontname=Arial"
         _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1"
         _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1"
         _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1"
         _StyleDefs(27)  =   "Splits(0).Style:id=87,.parent=1"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=89,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=91,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8,.namedParent=38"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
         _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
         _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.appearance=0,.borderSize=1"
         _StyleDefs(51)  =   ":id=34,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(52)  =   ":id=34,.fontname=Arial"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
         _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=1200,.italic=0,.underline=0"
         _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
         _StyleDefs(59)  =   ":id=36,.fontname=Arial"
         _StyleDefs(60)  =   "Named:id=37:Caption"
         _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(62)  =   "Named:id=38:HighlightRow"
         _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000008&,.appearance=0"
         _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=0,.fontsize=900,.italic=0,.underline=0"
         _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=38,.fontname=Arial"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame fraPedidos 
         BackColor       =   &H00D8D3CF&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   14655
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10320
         TabIndex        =   22
         Top             =   285
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6840
         TabIndex        =   21
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3240
         TabIndex        =   20
         Top             =   285
         Width           =   855
      End
      Begin VB.Label lblCodigoPedido 
         BackColor       =   &H8000000E&
         Caption         =   "Pedido 24000000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006BA838&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   14895
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2DEDB&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   14535
         Begin VB.VScrollBar VScroll2 
            Height          =   2175
            Left            =   6600
            TabIndex        =   37
            Top             =   600
            Width           =   615
         End
         Begin VB.Frame fraObservacionCombo 
            BackColor       =   &H00D8D3CF&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   7440
            TabIndex        =   34
            Top             =   1920
            Width           =   6855
            Begin VB.TextBox txtObservacionCombo 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Width           =   6735
            End
         End
         Begin VB.Frame fraPropiedadCombo 
            BackColor       =   &H00D8D3CF&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   7440
            TabIndex        =   31
            Top             =   600
            Width           =   6855
            Begin VB.TextBox txtPropiedadCombo 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Width           =   6735
            End
         End
         Begin TrueOleDBGrid80.TDBGrid grdCombo 
            Height          =   2175
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   3836
            _LayoutType     =   4
            _RowHeight      =   26
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   0
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerColor=   16777215
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MarqueeUnique   =   0   'False
            RowDividerStyle =   7
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
            RowDividerColor =   15263202
            RowSubDividerColor=   16777215
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            CollapseColor   =   16777215
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000009&,.locked=0,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
            _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Arial"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
            _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=2,.fontname=Arial"
            _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&,.bold=0"
            _StyleDefs(19)  =   ":id=6,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(20)  =   ":id=6,.fontname=Arial"
            _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(27)  =   "Splits(0).Style:id=87,.parent=1"
            _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
            _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=89,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=91,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8,.namedParent=38"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
            _StyleDefs(43)  =   "Named:id=33:Normal"
            _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.appearance=0,.borderSize=1"
            _StyleDefs(51)  =   ":id=34,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(52)  =   ":id=34,.fontname=Arial"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
            _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=1200,.italic=0,.underline=0"
            _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=36,.fontname=Arial"
            _StyleDefs(60)  =   "Named:id=37:Caption"
            _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(62)  =   "Named:id=38:HighlightRow"
            _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000008&,.appearance=0"
            _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=0,.fontsize=900,.italic=0,.underline=0"
            _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
            _StyleDefs(66)  =   ":id=38,.fontname=Arial"
            _StyleDefs(67)  =   "Named:id=39:EvenRow"
            _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(69)  =   "Named:id=40:OddRow"
            _StyleDefs(70)  =   ":id=40,.parent=33"
            _StyleDefs(71)  =   "Named:id=41:RecordSelector"
            _StyleDefs(72)  =   ":id=41,.parent=34"
            _StyleDefs(73)  =   "Named:id=42:FilterBar"
            _StyleDefs(74)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E2DEDB&
            Caption         =   "Observación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7560
            TabIndex        =   33
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label lblPropiedades 
            BackColor       =   &H00E2DEDB&
            Caption         =   "Propiedades"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7560
            TabIndex        =   30
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E2DEDB&
            Caption         =   "Combos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   14535
         Begin VB.Frame fraObservacion 
            BackColor       =   &H00D8D3CF&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   7560
            TabIndex        =   7
            Top             =   360
            Width           =   6735
            Begin VB.TextBox txtObservacion 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   0
               TabIndex        =   8
               Top             =   0
               Width           =   6615
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   7680
            TabIndex        =   6
            Top             =   720
            Width           =   3135
         End
         Begin VB.Frame fraPropiedad 
            BackColor       =   &H00D8D3CF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   855
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   7215
            Begin VB.TextBox txtPropiedad 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   7095
            End
         End
         Begin VB.Label Label6 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7680
            TabIndex        =   10
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label Label5 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Propiedades"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      Height          =   11040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.Frame Frame8 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         TabIndex        =   23
         Top             =   10080
         Width           =   14895
         Begin VB.PictureBox picReactivarPedido 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   660
            Left            =   11640
            Picture         =   "frmDetallePedido.frx":0000
            ScaleHeight     =   660
            ScaleWidth      =   2580
            TabIndex        =   26
            Top             =   120
            Width           =   2580
         End
         Begin VB.PictureBox Picture9 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   660
            Left            =   8280
            Picture         =   "frmDetallePedido.frx":2059
            ScaleHeight     =   660
            ScaleWidth      =   2580
            TabIndex        =   25
            Top             =   120
            Width           =   2580
         End
         Begin VB.PictureBox picVolver 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   480
            Picture         =   "frmDetallePedido.frx":3AFF
            ScaleHeight     =   660
            ScaleWidth      =   1920
            TabIndex        =   24
            Top             =   120
            Width           =   1920
         End
      End
      Begin VB.Label lblCerrar 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14400
         TabIndex        =   36
         Top             =   0
         Width           =   495
      End
      Begin VB.Label txtTitulo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalle de pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   14415
      End
   End
End
Attribute VB_Name = "frmDetallePedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrilla As Recordset
Dim RsTitulo As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsCombo As Recordset
Dim RsComboPropiedad As Recordset

Dim simboloSoles As String
Private Sub Form_Load()
    Call OrganizaComponentes
    Centrar Me

    inicializarComponentes
    
    lblCodigoPedido.Caption = "Pedido " & sPedido
    
    sVar3 = Calcular("SELECT rTRIM(ISNULL(tApellido, '') + ' ' + ISNULL(tNombre, '')) AS Codigo FROM MPEDIDO INNER JOIN TDELIVERY ON TDELIVERY.tCodigoDelivery = MPEDIDO.tClienteDelivery WHERE tCodigoPedido = '" & sPedido & "'", Cn)
    sVar3 = IIf(sVar3 = "0", "", sVar3)

    'pictxCliente
    PicboxText pictxCliente, sVar3, ""
    
    'Logica basado en frmPedido
    
    'Screen.MousePointer = vbHourglass
    simboloSoles = "S/. "
    
    'Ingrese el Titulo
    txtTitulo.Caption = "Detalle de pedido - Caja : " & sCaja & " - " & Trim(sUsuario) & " "
    'Ingrese el SubTitulo
    'If sTipo = "CajaRapida" Then
    '   TxtTotal.Caption = Format(Calcular("SELECT sum(nVenta) AS Codigo From " & frmCajaRapida.sDetalle & " WHERE tEstadoItem='N'", Cn), "#,###,##0.00")
    'Else
       grdPedidos.Caption = " Pedido : " & sPedido
       PicboxText pictxNumTotal, simboloSoles & Format(Calcular("SELECT sum(DPEDIDO.nVenta) AS Codigo From DPEDIDO WHERE DPEDIDO.tCodigoPedido='" & sPedido & "' AND DPEDIDO.tEstadoItem='N'", Cn), "#,###,##0.00"), ""
       
       
    'End If
    
    'Ingresar la Vista de la Grilla
    'If sTipo = "CajaRapida" Then
    '   Isql = "SELECT " & frmCajaRapida.sComboDetalle & ".tCodigoPedido, " & frmCajaRapida.sComboDetalle & ".tProducto, " & frmCajaRapida.sComboDetalle & ".tItem, " & frmCajaRapida.sComboDetalle & ".tItemCombo, " & frmCajaRapida.sComboDetalle & ".tProductoCombo, " & frmCajaRapida.sComboDetalle & ".nCantidad, " & frmCajaRapida.sComboDetalle & ".tCodigoGrupo, " & frmCajaRapida.sComboDetalle & ".tCodigoSubGrupo, dbo.TPRODUCTO.tDetallado AS Producto, dbo.MPEDIDO.tEstadoPedido, dbo.MPEDIDO.tCaja, " & frmCajaRapida.sComboDetalle & ".lImprimeArea, " & frmCajaRapida.sComboDetalle & ".lImprime, " & frmCajaRapida.sComboDetalle & ".nOrden, CONVERT(bit, ISNULL(DATALENGTH(" & frmCajaRapida.sComboDetalle & ".tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) AS lPropiedad, " & frmCajaRapida.sComboDetalle & ".tObservacion, " & frmCajaRapida.sComboDetalle & ".lCorte " & _
    '          "FROM " & frmCajaRapida.sComboDetalle & " LEFT OUTER JOIN (SELECT tItem, tItemCombo, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END AS nPropiedad From " & frmCajaRapida.sComboPropiedad & " " & _
    '          "GROUP BY tItem, tItemCombo) AS T1 ON " & frmCajaRapida.sComboDetalle & ".tItemCombo = T1.tItemCombo AND " & frmCajaRapida.sComboDetalle & ".tItem = T1.tItem LEFT OUTER JOIN dbo.TPRODUCTO ON " & frmCajaRapida.sComboDetalle & ".tProductoCombo = dbo.TPRODUCTO.tCodigoProducto LEFT OUTER JOIN dbo.MPEDIDO ON " & frmCajaRapida.sComboDetalle & ".tCodigoPedido = dbo.MPEDIDO.tCodigoPedido "
    '
    '   Set RsCombo = Lib.OpenRecordset(Isql, Cn)
    '   Set RsTitulo = Lib.OpenRecordset("select nAdulto, nNino from MPEDIDO where tCodigoPedido = ''", Cn)
    '
    '   Isql = "SELECT " & frmCajaRapida.sDetalle & ".*, dbo.TPRODUCTO.tDetallado AS Producto, dbo.vCortesia.Descripcion AS Cortesia, CASE " & frmCajaRapida.sDetalle & ".nDescuento WHEN 0 THEN 0 ELSE " & frmCajaRapida.sDetalle & ".nDescuento * 100 / " & frmCajaRapida.sDetalle & ".nPrecioOficial END AS Descuento, " & _
    '          "dbo.TPRODUCTO.lDescuento AS lDescuento, dbo.TPRODUCTO.lModificable AS lModificable, CONVERT(bit, ISNULL(DATALENGTH(" & frmCajaRapida.sDetalle & ".tObservacion), 0)) AS lObservacion, ISNULL(T1.nPropiedad, 0) As lPropiedad, " & frmCajaRapida.sDetalle & ".lCorte " & _
    '          "FROM " & frmCajaRapida.sDetalle & " LEFT OUTER JOIN (SELECT tItem, CASE WHEN COUNT(tProducto) > 0 THEN 1 ELSE 0 END as nPropiedad FROM " & frmCajaRapida.sProductoPropiedad & " GROUP by tItem ) as T1 " & _
    '          "ON " & frmCajaRapida.sDetalle & ".tItem = T1.tItem LEFT OUTER JOIN " & _
    '          "dbo.vCortesia ON " & frmCajaRapida.sDetalle & ".tCortesia = dbo.vCortesia.Codigo LEFT OUTER JOIN dbo.TPRODUCTO ON " & frmCajaRapida.sDetalle & ".tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto " & _
    '          "ORDER BY " & frmCajaRapida.sDetalle & ".tCodigoPedido, " & frmCajaRapida.sDetalle & ".tItem"
    '   sPedido = frmCajaRapida.Pedido
    'Else
       Set RsTitulo = Lib.OpenRecordset("select nAdulto, nNino, fregistro, tObservacion from MPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn)
       Set RsCombo = Lib.OpenRecordset("select * from vPedidoCombo where tCodigoPedido='" & sPedido & "'", Cn)
       Isql = "select * from vPedidoDetalle where tCodigoPedido = '" & sPedido & "' order by tItem"
    'End If
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
        
    Call ConfGrilla(18, grdPedidos, "Itm", 2, "tItem", 430, 2, 0, "", _
                                   "Ord", 2, "nOrden", 400, 2, 0, "", _
                                   "-", 2, "lCorte", 250, 2, 4, "", _
                                   "Producto", 2, "Producto", 1900, 0, 0, "", _
                                   "P.Oficial", 2, "nPrecioOficial", 780, 1, 0, "###,###,##0.00", _
                                   "Dcto", 2, "nDescuento", 600, 1, 0, "##0.00", _
                                   "P.Venta", 2, "nPrecioVenta", 780, 1, 0, "###,###,##0.00", _
                                   "Cant.", 2, "nCantidad", 700, 1, 0, "##,##0.00", _
                                   "SubTotal", 2, "nVenta", 900, 1, 0, "###,###,##0.00", _
                                   "( % )", 2, "Descuento", 550, 1, 0, "##0.00", _
                                   "F", 2, "tFacturado", 270, 0, 0, "", _
                                   "E", 2, "lImprime", 270, 0, 4, "", _
                                   "P", 2, "lPropiedad", 270, 0, 4, "", _
                                   "O", 2, "lObservacion", 270, 0, 4, "", _
                                   "Documento", 2, "tDocumento", 1500, 0, 0, "@@@@@@@@@@@@@@@", _
                                   "Usuario", 2, "tUsuariod", 1500, 0, 0, "", _
                                   "Envio", 2, "fEnvio", 1600, 0, 0, "", _
                                   "Comanda", 2, "tComanda", 1500, 0, 0, "")
    Set grdPedidos.DataSource = RsGrilla
    
    ' Llena cantidad de productos
    PicboxText pictxNumProductos, RsGrilla.RecordCount & " Producto(s)", ""
    Dim sumaDescuento As Double
    Dim numCantidad As Double
    sumaDescuento = 0
    
    If Not RsGrilla.EOF Then
        RsGrilla.MoveFirst
        Do While Not RsGrilla.EOF
            sumaDescuento = sumaDescuento + (val(RsGrilla!nDescuento) * val(RsGrilla!nCantidad))
            RsGrilla.MoveNext
        Loop
    End If
    
    
    
    PicboxText pictxNumDescuento, simboloSoles & Format(sumaDescuento, "##0.00"), ""

    'PicboxText pictxNumDescuento, simboloSoles & "suma de nDescuento", ""
        
    Call ConfGrilla(7, grdCombo, "Itm", 2, "tItemCombo", 450, 2, 0, "", _
                                 "Ord", 2, "nOrden", 420, 2, 0, "", _
                                 "Producto", 2, "Producto", 2000, 0, 0, "", _
                                 "Cant.", 2, "nCantidad", 700, 1, 0, "##,##0.00", _
                                 "E", 2, "lImprime", 310, 0, 4, "", _
                                 "P", 2, "lPropiedad", 310, 0, 4, "", _
                                 "O", 2, "lObservacion", 310, 0, 4, "")
    Set grdCombo.DataSource = RsCombo
    
    If lAlmacen Then
       Dim xSql As String
       Dim RsOp As Recordset
       Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
       If RsOp.RecordCount > 0 Then
           'CASO CORP.FERNANDEZ
            If BDLink = "1" Then
                xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                       "nInsumo, nGasto, nManoObra " & _
                       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                       "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & " t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                       "FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto =  t3x.tCodigoProducto " & _
                       "Where lNoDescargo = 1"
            Else
                xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                       "nInsumo, nGasto, nManoObra " & _
                       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union " & _
                       "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & " t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0 " & _
                       "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto =  t3x.tCodigoProducto " & _
                       "Where lNoDescargo = 1"
            End If
       Else
          xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                 "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                 "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                 "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
       End If
    Else
          xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                 "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                 "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                 "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
    End If
    
    ''Producto vs Propiedad
    'If sTipo = "CajaRapida" Then
    '   Isql = "SELECT " & frmCajaRapida.sProductoPropiedad & ".tItem, " & frmCajaRapida.sProductoPropiedad & ".tCodigoPropiedad, " & frmCajaRapida.sProductoPropiedad & ".tProducto, t1.Descripcion AS Descripcion, Operador , " & frmCajaRapida.sProductoPropiedad & ".ncantidad ncantidad " & _
    '          "FROM " & frmCajaRapida.sProductoPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & frmCajaRapida.sProductoPropiedad & ".tCodigoPropiedad = T1.Codigo and " & frmCajaRapida.sProductoPropiedad & ".tProducto = T1.tProducto and " & frmCajaRapida.sProductoPropiedad & ".tenlace= t1.tenlace "
    'Else
       ''Producto Propiedad
       Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador , isnull(tproductopropiedad.ncantidad,1) ncantidad " & _
              "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo and dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto and dbo.TPRODUCTOPROPIEDAD.tenlace= t1.tenlace " & _
              "where tCodigoPedido = '" & sPedido & "'"
       'txtAdulto.Caption = IIf(IsNull(RsTitulo!nAdulto), 0, RsTitulo!nAdulto)
       'txtNino.Caption = IIf(IsNull(RsTitulo!nNino), 0, RsTitulo!nNino)
       PicboxText pictxFecha, IIf(IsNull(RsTitulo!fRegistro), "", RsTitulo!fRegistro), ""
       PicboxText pictxObs, IIf(IsNull(RsTitulo!tObservacion), "", RsTitulo!tObservacion), ""
    'End If
    Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
    
    'Combo Propiedad
    'If sTipo = "CajaRapida" Then
    '   Isql = "SELECT " & frmCajaRapida.sComboPropiedad & ".tItem, " & frmCajaRapida.sComboPropiedad & ".tItemCombo, T1.Descripcion, T1.Operador ," & frmCajaRapida.sComboPropiedad & ".ncantidad  " & _
    '          "FROM " & frmCajaRapida.sComboPropiedad & " INNER JOIN (" & xSql & ") T1 ON " & frmCajaRapida.sComboPropiedad & ".tCodigoPropiedad = T1.Codigo AND " & frmCajaRapida.sComboPropiedad & ".tProducto = T1.tProducto AND dbo." & frmCajaRapida.sComboPropiedad & ".tEnlace = T1.tEnlace"
    'Else
       Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador , isnull(tcombopropiedad.ncantidad,1) ncantidad " & _
              "FROM dbo.TCOMBOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace " & _
              "WHERE tcodigoPedido='" & sPedido & "'"
    'End If
    Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
    
    
End Sub
Private Sub inicializarComponentes()
    PicboxBorder pictxNumProductos, "#CFD3D8"
    PicboxBgColor pictxNumProductos, "#FFFFFF"
    PicboxText pictxNumProductos, "0 Producto(s)", ""
    PicboxTextBold pictxNumProductos, 1
    PicboxTextSize pictxNumProductos, 12
    PicboxTextColor pictxNumProductos, "#232323"
    
    PicboxBorder pictxIcbper, "#CFD3D8"
    PicboxBgColor pictxIcbper, "#EDFAF2"
    PicboxText pictxIcbper, "ICBPER", ""
    PicboxTextBold pictxIcbper, 1
    PicboxTextSize pictxIcbper, 12
    PicboxTextColor pictxIcbper, "#38A86B"
    
    PicboxBorder pictxNumIcbper, "#CFD3D8"
    PicboxBgColor pictxNumIcbper, "#FFFFFF"
    PicboxText pictxNumIcbper, "S/. 0.00", ""
    PicboxTextBold pictxNumIcbper, 1
    PicboxTextSize pictxNumIcbper, 12
    PicboxTextColor pictxNumIcbper, "#CFD3D8"
    
    PicboxBorder pictxDescuento, "#CFD3D8"
    PicboxBgColor pictxDescuento, "#EDFAF2"
    PicboxText pictxDescuento, "Descuento", ""
    PicboxTextBold pictxDescuento, 1
    PicboxTextSize pictxDescuento, 12
    PicboxTextColor pictxDescuento, "#38A86B"
    
    PicboxBorder pictxNumDescuento, "#CFD3D8"
    PicboxBgColor pictxNumDescuento, "#FFFFFF"
    PicboxText pictxNumDescuento, "S/. 0.00", ""
    PicboxTextBold pictxNumDescuento, 1
    PicboxTextSize pictxNumDescuento, 12
    PicboxTextColor pictxNumDescuento, "#FF655A"
    
    PicboxBorder pictxTotal, "#CFD3D8"
    PicboxBgColor pictxTotal, "#EDFAF2"
    PicboxText pictxTotal, "Total", ""
    PicboxTextBold pictxTotal, 1
    PicboxTextSize pictxTotal, 12
    PicboxTextColor pictxTotal, "#38A86B"
    
    PicboxBorder pictxNumTotal, "#CFD3D8"
    PicboxBgColor pictxNumTotal, "#FFFFFF"
    PicboxText pictxNumTotal, "S/. 158.00", ""
    PicboxTextBold pictxNumTotal, 1
    PicboxTextSize pictxNumTotal, 12
    PicboxTextColor pictxNumTotal, "#232323"
End Sub

Private Sub OrganizaComponentes()
    On Error GoTo fin
    
        frmDetallePedido.Width = 14895 '14805 '15585
        frmDetallePedido.Height = 11040 '11540
        
        'PICBOX CLIENTE
        'txtCliente.Height = fraCliente.Height - 40
        'txtCliente.Width = fraCliente.Width - 40
        'txtCliente.Top = 20
        'txtCliente.Left = 20
        PicboxBorder pictxCliente, "#CFD3D8"
        PicboxBgColor pictxCliente, "#FFFFFF"
        PicboxTextSize pictxCliente, 10
        'PICBOX CLIENTE
        
        'PICBOX FECHA
        'txtFecha.Height = fraFecha.Height - 40
        'txtFecha.Width = fraFecha.Width - 40
        'txtFecha.Top = 20
        'txtFecha.Left = 20
        PicboxBorder pictxFecha, "#CFD3D8"
        PicboxBgColor pictxFecha, "#FFFFFF"
        PicboxTextSize pictxFecha, 10
        'FRAME FECHA
        
        'PICBOX OBS
        'txtObs.Height = fraObservacion.Height - 40
        'txtObs.Width = fraObservacion.Width - 40
        'txtObs.Top = 20
        'txtObs.Left = 20
        PicboxBorder pictxObs, "#CFD3D8"
        PicboxBgColor pictxObs, "#FFFFFF"
        PicboxTextSize pictxObs, 10
        'PICBOX OBS
        
        'FRAME PROPIEDAD
        txtPropiedad.Height = fraPropiedad.Height - 40
        txtPropiedad.Width = fraPropiedad.Width - 40
        txtPropiedad.Top = 20
        txtPropiedad.Left = 20
        'FRAME PROPIEDADES1
        
        'FRAME OBSERVACION
        txtObservacion.Height = fraObservacion.Height - 40
        txtObservacion.Width = fraObservacion.Width - 40
        txtObservacion.Top = 20
        txtObservacion.Left = 20
        'FRAME OBSERVACION
        
        'FRAME PROPIEDADCOMBO
        txtPropiedadCombo.Height = fraPropiedadCombo.Height - 40
        txtPropiedadCombo.Width = fraPropiedadCombo.Width - 40
        txtPropiedadCombo.Top = 20
        txtPropiedadCombo.Left = 20
        'FRAME PROPIEDADCOMBO
        
        'FRAME OBSERVACIONCOMBO
        txtObservacionCombo.Height = fraObservacionCombo.Height - 40
        txtObservacionCombo.Width = fraObservacionCombo.Width - 40
        txtObservacionCombo.Top = 20
        txtObservacionCombo.Left = 20
        'FRAME OBSERVACIONCOMBO
        
        'FRAME PEDIDOS
        grdPedidos.Height = fraPedidos.Height - 40
        grdPedidos.Width = fraPedidos.Width - vscrPedidos.Width - 40
        grdPedidos.Top = fraPedidos.Top + 20
        grdPedidos.Left = fraPedidos.Left + 20
        
        vscrPedidos.Height = fraPedidos.Height - 40
        vscrPedidos.Width = fraPedidos.Width - grdPedidos.Width - 40
        vscrPedidos.Top = fraPedidos.Top + 20
        vscrPedidos.Left = vscrPedidos.Left - 20
        'FRAME PEDIDOS
        
        'FRAME CANAL DE VENTA
        'FrameCanalVenta.Width = 6256.2
        'FrameCanalVenta.Height = 3135
        
        'FrameCanalVenta.Top = 780
        'FrameCanalVenta.Left = 9241.772
        'FRAME CANAL DE VENTA
        
    Exit Sub
fin:
    MsgBox "Error" & Error
End Sub

Private Sub grdCombo_Click()
    tmrConsultaCombo.Enabled = True
    'SubAsignar
End Sub



Private Sub picReactivarPedido_Click()
    Unload Me
    Unload frmMesas
    'frmCajaRapida.CargaPedido
End Sub

Private Sub tmrConsultaCombo_Timer()
    ' Detén el timer para que se ejecute solo una vez
    tmrConsultaCombo.Enabled = False

    SubAsignar
End Sub
Public Sub SubAsignar()
   If Not RsCombo.EOF Then
      txtPropiedadCombo.Text = ""
      txtObservacionCombo.Text = ""
      If Not RsCombo.EOF And Not RsGrilla.EOF Then
         RsComboPropiedad.Filter = "tItem='" & RsGrilla!tItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
         If Not RsComboPropiedad.EOF Then
            RsComboPropiedad.MoveFirst
            Do While Not RsComboPropiedad.EOF
               txtPropiedadCombo.Text = txtPropiedadCombo.Text & IIf(RsComboPropiedad!nCantidad = 1, "", "(" & RsComboPropiedad!nCantidad & ")") & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!descripcion) & ", "
               RsComboPropiedad.MoveNext
            Loop
         End If
      End If
      txtObservacionCombo.Text = IIf(IsNull(RsCombo!tObservacion), "", RsCombo!tObservacion)
   End If
End Sub
Private Sub grdPedidos_Click()
    tmrConsulta.Enabled = True
    'Asignar
End Sub
Private Sub tmrConsulta_Timer()
    ' Detén el timer para que se ejecute solo una vez
    tmrConsulta.Enabled = False

    Asignar
End Sub
Public Sub Asignar()
   If Not RsGrilla.EOF Then
      RsProductoPropiedad.Filter = "tItem='" & RsGrilla!tItem & "'"
      txtPropiedad.Text = ""
      If Not RsProductoPropiedad.EOF Then
         RsProductoPropiedad.MoveFirst
         Do While Not RsProductoPropiedad.EOF
            txtPropiedad.Text = txtPropiedad.Text & IIf(RsProductoPropiedad!nCantidad = 1, "", "(" & RsProductoPropiedad!nCantidad & ")") & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!descripcion) & ", "
            RsProductoPropiedad.MoveNext
         Loop
      End If
      txtObservacion.Text = IIf(IsNull(RsGrilla!tObservacion), "", RsGrilla!tObservacion)
      txtPropiedadCombo.Text = ""
      txtObservacionCombo.Text = ""
      RsCombo.Filter = "tItem='" & RsGrilla!tItem & "'"
   End If
End Sub


Private Sub lblCerrar_Click()
    Unload Me
End Sub

Private Sub picVolver_Click()
    Unload Me
End Sub
