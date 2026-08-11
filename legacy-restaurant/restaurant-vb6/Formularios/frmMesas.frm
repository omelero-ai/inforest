VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmMesas 
   BorderStyle     =   0  'None
   Caption         =   "Mesas"
   ClientHeight    =   12930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18915
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrActualiza 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   11880
   End
   Begin VB.Timer tmrConsulta 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   11880
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   10320
      Width           =   15615
      Begin VB.PictureBox picAceptar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   13200
         Picture         =   "frmMesas.frx":0000
         ScaleHeight     =   660
         ScaleWidth      =   2220
         TabIndex        =   80
         Top             =   240
         Width           =   2220
      End
      Begin VB.PictureBox picActualizar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   10800
         Picture         =   "frmMesas.frx":15D0
         ScaleHeight     =   660
         ScaleWidth      =   2220
         TabIndex        =   79
         Top             =   240
         Width           =   2220
      End
      Begin VB.PictureBox picJuntarMesas 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   8400
         Picture         =   "frmMesas.frx":2C4F
         ScaleHeight     =   660
         ScaleWidth      =   2220
         TabIndex        =   78
         Top             =   240
         Width           =   2220
      End
      Begin VB.PictureBox picVolver 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         Picture         =   "frmMesas.frx":44F2
         ScaleHeight     =   660
         ScaleWidth      =   1920
         TabIndex        =   77
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E8E5E2&
      BorderStyle     =   0  'None
      Height          =   10155
      Left            =   11850
      TabIndex        =   3
      Top             =   230
      Width           =   3735
      Begin TrueOleDBGrid80.TDBGrid grdPedidos 
         Height          =   9855
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   17383
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
         Caption         =   "Pedidos sin mesa"
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
      Begin VB.Line Line3 
         BorderColor     =   &H009E9188&
         X1              =   -120
         X2              =   3960
         Y1              =   10
         Y2              =   10
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F7F6F5&
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   11895
      Begin VB.PictureBox pictxNumAdultos 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   4800
         ScaleHeight     =   660
         ScaleWidth      =   495
         TabIndex        =   82
         Top             =   7320
         Width           =   495
      End
      Begin VB.PictureBox pictxNumNinios 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   8520
         ScaleHeight     =   660
         ScaleWidth      =   495
         TabIndex        =   81
         Top             =   7320
         Width           =   495
      End
      Begin VB.PictureBox picNumNinios 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   5880
         Picture         =   "frmMesas.frx":5DFF
         ScaleHeight     =   660
         ScaleWidth      =   2610
         TabIndex        =   76
         Top             =   7320
         Width           =   2610
      End
      Begin VB.PictureBox picNumAdultos 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   2160
         Picture         =   "frmMesas.frx":7920
         ScaleHeight     =   660
         ScaleWidth      =   2610
         TabIndex        =   75
         Top             =   7320
         Width           =   2610
      End
      Begin VB.PictureBox Picture67 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   960
         Picture         =   "frmMesas.frx":9588
         ScaleHeight     =   270
         ScaleWidth      =   9225
         TabIndex        =   74
         Top             =   6720
         Width           =   9225
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   64
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   73
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   63
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   72
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   62
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   71
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   61
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   70
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   60
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   69
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   59
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   68
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   58
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   67
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   57
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   66
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   56
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   65
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   55
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   64
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   54
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   63
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   53
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   62
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   52
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   61
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   51
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   60
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   50
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   59
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   49
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   58
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   48
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   57
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   47
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   56
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   46
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   55
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   45
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   54
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   44
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   53
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   43
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   52
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   42
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   51
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   41
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   50
         Top             =   4320
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   40
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   49
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   39
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   48
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   38
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   47
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   37
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   46
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   36
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   45
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   35
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   44
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   34
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   43
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   33
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   42
         Top             =   3600
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   32
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   41
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   31
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   40
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   30
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   39
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   29
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   38
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   28
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   37
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   27
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   36
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   26
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   35
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   25
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   34
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   24
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   33
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   23
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   32
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   22
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   21
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   30
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   20
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   29
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   19
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   28
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   18
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   27
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   17
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   26
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   16
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   25
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   15
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   14
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   23
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   13
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   12
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   11
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   10
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   9
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   8
         Left            =   10320
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   7
         Left            =   8880
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   6
         Left            =   7440
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   5
         Left            =   6000
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   4
         Left            =   4560
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   3
         Left            =   3120
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   2
         Left            =   1680
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picMesa 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblMesas 
         BackColor       =   &H00F7F6F5&
         Caption         =   "Salón 1 tiene 23/72 mesas libres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   10935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   11895
      Begin VB.PictureBox picSalon 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   7
         Left            =   9480
         ScaleHeight     =   660
         ScaleWidth      =   1455
         TabIndex        =   89
         Top             =   600
         Width           =   1455
      End
      Begin VB.PictureBox picSalon 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   6
         Left            =   7920
         ScaleHeight     =   660
         ScaleWidth      =   1455
         TabIndex        =   88
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picSalon 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   5
         Left            =   6360
         ScaleHeight     =   660
         ScaleWidth      =   1455
         TabIndex        =   87
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picSalon 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   4
         Left            =   4800
         ScaleHeight     =   660
         ScaleWidth      =   1455
         TabIndex        =   86
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picSalon 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   3
         Left            =   3240
         ScaleHeight     =   660
         ScaleWidth      =   1455
         TabIndex        =   85
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picSalon 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   2
         Left            =   1680
         ScaleHeight     =   660
         ScaleWidth      =   1455
         TabIndex        =   84
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picSalon 
         BorderStyle     =   0  'None
         Height          =   660
         Index           =   1
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   1455
         TabIndex        =   83
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox picBuscar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   11040
         Picture         =   "frmMesas.frx":CADA
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   7
         Top             =   600
         Width           =   660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009E9188&
         X1              =   -240
         X2              =   11640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Salones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label lblCerrar 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
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
      Left            =   15080
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Label txtTitulo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mesas"
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
      TabIndex        =   0
      Top             =   0
      Width           =   15465
   End
End
Attribute VB_Name = "frmMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsMesa As Recordset
Dim RsSalon As Recordset
Dim RsPedido As Recordset
Dim RsGrilla As Recordset

Dim sTitulo As String
Dim sMesaActual As String
Dim nIndex As Integer
Dim xIndex As Integer
Dim xSalon As String
Dim sMesa() As String
Dim Estado As String
Dim wClick As Boolean
Dim sSinMesa As String

Dim sSalonActual As String

Dim tcolorSeleccion As String
Dim tcolorDeseleccion As String
Private SelectedMesaIndex As Integer  ' Índice de la mesa actualmente seleccionada

Private Sub Form_Load()
    Call OrganizaComponentes
    Centrar Me
    
    tcolorSeleccion = "#40160d"
    tcolorDeseleccion = "#CFD3D8"
    
    ' Inicializamos sin ninguna mesa seleccionada
    SelectedMesaIndex = -1
    
    Dim i As Integer
    wClick = False
    sCodigo = ""
    sDescrip = ""
    sTemp = ""
    xIndex = 0
    wMesa = False
    wCambioMesa = False
    
    
    
    'Set RsPedido = Lib.OpenRecordset("select tCodigoPedido as Codigo, tMesa, fFecha, nPrecuenta from MPEDIDO where tEstadoPedido = '01' order by tMesa", Cn)
    'Set RsGrilla = Lib.OpenRecordset("select tObservacion from MPEDIDO where tEstadoPedido = '01' and tTipoPedido <> '04' and not len(rtrim(tObservacion)) = 0 and tCaja = '" & sCaja & "' order by tObservacion", Cn)
    
    'Configuración de Grillas
    'Call ConfGrilla(1, grdGrilla, "Mesa", 2, "tObservacion", 2650, 0, 0, "")
    'Set grdGrilla.DataSource = RsGrilla
    
    'Salon

    inicializarSalon 7
    inicializarMesa 64
    SalonSeleccionado 7, picSalon(1)
    'asignarMesa "01"
    Set RsPedido = Lib.OpenRecordset("select tCodigoPedido as Codigo, tMesa, fFecha, nPrecuenta from MPEDIDO where tEstadoPedido = '01' order by tMesa", Cn)
    
    'Set RsGrilla = Lib.OpenRecordset("select tCodigoPedido,tObservacion from MPEDIDO where tEstadoPedido = '01' and tTipoPedido <> '04' and not len(rtrim(tObservacion)) = 0 and tCaja = '" & sCaja & "' order by tObservacion", Cn)
    Set RsGrilla = Lib.OpenRecordset("select tCodigoPedido,tObservacion from MPEDIDO where tEstadoPedido = '01' and tTipoPedido <> '04' and len(rtrim(tMesa)) = 0 and tCaja = '" & sCaja & "' order by tCodigoPedido desc", Cn)
    
    'Configuración de Grillas
    Call ConfGrilla(2, grdPedidos, "Pedido", 2, "tCodigoPedido", 1200, 0, 0, "", _
                                   "Observación", 2, "tObservacion", 2350, 0, 0, "")
    Set grdPedidos.DataSource = RsGrilla
    
    inicializarCantidadClientes
    
End Sub
Private Sub inicializarCantidadClientes()
    PicboxBorder pictxNumAdultos, "#CFD3D8"
    PicboxText pictxNumAdultos, "0", ""
    'PicboxTextColor pictxNumAdultos, "#CFD3D8"
    
    PicboxText pictxNumNinios, "0", ""
    PicboxBorder pictxNumNinios, "#CFD3D8"
    'PicboxTextColor pictxNumNinios, "#CFD3D8"
End Sub
Private Sub SalonSeleccionado(cantidadMaxSalon As Integer, pbx As PictureBox)
    If RsSalon.RecordCount = 0 Then
        Exit Sub
    End If
    
    RsSalon.MoveFirst
    RsSalon.Move pbx.Index - 1
    
    sSalonActual = Trim(PicboxGetText(picSalon(pbx.Index), 1) & " " & PicboxGetText(picSalon(pbx.Index), 2))
    
    Dim i As Integer
    For i = 1 To cantidadMaxSalon
        If picSalon(i) Is pbx Then
            PicboxBgColor picSalon(i), "#42C67E"
        Else
            PicboxBgColor picSalon(i), "#F5F6F7"
        End If
    Next i
    asignarMesa RsSalon!tSalon
    
End Sub
Private Sub asignarMesa(codigoSalon As String)
'colores disponibles:
    'vbHexLista "01"
    'vbHexOcupada "02"
    'vbHexPrecuenta "--"
    'vbHexSucia "04"
    'vbHexMesasjuntas "--"
    'vbHexReservada "03"
    'vbHexBloqueada "05"
    'vbHexFServicio "06"
    limpiarMesas 64
    Dim countLibre As Integer
    countLibre = 0
    
    Dim i As Integer
    'RsMesa.Filter = "tSalon = '" & RsSalon!tSalon & "'"
    RsMesa.Filter = "tSalon = '" & codigoSalon & "'"
    
    If RsMesa.RecordCount = 0 Then
      Exit Sub
    End If
    
    RsMesa.MoveFirst
    For i = 1 To RsMesa.RecordCount
            
        PicboxText picMesa(i), RsMesa!tResumido, ""
        Select Case RsMesa!tEstadoMesa
            Case Is = "01" 'Lista
                PicboxBgColor picMesa(i), vbHexLista
                PicboxTextColor picMesa(i), "#232323"
                countLibre = countLibre + 1
            
            Case Is = "02" 'Ocupada
                PicboxBgColor picMesa(i), vbHexOcupada
                PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "03" 'Reservada
                PicboxBgColor picMesa(i), vbHexReservada
                PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "04" 'Sucia
                PicboxBgColor picMesa(i), vbHexSucia
                PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "05" 'Bloqueada
                PicboxBgColor picMesa(i), vbHexBloqueada
                PicboxTextColor picMesa(i), "#232323"
            
            Case Is = "06" 'Fuera de servicio
                PicboxBgColor picMesa(i), vbHexFServicio
                PicboxTextColor picMesa(i), "#98ACB3"
            
            'Case Is = "07" 'Original
        End Select
            
        picMesa(i).Visible = True
        RsMesa.MoveNext
    Next i
    
    
    'RsMesa.MoveFirst
    'Do While Not RsMesa.EOF
    '    If RsMesa!tEstadoMesa = "01" Then
    '        countLibre = countLibre + 1
    '    End If
    '    RsMesa.MoveNext
    'Loop

    
    lblMesas.Caption = sSalonActual & " " & countLibre & "/" & RsMesa.RecordCount & " mesas libres"
End Sub

Private Sub inicializarMesa(cantidadMesas As Integer)
    Set RsMesa = Lib.OpenRecordset("select * from TMESA where tCodigoMesa <> '000' and lActivo = 1 order by tCodigoMesa", Cn)
    Dim i As Integer
    'For i = 1 To RsMesa.RecordCount
    
    ' Si el recordset está vacío, salir del sub
    RsMesa.Requery
    If RsMesa.EOF And RsMesa.BOF Then Exit Sub
    RsMesa.MoveFirst
    
    For i = 1 To cantidadMesas
        PicboxText picMesa(i), "-", ""
        PicboxBgColor picMesa(i), "#FFFFFF" '#FFFFFF
        PicboxBorder picMesa(i), "#CFD3D8" '#FFFFFF
        'PicboxTextColor picMesa(i), "#FF0000"
        PicboxTextBold picMesa(i), 0
        'PicboxTextSize picMesa(2), 12
        picMesa(i).Visible = False
        If Not RsMesa.EOF Then
            RsMesa.MoveNext
        End If
    Next i
End Sub
Private Sub limpiarMesas(cantidadMaxMesas As Integer)
    Dim i As Integer
    'For i = 1 To RsMesa.RecordCount
    ' Si el recordset está vacío, salir del sub
    RsMesa.Requery
    If RsMesa.EOF And RsMesa.BOF Then Exit Sub
    RsMesa.MoveFirst
    
    For i = 1 To cantidadMaxMesas
        PicboxText picMesa(i), "-", ""
        PicboxBgColor picMesa(i), "#FFFFFF" '#FFFFFF
        PicboxBorder picMesa(i), "#CFD3D8" '#FFFFFF
        'PicboxTextColor picMesa(i), "#FF0000"
        PicboxTextBold picMesa(i), 0
        'PicboxTextSize picMesa(2), 12
        picMesa(i).Visible = False
        ' Mover al siguiente registro sólo si no es el último
        If Not RsMesa.EOF Then
            RsMesa.MoveNext
        End If
    Next i
End Sub
Private Sub inicializarSalon(cantidadMaxSalon As Integer)
    Dim i As Integer
    Set RsSalon = Lib.OpenRecordset("select tCodigo as tSalon, tDetallado as dSalon, tiCono from TTABLA where tTABLA ='Salon' and lActivo = 1 ", Cn)
    
    For i = 1 To cantidadMaxSalon
        picSalon(i).Visible = False
    Next i
    'Call MatrizBotones(IIf(RsSalon.RecordCount >= 5, 5, RsSalon.RecordCount), 2, 60, 60, cmdSalon)
    For i = 1 To RsSalon.RecordCount
        ' Botones cabecera picSalon
        ' La primera condicional sirve para mostrar el texto en 2 lineas separadas, dependiendo los caracteres disponibles
        If Len(RsSalon!dSalon) >= 13 Then ' si tiene 13 caracteres o mas, va a dividir los textos en 2 lineas
            Dim salonText1 As String
            Dim salonText2 As String
            Dim posSpace As Integer
            ' Buscar el primer espacio a partir del quinto carácter
            posSpace = InStr(5, RsSalon!dSalon, " ")
        
            If posSpace > 0 Then
                ' Si se encontró espacio, separar el texto en dos partes:
                ' la primera parte contiene desde el inicio hasta el carácter anterior al espacio
                salonText1 = Left(RsSalon!dSalon, posSpace - 1)
                ' la segunda parte contiene el resto del texto después del espacio
                salonText2 = Mid(RsSalon!dSalon, posSpace + 1)
            Else
                ' Si no se encuentra espacio, asigna el texto completo a salonText1
                salonText1 = RsSalon!dSalon
                salonText2 = ""
            End If
            
            PicboxText picSalon(i), salonText1, salonText2
        Else
            PicboxText picSalon(i), RsSalon!dSalon, ""
        End If
        PicboxBgColor picSalon(i), "#F5F6F7" '#FFFFFF
        PicboxBorder picSalon(i), "#CFD3D8"
        PicboxTextColor picSalon(i), "#232323"
        PicboxTextBold picSalon(i), 1
        PicboxTextSize picSalon(i), 8
        
        picSalon(i).Visible = True
        
        RsSalon.MoveNext
    Next i
    RsSalon.MoveFirst
End Sub

Private Sub inicializarMesas_old()
    'Cantidad niños
    PicboxBgColor pictxNumNinios, "#F0F0F0" '#FF8484 ' #F0F0F0
    PicboxBorder pictxNumNinios, "#CFD3D8"
    PicboxText pictxNumNinios, "0", ""
    
    'Cantidad Adultos
    PicboxBgColor pictxNumAdultos, "#F0F0F0" '#FF8484 ' #F0F0F0
    PicboxBorder pictxNumAdultos, "FF0000" ' #CFD3D8
    PicboxText pictxNumAdultos, "0", ""
    
    'picbox_2
    PicboxBgColor picMesa(1), "#FFFFFF" '#FFFFFF
    PicboxText picMesa(1), "Mesa 999-B5", "hace 100 min"
    PicboxTextBold picMesa(1), 1
    
    PicboxText picMesa(2), "Mesa 2", ""
    PicboxBgColor picMesa(2), "89ac76" '#FFFFFF
    PicboxBorder picMesa(2), "FF0000" '#FFFFFF
    PicboxTextColor picMesa(2), "#FF0000"
    PicboxTextBold picMesa(2), 0
    'PicboxTextSize picMesa(2), 12
    
    'RefreshPictureBox picMesa(2)
End Sub
Private Sub matrizPicboxSalon(pbx As PictureBox)

End Sub

Private Sub OrganizaComponentes()
    On Error GoTo fin
    
        frmMesas.Width = 15585
        frmMesas.Height = 11540
        
        'FRAME ELIMINACION
        'fraEliminacion.Height = 9015
        'fraEliminacion.Width = 6226.194
        'fraEliminacion.Top = 1440
        'fraEliminacion.Left = 9241.771
        'FRAME ELIMINACION
        
        'FRAME PROPIEDADES
        'fraPropiedad.Height = 10215
        'fraPropiedad.Width = 6226.194
        'fraPropiedad.Top = 240
        'fraPropiedad.Left = 9241.771
        'FRAME PROPIEDADES
        
        'FRAME MOTORIZADO
        'fraMorotizado.Height = 9015
        'fraMorotizado.Width = 6226.194
        'fraMorotizado.Top = 1440
        'fraMorotizado.Left = 9241.771
        'FRAME MOTORIZADO
        
        'FRAME MESERO
        'fraMozo.Height = 9015
        'fraMozo.Width = 6226.194
        'fraMozo.Top = 1440
        'fraMozo.Left = 9241.771
        'FRAME MESERO
        
        'FRAME ORIGEN DE VENTAS
        'fraOrigenVentas.Height = 9015
        'fraOrigenVentas.Width = 6226.194
        'fraOrigenVentas.Top = 1440
        'fraOrigenVentas.Left = 9241.771
        'FRAME ORIGEN DE VENTAS
        
        'FRAME PUNTO VENTA
        'fraPuntoVenta.Height = 9015
        'fraPuntoVenta.Width = 6226.194
        'fraPuntoVenta.Top = 1440
        'fraPuntoVenta.Left = 9241.771
        'FRAME PUNTO VENTA
        
        'FRAME PRODUCTO COMBO
        'fraProductoCombo.Height = 11325
        'fraProductoCombo.Width = 6226.193
        'fraProductoCombo.Top = 240
        'fraProductoCombo.Left = 9241.771
        'FRAME PRODUCTO COMBO
        
        'FRAME DET COMBO
        'fraCombo.Height = 9975
        'fraCombo.Width = 9256.773
        'fraCombo.Top = 1560
        'fraCombo.Left = 0
        'FRAME DET COMBO
        
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

Private Sub grdPedidos_Click()
    'frmDetallePedido.Show vbModal
    'If RsGrilla.RecordCount > 0 Then
    '  sPedido = Calcular("select tCodigoPedido as Codigo from MPEDIDO where tObservacion = '" & grdPedidos.Columns(1).value & "' and tEstadoPedido = '01'", Cn)
    '  frmPedido.Show vbModal
    'End If
    
    'Se agregó este timer ya que se demoraba al actualizar selección de la grilla, ocasionando consultar la selección anterior y no la actual
    tmrConsulta.Enabled = True
End Sub

Private Sub picAceptar_Click()
    If SelectedMesaIndex <> -1 Then
        RsMesa.MoveFirst
        RsMesa.Move SelectedMesaIndex - 1
        'SelectedMesaIndex = -1

        Dim xmesaActual As String
        Dim xnumAdultoActual As String
        Dim xnumNInioActual As String
        Dim respuesta As VbMsgBoxResult
        If IsPedidoCajaRapida Then
            'sPedido
            respuesta = MsgBox("¿Desea asignar una mesa?", vbQuestion + vbYesNo, "Mesa: " & RsMesa!tCodigoMesa & " - " & RsMesa!tResumido)
            If respuesta = vbYes Then
                xmesaActual = IIf(IsNull(RsMesa!tCodigoMesa), "", RsMesa!tCodigoMesa)
                xnumAdultoActual = PicboxGetText(pictxNumAdultos, 1)
                xnumNInioActual = PicboxGetText(pictxNumNinios, 1)
                'frmCajaRapida.AsignaMesaCajaRapida xmesaActual, xnumAdultoActual, xnumNInioActual
                Unload Me
            End If
        Else
            respuesta = MsgBox("¿Desea crear un pedido con esta mesa?", vbQuestion + vbYesNo, "Mesa: " & RsMesa!tCodigoMesa & " - " & RsMesa!tResumido)
            If respuesta = vbYes Then
                xmesaActual = IIf(IsNull(RsMesa!tCodigoMesa), "", RsMesa!tCodigoMesa)
                xnumAdultoActual = PicboxGetText(pictxNumAdultos, 1)
                xnumNInioActual = PicboxGetText(pictxNumNinios, 1)
                'frmCajaRapida.CrearPedidoMesa xmesaActual, xnumAdultoActual, xnumNInioActual
                Unload Me
            End If
        End If
        
    Else
        MsgBox "No ha seleccionado una mesa lista para crear un pedido.", vbExclamation, "Aviso"
    End If
End Sub

Private Sub picActualizar_Click()
    inicializarMesa 64
    Set grdPedidos.DataSource = Nothing
    
    tmrActualiza.Enabled = True
End Sub

Private Sub tmrActualiza_Timer()
' Este metodo es usado para dar un tiempo de espera y dar la sensacion que si se actualizó (es decir, es para uso visual)
    ' Detén el timer para que se ejecute solo una vez
    tmrActualiza.Enabled = False
    
    LimpiarSeleccion
    inicializarCantidadClientes
    
    asignarMesa RsSalon!tSalon
    
    'Set RsGrilla = Lib.OpenRecordset("select tCodigoPedido,tObservacion from MPEDIDO where tEstadoPedido = '01' and tTipoPedido <> '04' and not len(rtrim(tObservacion)) = 0 and tCaja = '" & sCaja & "' order by tObservacion", Cn)
    Set RsGrilla = Lib.OpenRecordset("select tCodigoPedido,tObservacion from MPEDIDO where tEstadoPedido = '01' and tTipoPedido <> '04' and len(rtrim(tMesa)) = 0 and tCaja = '" & sCaja & "' order by tCodigoPedido desc", Cn)
    'Configuración de Grillas
    Call ConfGrilla(2, grdPedidos, "Pedido", 2, "tCodigoPedido", 1200, 0, 0, "", _
                                   "Observación", 2, "tObservacion", 2350, 0, 0, "")
    Set grdPedidos.DataSource = RsGrilla
End Sub

Private Sub picJuntarMesas_Click()
    'frmDetallePedido.Show vbModal
    sSalon = RsSalon!tSalon
    'frmJuntarMesas.parametrosIni "uhm"
    'frmJuntarMesas.Show vbModal
    
    If SelectedMesaIndex <> -1 Then
        RsMesa.MoveFirst
        RsMesa.Move SelectedMesaIndex - 1
        'SelectedMesaIndex = -1

        Dim xmesaActual As String
        Dim xnumAdultoActual As String
        Dim xnumNInioActual As String
        Dim respuesta As VbMsgBoxResult
        'If IsPedidoCajaRapida Then
            ''sPedido
            'respuesta = MsgBox("¿Desea asignar una mesa?", vbQuestion + vbYesNo, "Mesa: " & RsMesa!tCodigoMesa & " - " & RsMesa!tResumido)
            'If respuesta = vbYes Then
            '    xmesaActual = IIf(IsNull(RsMesa!tCodigoMesa), "", RsMesa!tCodigoMesa)
            '    xnumAdultoActual = PicboxGetText(pictxNumAdultos, 1)
            '    xnumNInioActual = PicboxGetText(pictxNumNinios, 1)
            '    frmCajaRapida.AsignaMesaCajaRapida xmesaActual, xnumAdultoActual, xnumNInioActual
            '    Unload Me
            'End If
        'Else
            'respuesta = MsgBox("¿Desea crear un pedido con esta mesa?", vbQuestion + vbYesNo, "Mesa: " & RsMesa!tCodigoMesa & " - " & RsMesa!tResumido)
            'If respuesta = vbYes Then
                xmesaActual = IIf(IsNull(RsMesa!tCodigoMesa), "", RsMesa!tCodigoMesa)
                xnumAdultoActual = PicboxGetText(pictxNumAdultos, 1)
                xnumNInioActual = PicboxGetText(pictxNumNinios, 1)
                
        frmJuntarMesas.parametrosIni xmesaActual, xnumAdultoActual, xnumNInioActual
        frmJuntarMesas.Show vbModal
                'frmCajaRapida.CrearPedidoMesa xmesaActual, xnumAdultoActual, xnumNInioActual
                'Unload Me
            'End If
        'End If
        
    Else
        MsgBox "No ha seleccionado una mesa lista.", vbExclamation, "Aviso"
    End If
    
End Sub

Private Sub picMesa_Click(Index As Integer)
'colores disponibles:
    'vbHexLista "01"
    'vbHexOcupada "02"
    'vbHexPrecuenta "--"
    'vbHexSucia "04"
    'vbHexMesasjuntas "--"
    'vbHexReservada "03"
    'vbHexBloqueada "05"
    'vbHexFServicio "06"
    
    Dim actualColorHex As String
    'Dim estadoMesa     As String
    
    
    RsMesa.MoveFirst
    RsMesa.Move Index - 1
    
    'estadoMesa = Trim(RsMesa!tEstadoMesa)
    actualColorHex = PicboxGetBorderHex(picMesa(Index))
    
    'PicboxText picMesa(i), RsMesa!tResumido, ""
        Select Case RsMesa!tEstadoMesa
            Case Is = "01" 'Lista
                'Logica de seleccion
                ' 1) Si ya estaba seleccionada y hacen clic de nuevo, la deseleccionamos
                If SelectedMesaIndex = Index Then
                    PicboxBorder picMesa(Index), tcolorDeseleccion
                    SelectedMesaIndex = -1
                    
                Else
                    ' 2) Si había otra mesa seleccionada, la deseleccionamos
                    If SelectedMesaIndex <> -1 Then
                        PicboxBorder picMesa(SelectedMesaIndex), tcolorDeseleccion
                    End If
                    
                    ' 3) Seleccionamos la nueva mesa
                    PicboxBorder picMesa(Index), tcolorSeleccion
                    SelectedMesaIndex = Index
                    
                End If
            
            Case Is = "02" 'Ocupada
                sPedido = Calcular("select tCodigoPedido as Codigo from MPEDIDO where tMesa = '" & RsMesa!tCodigoMesa & "' and tEstadoPedido = '01'", Cn)
                frmDetallePedido.Show vbModal
                'End If
            
            Case Is = "03" 'Reservada
                'No hace nada
            
            Case Is = "04" 'Sucia
                'Logica de seleccion
                ' 1) Si ya estaba seleccionada y hacen clic de nuevo, la deseleccionamos
                If SelectedMesaIndex = Index Then
                    PicboxBorder picMesa(Index), tcolorDeseleccion
                    SelectedMesaIndex = -1
                    
                Else
                    ' 2) Si había otra mesa seleccionada, la deseleccionamos
                    If SelectedMesaIndex <> -1 Then
                        PicboxBorder picMesa(SelectedMesaIndex), tcolorDeseleccion
                    End If
                    
                    ' 3) Seleccionamos la nueva mesa
                    PicboxBorder picMesa(Index), tcolorSeleccion
                    SelectedMesaIndex = Index
                    
                End If
            
            Case Is = "05" 'Bloqueada
                'No hace nada
            
            Case Is = "06" 'Fuera de servicio
                'No hace nada
        End Select
End Sub

Private Sub picNumAdultos_Click()
    Dim sAdultoTexto As String
    sAdultoTexto = PicboxGetText(pictxNumAdultos, 1)
    frmNumPad.Show vbModal
    PicboxText pictxNumAdultos, IIf(wEnter = True, Format(sDescrip, "#,##0"), sAdultoTexto), ""
End Sub

Private Sub picNumNinios_Click()
    Dim sNinioTexto As String
    sNinioTexto = PicboxGetText(pictxNumNinios, 1)
    frmNumPad.Show vbModal
    PicboxText pictxNumNinios, IIf(wEnter = True, Format(sDescrip, "#,##0"), sNinioTexto), ""
End Sub

Private Sub picVolver_Click()
    Unload Me
End Sub
Private Sub tmrConsulta_Timer()
    ' Detén el timer para que se ejecute solo una vez
    tmrConsulta.Enabled = False

    If RsGrilla.RecordCount > 0 Then
        'sPedido = Calcular("select tCodigoPedido as Codigo from MPEDIDO where tObservacion = '" & grdPedidos.Columns(1).value & "' and tEstadoPedido = '01'", Cn)
        sPedido = Calcular("select tCodigoPedido as Codigo from MPEDIDO where tCodigoPedido = '" & grdPedidos.Columns(0).value & "' and tEstadoPedido = '01'", Cn)
        'sTipo = "CajaRapida"
        'frmPedido.Show vbModal
        frmDetallePedido.Show vbModal
    End If
End Sub
Private Sub lblCerrar_Click()
    Unload Me
End Sub

Private Sub picBuscar_Click()
Dim xTipo As String
    xTipo = sTipo
    If sTipo = "M" And wClick Then
      wClick = False
    End If

    sTemp = ""
    Isql = "select * from vSalon where lActivo = 1"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    
    If wEnter Then
       RsSalon.MoveFirst
       RsSalon.Find "tSalon='" & sCodigo & "'"
       'sTemp = grdPedido.
       'fraMesa.Caption = " " & RsSalon!dSalon & " "
       'InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
       'Desasigna (RsMesa.RecordCount)
       'AsignaMesa (sTipo)
       SalonSeleccionado 7, picSalon(RsSalon.AbsolutePosition)
    End If
    sTipo = xTipo
End Sub

Private Sub picSalon_Click(Index As Integer)
    SalonSeleccionado 7, picSalon(Index)
    LimpiarSeleccion
End Sub


Private Sub LimpiarSeleccion()
    ' 1) Deselecciona la mesa activa, si existe
    If SelectedMesaIndex <> -1 Then
        PicboxBorder picMesa(SelectedMesaIndex), tcolorDeseleccion
        SelectedMesaIndex = -1
    End If
    
    ' 2) Reinicia los contadores de clientes
    inicializarCantidadClientes
End Sub


