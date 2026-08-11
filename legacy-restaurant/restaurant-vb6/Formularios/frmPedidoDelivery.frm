VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPedidoDelivery 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   7620
   ClientLeft      =   -3525
   ClientTop       =   -1125
   ClientWidth     =   11880
   Icon            =   "frmPedidoDelivery.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7594.341
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrilla 
      Height          =   4575
      Left            =   0
      TabIndex        =   9
      Top             =   -45
      Width           =   11835
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   4260
         Left            =   45
         TabIndex        =   15
         Top             =   180
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   7514
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
         Caption         =   "Correlativo"
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
   Begin VB.Frame Frame 
      Caption         =   " Datos Adicionales "
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
      Height          =   2310
      Left            =   0
      TabIndex        =   16
      Top             =   4590
      Width           =   11835
      Begin VB.Frame Frame1 
         Caption         =   " Otros Datos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   90
         TabIndex        =   17
         Top             =   270
         Width           =   5100
         Begin VB.Label txtDireccion 
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
            ForeColor       =   &H00C00000&
            Height          =   465
            Left            =   990
            TabIndex        =   23
            Top             =   315
            Width           =   4020
         End
         Begin VB.Label txtZona 
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
            ForeColor       =   &H00C00000&
            Height          =   465
            Left            =   990
            TabIndex        =   22
            Top             =   810
            Width           =   4020
         End
         Begin VB.Label txtReferencia 
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
            ForeColor       =   &H00C00000&
            Height          =   465
            Left            =   990
            TabIndex        =   21
            Top             =   1305
            Width           =   4020
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Dirección :"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   20
            Top             =   315
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Referencia :"
            Height          =   195
            Index           =   3
            Left            =   45
            TabIndex        =   19
            Top             =   1305
            Width           =   870
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Zona :"
            Height          =   195
            Index           =   2
            Left            =   465
            TabIndex        =   18
            Top             =   810
            Width           =   465
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetalle 
         Height          =   1860
         Left            =   5220
         TabIndex        =   24
         Top             =   315
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   3281
         _LayoutType     =   4
         _RowHeight      =   23
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
         Caption         =   "Detalle del Pedido"
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
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   6915
      Width           =   11880
      Begin VB.Frame fraCriterio 
         Caption         =   " Criterio "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5940
         TabIndex        =   12
         Top             =   0
         Width           =   2190
         Begin VB.TextBox txtResultado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraCampo 
         Caption         =   " Campo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4095
         TabIndex        =   10
         Top             =   0
         Width           =   1830
         Begin VB.ComboBox cboCriterio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   75
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Mostrar Datos"
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
         Left            =   8190
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   1200
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "KeyBoard"
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
         Index           =   7
         Left            =   9390
         Picture         =   "frmPedidoDelivery.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Width           =   1200
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
         Left            =   10590
         Picture         =   "frmPedidoDelivery.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   1200
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   3990
         TabIndex        =   1
         Top             =   45
         Width           =   4050
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   495
            Picture         =   "frmPedidoDelivery.frx":0636
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmPedidoDelivery.frx":0B78
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   3510
            Picture         =   "frmPedidoDelivery.frx":10BA
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3015
            Picture         =   "frmPedidoDelivery.frx":15FC
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   495
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1035
            TabIndex        =   14
            Top             =   150
            Width           =   1920
         End
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   7905
      Top             =   6990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPedidoDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim RsDetalle As Recordset
Dim nColumna As Integer
Dim sCriterio As String

Sub LlenaBusqueda()
    Dim i As Integer
    With cboCriterio
        For i = 0 To grdGrilla.Columns.Count - 1
            If grdGrilla.Columns(i).ValueItems.Presentation <> 4 Then
                .AddItem grdGrilla.Columns(i).Caption
                .ItemData(.NewIndex) = i
            End If
        Next i
    End With
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

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 2  'Mostrar
                If cmdOpcion(2).Caption = "Mostrar Datos" Then
                   cmdOpcion(2).Caption = "No Mostrar Datos"
                   'fraGrilla.Height = IIf(Me.Height - 3790 > 0, Me.Height - 3790, 0)
                   fraGrilla.Height = IIf(Me.Height - 3500 > 0, Me.Height - 3500, 0)
                   grdGrilla.Height = IIf(fraGrilla.Height - 250 > 0, fraGrilla.Height - 250, 0)
                Else
                   cmdOpcion(2).Caption = "Mostrar Datos"
                   fraGrilla.Height = IIf(Me.Height - 1150 > 0, Me.Height - 1150, 0)
                   grdGrilla.Height = IIf(fraGrilla.Height - 250 > 0, fraGrilla.Height - 250, 0)
                End If
                
           Case Is = 3  'Filtrar
           
           Case Is = 4  'No Filtrar
           
           Case Is = 6  'Salir
                Unload Me
           
           Case Is = 7  'Keyboard
                frmKeyBoard.Show vbModal
                txtResultado.Text = IIf(wEnter, sDescrip, "")
                txtResultado.SelStart = Len(txtResultado)
           
           Case Is = 8  'Emite / Precuenta
                
           Case Is = 9  'Impresion de Pedidos
                      
           Case Is = 10 'Cambio de Propina
           
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
 
    ' Ingrese el Titulo
    Me.Caption = " Seguimientos de Pedidos Delivery "
    grdGrilla.Caption = Me.Caption
    nColumna = 0
    Centrar Me
           
    ' Ingrese la Vista
    Isql = "select * from vDespachador where tTipoPedido='02' and tEstadoPedido='02' and isnull(fLlegada,0)=0"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Configuración de la Grilla
    Call ConfGrilla(11, grdGrilla, "Pedido", 2, "tCodigoPedido", 1000, 2, 0, "", _
                                   "Fecha", 2, "fFecha", 950, 0, 0, "dd/MM HH:mm", _
                                   "Usuario", 2, "tUsuario", 1000, 0, 0, "", _
                                   "Caja", 2, "tCaja", 500, 2, 0, "", _
                                   "Teléfono", 2, "tTelefono", 900, 0, 0, "", _
                                   "Cliente", 2, "Cliente", 1800, 0, 0, "", _
                                   "Direccion", 2, "tDireccion", 1450, 0, 0, "", _
                                   "Motorizado", 2, "Motorizado", 1400, 0, 0, "", _
                                   "Asigna", 2, "fAsignacion", 700, 1, 0, "HH:mm", _
                                   "Salida", 2, "fSalida", 700, 1, 0, "HH:mm", _
                                   "T.Rest", 2, "Restante", 700, 1, 0, "")
    Set grdGrilla.DataSource = RsCabecera
    LlenaBusqueda
    Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
    Call ConfGrilla(5, grdDetalle, "Item", 2, "tItem", 600, 2, 0, "", _
                                   "Producto", 2, "Producto", 2300, 0, 0, "", _
                                   "P.Venta", 2, "nPrecioVenta", 1000, 1, 0, "###,###,##0.00", _
                                   "Cantidad", 2, "nCantidad", 950, 1, 0, "##,##0.00", _
                                   "SubTotal", 2, "nVenta", 1050, 1, 0, "###,###,##0.00")
    Asignar
    
    txtResultado.Text = ""
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
    cboCriterio.Text = "Teléfono"
End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 1150 > 0, Me.Height - 1150, 0)
   grdGrilla.Height = IIf(fraGrilla.Height - 250 > 0, fraGrilla.Height - 250, 0)
   Frame.Top = IIf(fraGrilla.Height - 2350 > 0, fraGrilla.Height - 2350, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsCabecera = Nothing
   Set frmPedidoDelivery = Nothing
End Sub

Private Sub grdGrilla_DblClick()
   cmdOpcion_Click (1)
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
   nColumna = ColIndex
   RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   txtResultado.SetFocus
   
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   If RsCabecera.RecordCount > 0 Then
      Asignar
   End If
End Sub

Sub Asignar()
    'Asigna la Cabecera
    If RsCabecera.RecordCount > 0 Then
       sPedido = RsCabecera!tCodigoPedido
       Isql = "select tItem, Producto, nPrecioVenta, nCantidad, nVenta from vPedidoDetalle where tCodigoPedido ='" & sPedido & "'"
       Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
       Set grdDetalle.DataSource = RsDetalle
       txtReferencia.Caption = IIf(IsNull(RsCabecera!Referencia), "", RsCabecera!Referencia)
       txtDireccion.Caption = IIf(IsNull(RsCabecera!tDireccion), "", RsCabecera!tDireccion)
       txtZona.Caption = IIf(IsNull(RsCabecera!Zona), "", RsCabecera!Zona)
       
       
       
       
       
 '      Dim RsDireccion As Recordset
               
'       Set RsDireccion = Lib.OpenRecordset("select * from vDireccion where Codigo='" & RsCabecera!tDireccion & "'", Cn)
'       If RsDireccion.RecordCount > 0 Then
'          txtDireccion.Caption = IIf(IsNull(RsDireccion!TipoUbicacion), "", RsDireccion!TipoUbicacion) & " " & _
'                                 IIf(IsNull(RsDireccion!Descripcion), "", RsDireccion!Descripcion) & " " & _
'                                 IIf(IsNull(RsCabecera!tNumero), "", RsCabecera!tNumero)
'
'           txtZona.Caption = IIf(IsNull(RsDireccion!TipoZona), "", RsDireccion!TipoZona) & " " & _
'                             IIf(IsNull(RsDireccion!Zona), "", RsDireccion!Zona)
'
'           txtDistancia.Caption = Format(IIf(IsNull(RsDireccion!ndistancia), 0, RsDireccion!ndistancia), "###,##0.00")
'           txtTiempo.Caption = Format(IIf(IsNull(RsDireccion!nTiempo), "", RsDireccion!nTiempo), "###,##0.00")
'       Else
'           txtDireccion.Caption = ""
'           txtZona.Caption = ""
'           txtDistancia.Caption = "0.00"
'           txtTiempo.Caption = "0.00"
'       End If
       
  '     Set RsDireccion = Nothing
   End If
End Sub

Public Sub Filtrar()
   If Len(Trim(txtResultado.Text)) > 0 Then
   
      Select Case VarType(grdGrilla.Columns((grdGrilla.Columns(cboCriterio.Text).DataField)).Value)
             Case 1
                  sCriterio = ""
             Case 2 To 6
             sCriterio = Trim(grdGrilla.Columns(cboCriterio.Text).DataField) & "= " & Val(txtResultado.Text)
             Case 7
                  sCriterio = ""
             'sCriterio = Trim(grdGrilla.Columns(cboCriterio.Text).DataField) & " >= #" & txtResultado.Text & "# and " & Trim(grdGrilla.Columns(cboCriterio.Text).DataField) & " <= #" & txtResultado.Text & " 23:59#"
             Case Else
             sCriterio = Trim(grdGrilla.Columns(cboCriterio.Text).DataField) & " Like " & "'*" & txtResultado.Text & "*'"
      End Select
      
      Screen.MousePointer = vbHourglass
      With RsCabecera
           .Filter = sCriterio
           If .BOF = False Then
              .MoveFirst
           End If
      End With
      Screen.MousePointer = vbDefault
   Else
      Screen.MousePointer = vbHourglass
      RsCabecera.Filter = adFilterNone
      RsCabecera.Requery
      If RsCabecera.RecordCount <> 0 Then
         RsCabecera.MoveFirst
      End If
      Screen.MousePointer = vbDefault
      sCriterio = ""
   End If
End Sub

Private Sub txtresultado_Change()
   Filtrar
End Sub
