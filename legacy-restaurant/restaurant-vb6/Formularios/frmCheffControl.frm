VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCheffControl 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10830
   ClientLeft      =   2520
   ClientTop       =   1710
   ClientWidth     =   15240
   Icon            =   "frmCheffControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrilla 
      Height          =   10065
      Left            =   0
      TabIndex        =   14
      Top             =   -15
      Width           =   15225
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Atendido"
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
         Left            =   7020
         Picture         =   "frmCheffControl.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   9315
         Width           =   1305
      End
      Begin VB.Timer TimerMensaje 
         Interval        =   60000
         Left            =   7515
         Top             =   9180
      End
      Begin VB.Frame Frame2 
         Caption         =   " Detalle del Pedido  "
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
         Height          =   5955
         Left            =   8415
         TabIndex        =   23
         Top             =   135
         Width           =   6720
         Begin TrueOleDBGrid80.TDBGrid grdDetalle 
            Height          =   4410
            Left            =   90
            TabIndex        =   24
            Top             =   225
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   7779
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
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Observación :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3375
            TabIndex        =   28
            Top             =   4725
            Width           =   1200
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Propiedad :"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   27
            Top             =   4725
            Width           =   990
         End
         Begin VB.Label txtObservacion 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   915
            Left            =   3375
            TabIndex        =   26
            Top             =   4950
            Width           =   3210
         End
         Begin VB.Label txtPropiedad 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000C0&
            Height          =   915
            Left            =   90
            TabIndex        =   25
            Top             =   4950
            Width           =   3210
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Combos "
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
         Height          =   3795
         Left            =   8415
         TabIndex        =   17
         Top             =   6165
         Width           =   6720
         Begin TrueOleDBGrid80.TDBGrid grdCombo 
            Height          =   2235
            Left            =   90
            TabIndex        =   18
            Top             =   270
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   3942
            _LayoutType     =   4
            _RowHeight      =   21
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
         Begin VB.Label txtPropiedadCombo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000C0&
            Height          =   915
            Left            =   3405
            TabIndex        =   22
            Top             =   2790
            Width           =   3210
         End
         Begin VB.Label txtObservacionCombo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   915
            Left            =   90
            TabIndex        =   21
            Top             =   2790
            Width           =   3210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Propiedad :"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   20
            Top             =   2565
            Width           =   990
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Observación :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   3405
            TabIndex        =   19
            Top             =   2565
            Width           =   1200
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdCabecera 
         Height          =   8955
         Left            =   45
         TabIndex        =   16
         Top             =   225
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   15796
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
      Begin VB.Label Mensaje 
         Alignment       =   2  'Center
         Caption         =   "Mensaje de ejemplo, mensaje de ejemplo, mensaje de ejemplo mensaje de ejemplo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   135
         TabIndex        =   33
         Top             =   9225
         Width           =   6765
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   15180
      TabIndex        =   1
      Top             =   10080
      Width           =   15240
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Recuperar Pedido"
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
         Left            =   12600
         Picture         =   "frmCheffControl.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   45
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
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
         Height          =   615
         Index           =   5
         Left            =   10080
         Picture         =   "frmCheffControl.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   45
         Width           =   1260
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Reimpresion"
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
         Left            =   11340
         Picture         =   "frmCheffControl.frx":0610
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   45
         Width           =   1260
      End
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
         Left            =   7065
         TabIndex        =   13
         Top             =   0
         Width           =   1710
         Begin VB.TextBox txtCriterio 
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
            TabIndex        =   0
            Top             =   240
            Width           =   1605
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
         Left            =   5220
         TabIndex        =   11
         Top             =   0
         Width           =   1845
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
            Left            =   45
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   1740
         End
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
         Left            =   13860
         Picture         =   "frmCheffControl.frx":0B42
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   1260
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
         Left            =   8820
         Picture         =   "frmCheffControl.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   1260
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   -30
         ScaleHeight     =   555
         ScaleWidth      =   5145
         TabIndex        =   2
         Top             =   30
         Width           =   5205
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmCheffControl.frx":0D36
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmCheffControl.frx":1278
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
            Picture         =   "frmCheffControl.frx":17BA
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4575
            Picture         =   "frmCheffControl.frx":1CFC
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4035
            Picture         =   "frmCheffControl.frx":223E
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3495
            Picture         =   "frmCheffControl.frx":2780
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
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
            Height          =   315
            Left            =   1695
            TabIndex        =   15
            Top             =   120
            Width           =   1725
         End
      End
   End
End
Attribute VB_Name = "frmCheffControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCriterio As String
Dim nColumna As Integer

Dim RsCabecera As Recordset
Dim RsDetalle As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsCombo As Recordset
Dim RsComboPropiedad As Recordset
Dim RsArea As Recordset
Dim RsImpresion As Recordset
Dim RsParametro As Recordset
Dim sFiltro As String
Dim i As Integer
Dim nDias As Integer

Sub LlenaBusqueda()
    Dim i As Integer
    With cboCriterio
        For i = 0 To grdCabecera.Columns.Count - 1
            If grdCabecera.Columns(i).ValueItems.Presentation <> dbgCheckBox Then
                .AddItem grdCabecera.Columns(i).Caption
                .ItemData(.NewIndex) = i
            End If
        Next i
    End With
End Sub

Private Sub cboCriterio_Click()
   txtCriterio.Text = ""
   txtCriterio.SetFocus
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdCabecera
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdCabecera
           Case Is = 2 'Previo
                MoverPuntero previo, grdCabecera
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdCabecera
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdCabecera
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdCabecera
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0  'Aceptar
               If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
            
                Screen.MousePointer = vbHourglass
                Cn.Execute "update MPEDIDO set fAtencion=getdate() where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                'Cn.Execute "update MPEDIDO set tEstadoPedido='02', fAtencion=getdate() where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                RsCabecera.Requery
                If Not RsCabecera.EOF Then
                   RsCabecera.MoveLast
                End If
                Screen.MousePointer = vbDefault
                txtCriterio.SetFocus
           
           Case Is = 1  'Recuperar
                Isql = "SELECT dbo.MPEDIDO.tCodigoPedido as Codigo, dbo.MPEDIDO.fRegistro, dbo.TCAJA.tDescripcion AS Caja, dbo.vSalon.Descripcion AS Salon, dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tObservacion as Descripcion, dbo.MPEDIDO.tUsuario, dbo.MPEDIDO.tSalon " & _
                       "FROM dbo.MPEDIDO INNER JOIN dbo.vSalon ON dbo.MPEDIDO.tSalon = dbo.vSalon.Codigo INNER JOIN dbo.TCAJA ON dbo.MPEDIDO.tCaja = dbo.TCAJA.tCaja LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa " & _
                       "where fAtencion>='" & Format(Date, "mm/dd/yyyy") & "'" & _
                       "order by tCodigoPedido"
                
                'Configuración de la Grilla
                Call ConfGrilla(7, frmBusquedaRapida.grdGrilla, "Pedido", 2, "Codigo", 1050, 0, 0, "", _
                                                                 "Fecha", 2, "fRegistro", 1000, 0, 0, "dd/MM HH:mm", _
                                                                 "Caja", 2, "Caja", 1050, 0, 0, "", _
                                                                 "Salon", 2, "Salon", 1300, 0, 0, "", _
                                                                 "Mesa", 2, "Mesa", 1300, 0, 0, "", _
                                                                 "Observacion", 2, "Observacion", 1300, 0, 0, "", _
                                                                 "Usuario", 2, "tUsuario", 1200, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 0
                frmBusquedaRapida.Show vbModal
               
                If wEnter Then
                   Screen.MousePointer = vbHourglass
                   Cn.Execute "update MPEDIDO set fAtencion=null where tCodigoPedido='" & sCodigo & "'"
                   RsCabecera.Requery
                   If Not RsCabecera.EOF Then
                      RsCabecera.Find "tCodigoPedido='" & sCodigo & "'"
                   End If
                   txtCriterio.SetFocus
                   Screen.MousePointer = vbDefault
                End If
           
           Case Is = 2  'Reimprimir
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
            
                If RsArea.RecordCount = 0 Then
                   MsgBox "No existe area definida", vbInformation, sMensaje
                   Exit Sub
                End If
                 
                Screen.MousePointer = vbHourglass
                
                If lOrden Then
                   Isql = "select * from vPedido " & _
                          "Where Codigo = '" & sPedido & "' and nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & RsCabecera!tCodigoPedido & "' and lImprime = 1 ) " & _
                          "ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                Else
                   Isql = "select * from vPedido " & _
                          "Where Codigo = '" & RsCabecera!tCodigoPedido & "' And lImprimeArea = 1 and (lImprime = 1 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 1 ))" & _
                          "ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                End If
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                 
                If Not RsImpresion.EOF Then
                   RsArea.MoveFirst
                   For i = 1 To RsArea.RecordCount
                       RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                       If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                          If RsImpresion.RecordCount <> 0 Then
                             RsImpresion.MoveFirst
                             ImprimePedido RsImpresion, "N", RsArea!tImpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad
                          End If
                       End If
                       RsArea.MoveNext
                   Next i
                End If
                LimpiaRs
           
           Case Is = 5  'Refrescar
                Screen.MousePointer = vbHourglass
                RsCabecera.Requery
                If Not RsCabecera.EOF Then
                   RsCabecera.MoveLast
                End If
                Screen.MousePointer = vbDefault
                txtCriterio.SetFocus
                Mensaje.Caption = ""
                
           Case Is = 6  'Salir
                Unload Me
           
           Case Is = 7  'Keyboard
                frmKeyBoard.Show vbModal
                txtCriterio.Text = IIf(wEnter, sDescrip, "")
    End Select
End Sub

Private Sub Form_Activate()
   cboCriterio.ListIndex = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
         
    'Ingrese el Titulo
    Me.Caption = " Cheff Control "
    Mensaje.Caption = ""
    grdCabecera.Caption = Me.Caption
    Centrar Me
    nColumna = 2
    Set RsParametro = Lib.OpenRecordset("select lCheffControlCanal1, lCheffControlCanal2, lCheffControlCanal3, lCheffControlCanal4, lCheffControlCanal5, tCheffControlEstadoPedido, nDiasCheffControl from TPARAMETRO", Cn)
    If RsParametro!lCheffControlCanal1 = True Then
       sFiltro = "'01'"
    End If
    If RsParametro!lCheffControlCanal2 = True Then
       sFiltro = sFiltro & IIf(sFiltro = "", "'02'", ",'02'")
    End If
    If RsParametro!lCheffControlCanal3 = True Then
       sFiltro = sFiltro & IIf(sFiltro = "", "'03'", ",'03'")
    End If
    If RsParametro!lCheffControlCanal4 = True Then
       sFiltro = sFiltro & IIf(sFiltro = "", "'04'", ",'04'")
    End If
    If RsParametro!lCheffControlCanal5 = True Then
       sFiltro = sFiltro & IIf(sFiltro = "", "'05'", ",'05'")
    End If
    
    If sFiltro = "" Then
       MsgBox "Error fatal: Se debe tener al menos un Canal activado", vbCritical, sMensaje
       sFiltro = "'00'"
    End If
    
    nDias = RsParametro!nDiasCheffControl
    
    'Cabecera
    Isql = "SELECT dbo.MPEDIDO.tCodigoPedido, dbo.MPEDIDO.fRegistro, dbo.TCAJA.tDescripcion AS Caja, dbo.vSalon.Descripcion AS Salon, dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tObservacion, dbo.MPEDIDO.tUsuario, dbo.MPEDIDO.tSalon " & _
           "FROM dbo.MPEDIDO INNER JOIN dbo.vSalon ON dbo.MPEDIDO.tSalon = dbo.vSalon.Codigo INNER JOIN dbo.TCAJA ON dbo.MPEDIDO.tCaja = dbo.TCAJA.tCaja LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa " & _
           "where dbo.MPEDIDO.tTipoPedido in (" & sFiltro & ") and tEstadoPedido='" & RsParametro!tCheffControlEstadoPedido & "' and isnull(fAtencion,0)=0 and MPEDIDO.fregistro >= getdate() - " & nDias & " order by tCodigoPedido"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Detalle
    Isql = "select * from vPedidoDetalle where tCodigoPedido in (SELECT dbo.MPEDIDO.tCodigoPedido From dbo.MPEDIDO where dbo.MPEDIDO.tTipoPedido in (" & sFiltro & ") and tEstadoPedido='" & RsParametro!tCheffControlEstadoPedido & "' and isnull(fAtencion,0)=0 and dbo.MPEDIDO.fregistro >= getdate() - " & nDias & ") order by tItem"
    Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
    Set grdDetalle.DataSource = RsDetalle
          
    'Combos
    Isql = "SELECT * from vPedidoCombo " & _
           "where tCodigoPedido in (SELECT dbo.MPEDIDO.tCodigoPedido From dbo.MPEDIDO where dbo.MPEDIDO.tTipoPedido in (" & sFiltro & ") and tEstadoPedido='" & RsParametro!tCheffControlEstadoPedido & "' and isnull(fAtencion,0)=0 and MPEDIDO.fregistro >= getdate() - " & nDias & ") " & _
           "ORDER BY tCodigoPedido, tItemCombo"
    Set RsCombo = Lib.OpenRecordset(Isql, Cn)

    'Areas
    Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)

    'Configuración de la Grilla
    Call ConfGrilla(7, grdCabecera, "Pedido", 2, "tCodigoPedido", 1000, 0, 0, "", _
                                   "Fecha", 2, "fRegistro", 1000, 0, 0, "dd/MM HH:mm", _
                                   "Caja", 2, "Caja", 1000, 0, 0, "", _
                                   "Salon", 2, "Salon", 1200, 0, 0, "", _
                                   "Mesa", 2, "Mesa", 1200, 0, 0, "", _
                                   "Observacion", 2, "Observacion", 1200, 0, 0, "", _
                                   "Usuario", 2, "tUsuario", 1000, 0, 0, "")
    Set grdCabecera.DataSource = RsCabecera
    
    Call ConfGrilla(7, grdDetalle, "Itm", 2, "tItem", 450, 2, 0, "", _
                                   "Ord", 2, "nOrden", 450, 2, 0, "", _
                                   "Producto", 2, "Producto", 3300, 0, 0, "", _
                                   "Cant.", 2, "nCantidad", 800, 1, 0, "##,##0.00", _
                                   "E", 2, "lImprime", 300, 2, 4, "", _
                                   "P", 2, "lPropiedad", 300, 2, 4, "", _
                                   "O", 2, "lObservacion", 300, 2, 4, "")
    Set grdDetalle.DataSource = RsDetalle
    
    Call ConfGrilla(7, grdCombo, "Itm", 2, "tItem", 450, 2, 0, "", _
                                 "Ord", 2, "nOrden", 450, 2, 0, "", _
                                 "Producto", 2, "Producto", 3300, 0, 0, "", _
                                 "Cant.", 2, "nCantidad", 800, 1, 0, "##,##0.00", _
                                 "E", 2, "lImprime", 300, 2, 4, "", _
                                 "P", 2, "lPropiedad", 300, 2, 4, "", _
                                 "O", 2, "lObservacion", 300, 2, 4, "")
    Set grdCombo.DataSource = RsCombo
                                   
    LlenaBusqueda
    txtCriterio = ""
   
    If RsCabecera.RecordCount = 0 Then
       sPedido = ""
    Else
       RsCabecera.MoveFirst
       sPedido = RsCabecera!tCodigoPedido
    End If
    
    Asignar
    cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmCheffControl = Nothing
End Sub

Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   AsignaCombo
End Sub

Private Sub grdCabecera_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    RsCabecera.Sort = grdCabecera.Columns(ColIndex).DataField & " ASC"
End Sub

Private Sub grdCabecera_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Asignar
End Sub

Private Sub grdDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   SubAsignar
End Sub

Private Sub TimerMensaje_Timer()
   If RsCabecera.RecordCount < Calcular("select count(tCodigoPedido) as Codigo from MPEDIDO where dbo.MPEDIDO.tTipoPedido in (" & sFiltro & ") and tEstadoPedido='" & RsParametro!tCheffControlEstadoPedido & "' and isnull(fAtencion,0)=0", Cn) Then
      Mensaje.Caption = "Tienes pedidos en Cola, favor de Refrescar la Pantalla"
   End If
End Sub

Public Sub Filtrar()
   If Len(Trim(txtCriterio.Text)) > 0 Then
      Select Case VarType(grdCabecera.Columns((grdCabecera.Columns(cboCriterio.Text).DataField)).Value)
             Case 1
                  sCriterio = ""
             Case 2 To 6
                  sCriterio = Trim(grdCabecera.Columns(cboCriterio.Text).DataField) & "= " & Val(txtCriterio.Text)
             Case 7
                  'sCriterio = ""
                  sCriterio = Trim(grdCabecera.Columns(cboCriterio.Text).DataField) & " >= #" & txtCriterio.Text & "# and " & Trim(grdCabecera.Columns(cboCriterio.Text).DataField) & " <= #" & txtCriterio.Text & " 23:59#"
             Case Else
                  sCriterio = Trim(grdCabecera.Columns(cboCriterio.Text).DataField) & " Like " & "'*" & txtCriterio.Text & "*'"
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

Public Sub Asignar()
    RsDetalle.Requery
    RsCombo.Requery
    txtPropiedad.Caption = ""
    txtObservacion.Caption = ""

    If RsCabecera.RecordCount = 0 Then
       sPedido = ""
       Exit Sub
    End If
    
    sPedido = RsCabecera!tCodigoPedido
    RsDetalle.Filter = "tCodigoPedido='" & sPedido & "'"
    
    'Producto vs Propiedad
    Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, dbo.TPROPIEDAD.tDetallado AS Descripcion, dbo.vOperador.Descripcion AS Operador " & _
           "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto LEFT OUTER JOIN dbo.vOperador ON dbo.TPROPIEDAD.tOperador = dbo.vOperador.Codigo " & _
           "WHERE tcodigoPedido='" & sPedido & "'"
    Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
    
    'Combo Propiedad
    Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, dbo.TCOMBOPROPIEDAD.tCodigoPropiedad, Dbo.TCOMBOPROPIEDAD.tProducto, dbo.TPROPIEDAD.tDetallado AS Descripcion, dbo.vOperador.Descripcion AS Operador " & _
           "FROM dbo.TCOMBOPROPIEDAD LEFT OUTER JOIN dbo.TPROPIEDAD ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto LEFT OUTER JOIN dbo.vOperador ON dbo.TPROPIEDAD.tOperador = dbo.vOperador.Codigo " & _
           "WHERE tcodigoPedido='" & sPedido & "'"
    Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
End Sub

Public Sub SubAsignar()
   txtPropiedadCombo.Caption = ""
   txtObservacionCombo.Caption = ""

   If Not RsDetalle.EOF Then
      RsProductoPropiedad.Filter = "tItem='" & RsDetalle!tItem & "'"
      If Not RsProductoPropiedad.EOF Then
         RsProductoPropiedad.MoveFirst
         Do While Not RsProductoPropiedad.EOF
            txtPropiedad.Caption = txtPropiedad.Caption & LTrim(RsProductoPropiedad!Operador) & " " & LTrim(RsProductoPropiedad!Descripcion) & ", "
            RsProductoPropiedad.MoveNext
         Loop
      End If
      txtObservacion.Caption = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
      RsCombo.Filter = "tCodigoPedido='" & sPedido & "' and tItem='" & RsDetalle!tItem & "'"
   Else
      RsCombo.Requery
   End If
End Sub

Public Sub AsignaCombo()
   If Not RsCombo.EOF Then
      txtPropiedadCombo.Caption = ""
      txtObservacionCombo.Caption = ""
      If Not RsCombo.EOF And Not RsDetalle.EOF Then
         RsComboPropiedad.Filter = "tItem='" & RsDetalle!tItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
         If Not RsComboPropiedad.EOF Then
            RsComboPropiedad.MoveFirst
            Do While Not RsComboPropiedad.EOF
               txtPropiedadCombo.Caption = txtPropiedadCombo.Caption & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!Descripcion) & ", "
               RsComboPropiedad.MoveNext
            Loop
         End If
      End If
      txtObservacionCombo.Caption = IIf(IsNull(RsCombo!tObservacion), "", RsCombo!tObservacion)
   End If
End Sub

Private Sub txtCriterio_KeyUp(KeyCode As Integer, Shift As Integer)
   Filtrar
End Sub
