VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBusquedaDelivery 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "frmBusquedaDelivery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInvitado 
      BackColor       =   &H00F2FAED&
      Caption         =   "&Invitado"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4320
      Width           =   1275
   End
   Begin VB.CommandButton Tienda 
      BackColor       =   &H00F2FAED&
      Caption         =   "Tiendas"
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
      Left            =   8130
      Picture         =   "frmBusquedaDelivery.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   4320
      Width           =   1275
   End
   Begin VB.TextBox txtCopia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9300
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "&Nuevo Cliente"
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
      Index           =   1
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   4320
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "&Modificar Cliente"
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
      Index           =   2
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4320
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Datos Principales"
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
      Index           =   3
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   60
      TabIndex        =   34
      Top             =   4800
      Width           =   11760
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   55
         Left            =   9675
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2430
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   45
         Left            =   8085
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2430
         Width           =   1590
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   54
         Left            =   9675
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   53
         Left            =   8880
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   52
         Left            =   8085
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   51
         Left            =   9675
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   930
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   50
         Left            =   8880
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   930
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   49
         Left            =   8085
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   930
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   48
         Left            =   9675
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   47
         Left            =   8880
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   46
         Left            =   8085
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "BackSpace"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   42
         Left            =   10470
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   930
         Width           =   1170
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H007EC642&
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Index           =   43
         Left            =   10470
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1680
         Width           =   1170
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   41
         Left            =   10470
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "Space"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   44
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3255
         Width           =   11550
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Key"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txtTempo 
         Height          =   285
         Left            =   10170
         TabIndex        =   0
         Top             =   3375
         Width           =   765
      End
      Begin MSWinsockLib.Winsock Winsock 
         Index           =   0
         Left            =   900
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock 
         Index           =   1
         Left            =   1320
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   555
      Index           =   5
      Left            =   10995
      Picture         =   "frmBusquedaDelivery.frx":053C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   495
      Width           =   885
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   555
      Index           =   6
      Left            =   10995
      Picture         =   "frmBusquedaDelivery.frx":0E06
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1050
      Width           =   885
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   555
      Index           =   7
      Left            =   10995
      Picture         =   "frmBusquedaDelivery.frx":16D0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1605
      Width           =   885
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   555
      Index           =   8
      Left            =   10995
      Picture         =   "frmBusquedaDelivery.frx":1F9A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   885
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   555
      Index           =   9
      Left            =   10995
      Picture         =   "frmBusquedaDelivery.frx":2864
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2715
      Width           =   885
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   555
      Index           =   10
      Left            =   10995
      Picture         =   "frmBusquedaDelivery.frx":312E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3270
      Width           =   885
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla_olf 
      Height          =   3000
      Left            =   12840
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   5292
      _LayoutType     =   4
      _RowHeight      =   28
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
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Listado"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
      _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=4,.fontname=Arial"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=-1"
      _StyleDefs(13)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(14)  =   ":id=2,.fontname=Arial"
      _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HE7FAB6&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
      _StyleDefs(42)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
      _StyleDefs(43)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(44)  =   ":id=33,.fontname=Small Fonts"
      _StyleDefs(45)  =   "Named:id=34:Heading"
      _StyleDefs(46)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(47)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(48)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(49)  =   ":id=34,.fontname=Arial"
      _StyleDefs(50)  =   "Named:id=35:Footing"
      _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   "Named:id=36:Selected"
      _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
      _StyleDefs(54)  =   ":id=36,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(55)  =   ":id=36,.fontname=Arial"
      _StyleDefs(56)  =   "Named:id=37:Caption"
      _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(58)  =   "Named:id=38:HighlightRow"
      _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.bold=-1,.fontsize=825"
      _StyleDefs(60)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(61)  =   ":id=38,.fontname=Arial"
      _StyleDefs(62)  =   "Named:id=39:EvenRow"
      _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(64)  =   "Named:id=40:OddRow"
      _StyleDefs(65)  =   ":id=40,.parent=33"
      _StyleDefs(66)  =   "Named:id=41:RecordSelector"
      _StyleDefs(67)  =   ":id=41,.parent=34"
      _StyleDefs(68)  =   "Named:id=42:FilterBar"
      _StyleDefs(69)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame fraDato 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Datos del Cliente "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   4850
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3930
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "min Entrega Delivery"
         Height          =   195
         Left            =   1920
         TabIndex        =   79
         Top             =   4485
         Width           =   1455
      End
      Begin VB.Label txtTiempoDelivery 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1170
         TabIndex        =   78
         Top             =   4440
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo :"
         Height          =   195
         Left            =   540
         TabIndex        =   77
         Top             =   4485
         Width           =   615
      End
      Begin VB.Label txtZonaReferencia 
         BackColor       =   &H00F0F0F0&
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
         Height          =   375
         Left            =   1170
         TabIndex        =   69
         Top             =   2730
         Width           =   2670
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Zona :"
         Height          =   195
         Index           =   9
         Left            =   345
         TabIndex        =   68
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label txtDireccion 
         BackColor       =   &H00F0F0F0&
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
         Height          =   600
         Left            =   1170
         TabIndex        =   65
         Top             =   1170
         Width           =   2670
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "min"
         Height          =   195
         Left            =   1920
         TabIndex        =   51
         Top             =   4155
         Width           =   240
      End
      Begin VB.Label txtTiempo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   1170
         TabIndex        =   33
         Top             =   4110
         Width           =   645
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo :"
         Height          =   195
         Index           =   10
         Left            =   540
         TabIndex        =   32
         Top             =   4155
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1920
         TabIndex        =   24
         Top             =   3795
         Width           =   150
      End
      Begin VB.Label txtdescuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   1170
         TabIndex        =   23
         Top             =   3750
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento :"
         Height          =   195
         Left            =   285
         TabIndex        =   22
         Top             =   3788
         Width           =   870
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observac. :"
         Height          =   195
         Index           =   8
         Left            =   330
         TabIndex        =   20
         Top             =   3210
         Width           =   825
      End
      Begin VB.Label txtObservacion 
         BackColor       =   &H00F0F0F0&
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
         Height          =   540
         Left            =   1170
         TabIndex        =   19
         Top             =   3150
         Width           =   2670
      End
      Begin VB.Label txtZona 
         BackColor       =   &H00F0F0F0&
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
         Height          =   285
         Left            =   1170
         TabIndex        =   18
         Top             =   2400
         Width           =   2670
      End
      Begin VB.Label txtReferencia 
         BackColor       =   &H00F0F0F0&
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
         Height          =   540
         Left            =   1170
         TabIndex        =   17
         Top             =   1815
         Width           =   2670
      End
      Begin VB.Label txtTelefono 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1170
         TabIndex        =   16
         Top             =   270
         Width           =   2670
      End
      Begin VB.Label txtNombre 
         BackColor       =   &H00F0F0F0&
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
         Left            =   1170
         TabIndex        =   15
         Top             =   885
         Width           =   2670
      End
      Begin VB.Label txtApellido 
         BackColor       =   &H00F0F0F0&
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
         Left            =   1170
         TabIndex        =   14
         Top             =   600
         Width           =   2670
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección :"
         Height          =   195
         Index           =   5
         Left            =   390
         TabIndex        =   13
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. :"
         Height          =   195
         Index           =   4
         Left            =   885
         TabIndex        =   12
         Top             =   315
         Width           =   270
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia :"
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   11
         Top             =   1815
         Width           =   870
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona :"
         Height          =   195
         Index           =   2
         Left            =   690
         TabIndex        =   10
         Top             =   2450
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres (RS) :"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   915
         Width           =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos (RC) :"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   8
         Top             =   630
         Width           =   1080
      End
   End
   Begin VB.Frame fraPunto 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000C0&
      Height          =   4530
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   3780
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Left            =   1950
         TabIndex        =   75
         Top             =   3015
         Width           =   1680
      End
      Begin VB.Label txtPedidos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Left            =   1950
         TabIndex        =   74
         Top             =   2677
         Width           =   1680
      End
      Begin VB.Label txtDesde 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Left            =   1950
         TabIndex        =   73
         Top             =   2340
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumo Total :"
         Height          =   195
         Index           =   20
         Left            =   690
         TabIndex        =   72
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Pedidos :"
         Height          =   195
         Index           =   19
         Left            =   510
         TabIndex        =   71
         Top             =   2722
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde :"
         Height          =   195
         Index           =   18
         Left            =   1290
         TabIndex        =   70
         Top             =   2385
         Width           =   555
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   6
         Left            =   585
         TabIndex        =   61
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto :"
         Height          =   195
         Index           =   7
         Left            =   585
         TabIndex        =   60
         Top             =   645
         Width           =   540
      End
      Begin VB.Label txtFechaUltima 
         BackColor       =   &H00F0F0F0&
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
         Left            =   1185
         TabIndex        =   59
         Top             =   315
         Width           =   2445
      End
      Begin VB.Label txtMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Left            =   1185
         TabIndex        =   58
         Top             =   645
         Width           =   2445
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento :"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   57
         Top             =   975
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social :"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   56
         Top             =   1635
         Width           =   1035
      End
      Begin VB.Label txtDocumento 
         BackColor       =   &H00F0F0F0&
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
         Left            =   1185
         TabIndex        =   55
         Top             =   975
         Width           =   2445
      End
      Begin VB.Label txtRazonSocial 
         BackColor       =   &H00F0F0F0&
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
         Height          =   465
         Left            =   1185
         TabIndex        =   54
         Top             =   1635
         Width           =   2445
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identidad :"
         Height          =   195
         Index           =   13
         Left            =   375
         TabIndex        =   53
         Top             =   1305
         Width           =   750
      End
      Begin VB.Label txtRuc 
         BackColor       =   &H00F0F0F0&
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
         Left            =   1185
         TabIndex        =   52
         Top             =   1305
         Width           =   2445
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puntos Acumulados :"
         Height          =   195
         Index           =   16
         Left            =   345
         TabIndex        =   31
         Top             =   3525
         Width           =   1500
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puntos Utilizados :"
         Height          =   195
         Index           =   15
         Left            =   540
         TabIndex        =   30
         Top             =   3855
         Width           =   1305
      End
      Begin VB.Label txtAcumulado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Left            =   1950
         TabIndex        =   29
         Top             =   3525
         Width           =   1680
      End
      Begin VB.Label txtUtilizado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Left            =   1950
         TabIndex        =   28
         Top             =   3855
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puntos Disponibles :"
         Height          =   195
         Index           =   14
         Left            =   405
         TabIndex        =   27
         Top             =   4185
         Width           =   1440
      End
      Begin VB.Label txtDisponible 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Left            =   1950
         TabIndex        =   26
         Top             =   4185
         Width           =   1680
      End
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   3135
      Left            =   3960
      TabIndex        =   86
      Top             =   360
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   5530
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   2
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
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(19)  =   ":id=6,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(51)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(52)  =   ":id=34,.fontname=Arial"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
      _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=900,.italic=0,.underline=0"
      _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=36,.fontname=Arial"
      _StyleDefs(60)  =   "Named:id=37:Caption"
      _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(62)  =   "Named:id=38:HighlightRow"
      _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.appearance=0"
      _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=-1,.fontsize=675,.italic=0,.underline=0"
      _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(66)  =   ":id=38,.fontname=Small Fonts"
      _StyleDefs(67)  =   "Named:id=39:EvenRow"
      _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=40:OddRow"
      _StyleDefs(70)  =   ":id=40,.parent=33"
      _StyleDefs(71)  =   "Named:id=41:RecordSelector"
      _StyleDefs(72)  =   ":id=41,.parent=34"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invitado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   85
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label txtInvitado 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5190
      TabIndex        =   84
      Top             =   3900
      Width           =   5745
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tienda :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   82
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label txtTienda 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tienda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5190
      TabIndex        =   80
      Top             =   3480
      Width           =   5745
   End
   Begin VB.Label txtResultado 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   5910
      TabIndex        =   67
      Top             =   0
      Width           =   5040
   End
   Begin VB.Label txtPredeterm 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F2FAED&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   3975
      TabIndex        =   66
      Top             =   0
      Width           =   1980
   End
End
Attribute VB_Name = "frmBusquedaDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public wAdiciona As Boolean
Public RsGrilla As Recordset

Dim sCriterio As String
Dim RsCliente As Recordset
Dim RsDireccion As Recordset
Dim RsDatos As Recordset
Dim RsDatoTienda As Recordset
Dim nPredeterm As Integer
Dim lConectado As Boolean

'invitado2013
Dim rsDatosInvitado As Recordset

'invitado2013
Private Sub cmdInvitado_Click()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
     If RsGrilla.RecordCount = 0 Then
       Exit Sub
    End If
    
    If txtTelefono.Caption = "" Then
        Exit Sub
    End If
     
    Isql = "select * from vInvitado where codigodelivery like '" & RsGrilla!codigo & "%' Order by Invitado"
    Call ConfGrilla(2, frmBusquedaRapidaInvitado.grdGrilla, "Código", 2, "Codigo", 2000, 2, 0, "", _
                                                    "Invitado", 2, "Invitado", 6000, 0, 0, "")

    sTemp = txtTienda.Caption
    frmBusquedaRapidaInvitado.nPredeterm = 1
    frmBusquedaRapidaInvitado.scodigoDelivery = RsGrilla!codigo
    frmBusquedaRapidaInvitado.Show vbModal
    If wEnter = True Then
        
        sCodigoInvitado = sCodigo
         txtInvitado.Caption = sDescrip
        
'        Asigna


    End If
End Sub
'invitado2013
Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   Select Case Index
          Case Is = 1 ' Agregar Cliente
               If lCCVOX Then
                  RsGrilla.Requery
                  sTemporal = txtResultado.Caption
                  Filtrar
                  Asigna
                  If txtCopia.Text = "" Then
                     txtCopia.Text = sTemporal
                  End If
                  txtTempo.SetFocus
               Else
                  wAdiciona = True
                  If txtPredeterm.Caption = "Id." Then
                     frmNuevoDelivery.txtTelefono.Text = txtResultado.Caption
                  End If
                   
                  frmNuevoDelivery.Show vbModal
                  If wEnter Then
                     RsGrilla.Filter = "Codigo ='" & sCodigo & "'"
                     RsGrilla.Requery
                     Asigna
                     wEnter = False
                  End If
                  txtTempo.SetFocus
               End If

          Case Is = 2 'Modificar
               If RsGrilla.RecordCount > 0 Then
                  wAdiciona = False
                  frmNuevoDelivery.Caption = "Modificar Cliente Delivery"
                  frmNuevoDelivery.Show vbModal
                  If wEnter Then
                     RsGrilla.Filter = "Codigo ='" & sCodigo & "'"
                     RsGrilla.Requery
                     Asigna
                     wEnter = False
                  End If
               Else
                  Exit Sub
               End If
               txtTempo.SetFocus

          Case Is = 3 ' Otros Datos
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               If cmdOpcion(3).Caption = "Datos Principales" Then
                  Asigna
                  cmdOpcion(3).Caption = "Otros Datos"
               Else
                  Dim sPedCalc As String
                  'Segundo Calculo
                  Isql = "SELECT COUNT(DISTINCT dbo.MPEDIDO.tCodigoPedido) AS nPedido, SUM(dbo.DPEDIDO.nVenta) AS nVenta " & _
                         "FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                         "WHERE (dbo.MPEDIDO.tEstadoPedido <> '03') AND (dbo.DPEDIDO.tEstadoItem = 'N') AND (dbo.MPEDIDO.tTipoPedido = '02') AND (dbo.MPEDIDO.tClienteDelivery = '" & RsGrilla!codigo & "') and " & _
                         "MPEDIDO.fRegistro >=" & Date - nDiasDelivery & ""
                  Set RsCliente = Lib.OpenRecordset(Isql, Cn)
                  If RsCliente.RecordCount = 0 Then
                     txtDesde.Caption = ""
                     txtPedidos.Caption = ""
                     txtTotal.Caption = ""
                  Else
                     txtDesde.Caption = FechaServidor() - nDiasDelivery
                     txtPedidos.Caption = Format(IIf(IsNull(RsCliente!nPedido), "", RsCliente!nPedido), "###,##0.00")
                     txtTotal.Caption = Format(IIf(IsNull(RsCliente!nVenta), "", RsCliente!nVenta), "###,##0.00")
                  End If
                  
                  sPedCalc = Calcular("SELECT MAX(tCodigoPedido) AS Codigo From dbo.MPEDIDO where tClienteDelivery ='" & RsGrilla!codigo & "' and tEstadoPedido <>'03' GROUP BY tClienteDelivery", Cn)
                  Isql = "SELECT dbo.DPEDIDO.tCodigoPedido, SUM(dbo.DPEDIDO.nVenta) AS Monto, MAX(dbo.MDOCUMENTO.fRegistro) AS Fecha, MAX(dbo.MDOCUMENTO.tDocumento) AS Documento, MAX(dbo.vCliente.Descripcion) As Cliente, MAX(dbo.vCliente.tIdentidad) As Identidad " & _
                         "FROM dbo.vCliente RIGHT OUTER JOIN dbo.MDOCUMENTO ON dbo.vCliente.Codigo = dbo.MDOCUMENTO.tCodigoCliente RIGHT OUTER JOIN dbo.DPEDIDO ON dbo.MDOCUMENTO.tDocumento = dbo.DPEDIDO.tDocumento " & _
                         "where tCodigoPedido ='" & sPedCalc & "' " & _
                         "GROUP BY dbo.DPEDIDO.tCodigoPedido"
                  Set RsCliente = Lib.OpenRecordset(Isql, Cn)
                  If RsCliente.RecordCount = 0 Then
                     txtFechaUltima.Caption = ""
                     txtDocumento.Caption = ""
                     txtMonto.Caption = "0.00"
                     txtRazonSocial.Caption = ""
                     txtRuc.Caption = ""
                  Else
                     txtFechaUltima.Caption = IIf(IsNull(RsCliente!Fecha), "", RsCliente!Fecha)
                     txtDocumento.Caption = IIf(IsNull(RsCliente!Documento), "", RsCliente!Documento)
                     txtMonto.Caption = Format(IIf(IsNull(RsCliente!Monto), "", RsCliente!Monto), "###,##0.00")
                     txtRazonSocial.Caption = IIf(IsNull(RsCliente!Cliente), "", RsCliente!Cliente)
                     txtRuc.Caption = IIf(IsNull(RsCliente!Identidad), "", RsCliente!Identidad)
                  End If
                  cmdOpcion(3).Caption = "Datos Principales"
               End If
               fraDato.Visible = Not fraDato.Visible
               txtTempo.SetFocus
               
          Case Is = 5  ' Primero
               MoverPuntero Primero, grdGrilla
               txtTempo.SetFocus

          Case Is = 6  ' PG Up
               MoverPuntero pgup, grdGrilla
               txtTempo.SetFocus

          Case Is = 7  ' Previo
               MoverPuntero previo, grdGrilla
               txtTempo.SetFocus

          Case Is = 8  ' Siguiente
               MoverPuntero siguiente, grdGrilla
               txtTempo.SetFocus

          Case Is = 9  ' PG Dn
               MoverPuntero pgdn, grdGrilla
               txtTempo.SetFocus

          Case Is = 10 ' Ultimo
               MoverPuntero Ultimo, grdGrilla
               txtTempo.SetFocus
   End Select

End Sub

Private Sub Form_Activate()
   txtPredeterm.Caption = grdGrilla.Columns(nPredeterm).Caption
End Sub

Private Sub Form_Load()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Centrar Me
   wEnter = False
   sCodigo = ""
   sDescrip = ""
   MatrizBotones 40, 10, 1, 1, cmdKey
   LlenaKey
   nTiempoDelivery = Calcular("select nTiempoMinutoCD as Codigo FROM TPARAMETRO", Cn)
   txtTiempoDelivery.Caption = nTiempoDelivery
   
   'Isql debe ser defino antes de ser llamado este formulario
   Isql = "SELECT dbo.TDELIVERY.tCodigoDelivery AS Codigo, dbo.TDELIVERY.tTelefono, LTRIM(dbo.TDELIVERY.tApellido) + ' ' + LTRIM(dbo.TDELIVERY.tNombre) AS Cliente, dbo.TDELIVERY.nDescuento,ltrim(dbo.TDELIVERY.treferencia) as treferencia, dbo.vZona.Descripcion AS Zona FROM dbo.TDELIVERY  LEFT OUTER JOIN dbo.vZona ON dbo.TDELIVERY.tZona = dbo.vZona.Codigo Where (dbo.TDELIVERY.lActivo = 1)"
   Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
   
   Call ConfGrilla(3, grdGrilla, "Id.", 2, "tTelefono", 1150, 2, 0, "", _
                                 "Cliente", 2, "Cliente", 3700, 0, 0, "", _
                                 "Referencia", 2, "tReferencia", 1800, 0, 0, "")
      
   Set grdGrilla.DataSource = RsGrilla
   txtResultado.Caption = sTemp
   Screen.MousePointer = vbDefault
   nPredeterm = 0
   If lCCVOX Then
      txtResultado.Width = 3050
      txtCopia.Visible = True
      
      WinSock(0).Close
      WinSock(0).LocalPort = "5000"
      WinSock(0).Listen
      
      WinSock(1).RemoteHost = WinSock(1).LocalIP
      WinSock(1).RemotePort = "5001"
      WinSock(1).Close
      If ConectarWinSock Then
         Call WinSock(1).SendData("Ok" & vbCrLf)
      Else
         MsgBox "Error de concectividad con el CCVOX", vbCritical
         WinSock(1).Close
      End If
       
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If lCCVOX Then
      WinSock(1).Close
      WinSock(0).Close
   End If
   Set RsCliente = Nothing
   Set RsGrilla = Nothing
   Set RsDireccion = Nothing
   Set frmBusquedaDelivery = Nothing
End Sub

Private Sub grdGrilla_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
End Sub

Private Sub grdGrilla_DblClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   If RsGrilla.RecordCount = 0 Then
      Exit Sub
   End If
   
   wEnter = True
   sCodigo = IIf(RsGrilla.EOF = True, "", RsGrilla!codigo)
   sDescrip = IIf(RsGrilla.EOF = True, "", RsGrilla!Cliente)
   sVar1 = IIf(IsNull(RsGrilla!tTelefono), "", RsGrilla!tTelefono)
'   sVar2 = IIf(IsNull(RsGrilla!tTienda), "", RsGrilla!tTienda)
   nVar1 = IIf(IsNull(RsGrilla!nDescuento), 0, Format(RsGrilla!nDescuento, "##0.00"))
   Unload Me
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If

    Screen.MousePointer = vbHourglass
    If Right(RsGrilla.Sort, 3) = "ASC" Then
       RsGrilla.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    Else
       RsGrilla.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    End If
    RsGrilla.Requery
    Screen.MousePointer = vbDefault
    nPredeterm = ColIndex
    txtPredeterm.Caption = grdGrilla.Columns(nPredeterm).Caption
    txtTempo.SetFocus
End Sub

Public Sub Asigna()
   If RsGrilla.RecordCount > 0 Then
      Set RsDatos = Lib.OpenRecordset("select * from vDelivery where Codigo='" & RsGrilla!codigo & "'", Cn)
      Set RsDatoTienda = Lib.OpenRecordset("select * from vTienda where tCodigoDelivery ='" & RsGrilla!codigo & "' and Codigo='" & sTienda & "'", Cn)
        'invitado2013
      Set rsDatosInvitado = Lib.OpenRecordset("select * from tdeliveryinvitado where tcodigoinvitado='" & sCodigoInvitado & "'", Cn)
      'invitado2013
      txtNombre.Caption = IIf(IsNull(RsDatos!tNombre), "", RsDatos!tNombre)
      txtTelefono.Caption = IIf(IsNull(RsDatos!tTelefono), "", RsDatos!tTelefono)
      txtReferencia.Caption = IIf(IsNull(RsDatos!tReferencia), "", RsDatos!tReferencia)
      txtObservacion.Caption = IIf(IsNull(RsDatos!tObservacion), "", RsDatos!tObservacion)
      txtDescuento.Caption = IIf(IsNull(RsDatos!nDescuento), "0.00", Format(RsDatos!nDescuento, "##0.00"))
      txtAcumulado.Caption = IIf(IsNull(RsDatos!nAcumulado), "0.00", Format(RsDatos!nAcumulado, "##,###,##0.00"))
      txtUtilizado.Caption = IIf(IsNull(RsDatos!nUtilizado), "0.00", Format(RsDatos!nUtilizado, "##,###,##0.00"))
      txtDisponible.Caption = IIf(IsNull(RsDatos!nDisponible), "0.00", Format(RsDatos!nDisponible, "##,###,##0.00"))
      txtDireccion.Caption = IIf(txtTienda.Caption = "", IIf(IsNull(RsDatos!tDireccion), "", RsDatos!tDireccion), "")
      txtZona.Caption = IIf(IsNull(RsDatos!tZona), "", RsDatos!tZona)
      If RsDatoTienda.RecordCount > 0 Then
         txtApellido.Caption = IIf(IsNull(RsDatos!tApellido), "", RsDatos!tApellido) & IIf(txtTienda.Caption = "", "", " - " & txtTienda.Caption)
         txtTienda.Caption = IIf(IsNull(RsDatoTienda!Descripcion), "", RsDatoTienda!Descripcion)
      Else
         txtTienda.Caption = ""
         txtApellido.Caption = IIf(IsNull(RsDatos!tApellido), "", RsDatos!tApellido)
      End If
         'invitado2013
      If rsDatosInvitado.RecordCount > 0 Then
            txtInvitado = IIf(IsNull(rsDatosInvitado!tNombre), "", rsDatosInvitado!tNombre) & " " & IIf(IsNull(rsDatosInvitado!tApellido), "", rsDatosInvitado!tApellido)
      Else
            txtInvitado.Caption = ""
      End If
        'invitado2013
   Else
      txtResultado.Caption = ""
      txtNombre.Caption = ""
      txtApellido.Caption = ""
      txtTelefono.Caption = ""
      txtDireccion.Caption = ""
      txtReferencia.Caption = ""
      txtObservacion.Caption = ""
      txtZona.Caption = ""
      txtZonaReferencia.Caption = ""
      txtDescuento.Caption = ""
      txtTienda.Caption = ""
              'invitado2013
      txtInvitado.Caption = ""
        'invitado2013
   End If
   Set RsDatos = Nothing
   Set RsDatoTienda = Nothing
   txtTempo.SetFocus
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
   txtTempo.SetFocus
   Limpiar
   If Not RsGrilla.EOF Then
      Asigna
   End If
End Sub

Private Sub cmdkey_Click(Index As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If

    Select Case Index
           Case Is = 41 'Escape
                wEnter = False
                Unload Me
           
           Case Is = 42 'Backspace
                If Len(Trim(txtResultado.Caption)) > 0 Then
                   txtResultado.Caption = Mid(txtResultado.Caption, 1, Len(txtResultado.Caption) - 1)
                Else
                   txtResultado.Caption = ""
                End If
                txtTempo.SetFocus
                
           Case Is = 43 'Enter
                'If RsGrilla.RecordCount = 0 Or txtResultado.Caption = "" Then
                If RsGrilla.RecordCount = 0 Then
                   Exit Sub
                End If
                wEnter = True
                sCodigo = IIf(RsGrilla.EOF = True, "", RsGrilla!codigo)
                sDescrip = IIf(RsGrilla.EOF = True, "", RsGrilla!Cliente)
                'sVar1 = sTienda
                Unload Me
                
           Case Is = 44 'Espacio
                txtResultado.Caption = txtResultado.Caption + " "
                
           Case Else
                If cmdKey(Index).Caption = "&&" Then
                   txtResultado.Caption = txtResultado.Caption + "&"
                Else
                   txtResultado.Caption = txtResultado.Caption + cmdKey(Index).Caption
                End If
                txtTempo.SetFocus
    End Select
    
    '       '    numeros   letras    sp  numpad     delet  ñ
    '       Case 48 To 57, 65 To 90, 32, 96 To 105, 8, 46, 192
    '            Filtrar
End Sub

Public Sub LlenaKey()
   cmdKey(1).Caption = "!"
   cmdKey(2).Caption = "_"
   cmdKey(3).Caption = "#"
   cmdKey(4).Caption = "$"
   cmdKey(5).Caption = "%"
   cmdKey(6).Caption = "&&"
   cmdKey(7).Caption = "/"
   cmdKey(8).Caption = "("
   cmdKey(9).Caption = ")"
   cmdKey(10).Caption = "="
   cmdKey(11).Caption = "Q"
   cmdKey(12).Caption = "W"
   cmdKey(13).Caption = "E"
   cmdKey(14).Caption = "R"
   cmdKey(15).Caption = "T"
   cmdKey(16).Caption = "Y"
   cmdKey(17).Caption = "U"
   cmdKey(18).Caption = "I"
   cmdKey(19).Caption = "O"
   cmdKey(20).Caption = "P"
   cmdKey(21).Caption = "A"
   cmdKey(22).Caption = "S"
   cmdKey(23).Caption = "D"
   cmdKey(24).Caption = "F"
   cmdKey(25).Caption = "G"
   cmdKey(26).Caption = "H"
   cmdKey(27).Caption = "J"
   cmdKey(28).Caption = "K"
   cmdKey(29).Caption = "L"
   cmdKey(30).Caption = "Ñ"
   cmdKey(31).Caption = "Z"
   cmdKey(32).Caption = "X"
   cmdKey(33).Caption = "C"
   cmdKey(34).Caption = "V"
   cmdKey(35).Caption = "B"
   cmdKey(36).Caption = "N"
   cmdKey(37).Caption = "M"
   cmdKey(38).Caption = ","
   cmdKey(39).Caption = "."
   cmdKey(40).Caption = "-"
   cmdKey(41).Caption = "ESC"
   cmdKey(42).Caption = "BackSpace"
   cmdKey(43).Caption = "Enter"
   cmdKey(44).Caption = "SPACE"
   cmdKey(45).Caption = "0"
   cmdKey(46).Caption = "1"
   cmdKey(47).Caption = "2"
   cmdKey(48).Caption = "3"
   cmdKey(49).Caption = "4"
   cmdKey(50).Caption = "5"
   cmdKey(51).Caption = "6"
   cmdKey(52).Caption = "7"
   cmdKey(53).Caption = "8"
   cmdKey(54).Caption = "9"
   cmdKey(55).Caption = "."
End Sub

Private Sub Tienda_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    
    If RsGrilla.RecordCount = 0 Then
       Exit Sub
    End If
    
    
    Isql = "select * from vTienda where lActivo = 1 and tCodigoDelivery='" & IIf(RsGrilla.EOF = True, "", RsGrilla!codigo) & "' Order by Descripcion"
    Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                    "Tienda", 2, "Descripcion", 3000, 0, 0, "", _
                                                    "Direccion", 2, "tDireccion", 4100, 0, 0, "")

    sTemp = txtTienda.Caption
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
        sTienda = sCodigo
        txtTienda.Caption = sDescrip
        Asigna
        txtDireccion.Caption = Calcular("select tDireccion as Codigo FROM TTIENDA where tCodigoDelivery='" & RsGrilla!codigo & "' and tCodigoTienda='" & sCodigo & "'", Cn)
    End If
End Sub

Private Sub txtCopia_Change()
   txtResultado.Caption = txtCopia.Text
   Filtrar
End Sub

Private Sub txtresultado_Change()
    txtCopia.Text = txtResultado.Caption
    Filtrar
End Sub

Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   If Shift > 1 Then
      Select Case KeyCode
       Case 48
            txtResultado.Caption = txtResultado.Caption + "="
       Case 49
            txtResultado.Caption = txtResultado.Caption + "!"
       Case 50
            txtResultado.Caption = txtResultado.Caption + "_"
       Case 51
            txtResultado.Caption = txtResultado.Caption + "#"
       Case 52
            txtResultado.Caption = txtResultado.Caption + "$"
       Case 53
            txtResultado.Caption = txtResultado.Caption + "%"
       Case 54
            txtResultado.Caption = txtResultado.Caption + "&"
       Case 55
            txtResultado.Caption = txtResultado.Caption + "/"
       Case 56
            txtResultado.Caption = txtResultado.Caption + "("
       Case 57
            txtResultado.Caption = txtResultado.Caption + ")"
       Case Else
    
      End Select
   Else
   
      Select Case KeyCode
          Case 13
               Call cmdkey_Click(43)
          Case 27
               Call cmdkey_Click(41)
          Case 8
               Call cmdkey_Click(42)
          Case 32
               txtResultado.Caption = txtResultado.Caption + " "
          Case 65
               txtResultado.Caption = txtResultado.Caption + "A"
          Case 66
               txtResultado.Caption = txtResultado.Caption + "B"
          Case 67
               txtResultado.Caption = txtResultado.Caption + "C"
          Case 68
               txtResultado.Caption = txtResultado.Caption + "D"
          Case 69
               txtResultado.Caption = txtResultado.Caption + "E"
          Case 70
               txtResultado.Caption = txtResultado.Caption + "F"
          Case 71
               txtResultado.Caption = txtResultado.Caption + "G"
          Case 72
               txtResultado.Caption = txtResultado.Caption + "H"
          Case 73
               txtResultado.Caption = txtResultado.Caption + "I"
          Case 74
               txtResultado.Caption = txtResultado.Caption + "J"
          Case 75
               txtResultado.Caption = txtResultado.Caption + "K"
          Case 76
               txtResultado.Caption = txtResultado.Caption + "L"
          Case 77
               txtResultado.Caption = txtResultado.Caption + "M"
          Case 78
               txtResultado.Caption = txtResultado.Caption + "N"
          Case 79
               txtResultado.Caption = txtResultado.Caption + "O"
          Case 80
               txtResultado.Caption = txtResultado.Caption + "P"
          Case 81
               txtResultado.Caption = txtResultado.Caption + "Q"
          Case 82
               txtResultado.Caption = txtResultado.Caption + "R"
          Case 83
               txtResultado.Caption = txtResultado.Caption + "S"
          Case 84
               txtResultado.Caption = txtResultado.Caption + "T"
          Case 85
               txtResultado.Caption = txtResultado.Caption + "U"
          Case 86
               txtResultado.Caption = txtResultado.Caption + "V"
          Case 87
               txtResultado.Caption = txtResultado.Caption + "W"
          Case 88
              txtResultado.Caption = txtResultado.Caption + "X"
          Case 89
               txtResultado.Caption = txtResultado.Caption + "Y"
          Case 90
               txtResultado.Caption = txtResultado.Caption + "Z"
          Case 96, 48
               txtResultado.Caption = txtResultado.Caption + "0"
          Case 97, 49
               txtResultado.Caption = txtResultado.Caption + "1"
          Case 98, 50
               txtResultado.Caption = txtResultado.Caption + "2"
          Case 99, 51
               txtResultado.Caption = txtResultado.Caption + "3"
          Case 100, 52
               txtResultado.Caption = txtResultado.Caption + "4"
          Case 101, 53
               txtResultado.Caption = txtResultado.Caption + "5"
          Case 102, 54
               txtResultado.Caption = txtResultado.Caption + "6"
          Case 103, 55
               txtResultado.Caption = txtResultado.Caption + "7"
          Case 104, 56
               txtResultado.Caption = txtResultado.Caption + "8"
          Case 105, 57
               txtResultado.Caption = txtResultado.Caption + "9"
          Case 110, 190
               txtResultado.Caption = txtResultado.Caption + "."
          Case 106, 188
               txtResultado.Caption = txtResultado.Caption + ","
          Case 109, 189
               txtResultado.Caption = txtResultado.Caption + "-"
          Case 192
               txtResultado.Caption = txtResultado.Caption + "Ñ"
          Case Else

      End Select
    End If
End Sub

Public Sub Filtrar()
   If Len(Trim(txtResultado.Caption)) > 0 Then
      Select Case VarType(grdGrilla.Columns(nPredeterm).value)
             Case 2 To 6
             sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & "= " & val(txtResultado.Caption)
             Case 7
             sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & " >= #" & txtResultado.Caption & "# and " & Trim(grdGrilla.Columns(nPredeterm).DataField) & " <= #" & txtResultado.Caption & " 23:59#"
             Case Else
             sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & " Like " & "'*" & txtResultado.Caption & "*'"
      End Select
      
      Screen.MousePointer = vbHourglass
      With RsGrilla
           .Filter = sCriterio
           If .BOF = False Then
              .MoveFirst
           End If
      End With
      Screen.MousePointer = vbDefault
   Else
      Screen.MousePointer = vbHourglass
      RsGrilla.Filter = adFilterNone
      RsGrilla.Requery
      If RsGrilla.RecordCount <> 0 Then
         RsGrilla.MoveFirst
      End If
      Screen.MousePointer = vbDefault
      sCriterio = ""
   End If
   On Error GoTo g:
    txtTempo.SetFocus
   Exit Sub
g:
'
End Sub

Public Sub Limpiar()
      txtNombre.Caption = ""
      txtApellido.Caption = ""
      txtTelefono.Caption = ""
      txtDireccion.Caption = ""
      txtZona.Caption = ""
      txtZonaReferencia.Caption = ""
      txtTiempo.Caption = "0.00"
            
      txtReferencia.Caption = ""
      txtObservacion.Caption = ""
      txtDescuento.Caption = "0.00"
      
      txtAcumulado.Caption = ""
      txtUtilizado.Caption = ""
      txtDisponible.Caption = ""
      
      txtFechaUltima.Caption = ""
      txtDocumento.Caption = ""
      txtMonto.Caption = "0.00"
      txtRazonSocial.Caption = ""
      txtRuc.Caption = ""
      txtTienda.Caption = ""
      cmdOpcion(3).Caption = "Otros Datos"
      txtTempo.SetFocus
End Sub
Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   WinSock(0).Close
   WinSock(0).Accept requestID
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   Dim buffer As String
   If WinSock(0).State <> 8 Then
      WinSock(0).GetData buffer
      If Len(buffer) > 0 Then
         If IsNumeric(Right(buffer, 2)) Then
            txtCopia.Text = LTrim(buffer)
         Else
            txtCopia.Text = Left(LTrim(buffer), Len(buffer) - 2)
         End If
      End If
   End If
   WinSock(0).Close
   WinSock(0).Listen
End Sub

Public Function ConectarWinSock() As Boolean
   Dim i As Integer
   i = 0
   ConectarWinSock = False
   WinSock(1).Connect
   Do While WinSock(1).State <> sckConnected
      i = i + 1
      If i > 1000 Then
         ConectarWinSock = False
         Exit Function
      End If
      DoEvents
   Loop
   ConectarWinSock = True
End Function

Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   WinSock(0).Close
   WinSock(1).Close
End Sub
