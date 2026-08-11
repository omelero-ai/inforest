VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmBusquedaValeConsumo 
   BackColor       =   &H00F0F0F0&
   Caption         =   "Búsqueda de Vales de Consumo"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FAED&
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   4
      Left            =   10200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FAED&
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   3
      Left            =   10200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FAED&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   2
      Left            =   10200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1200
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FAED&
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   1
      Left            =   10200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   480
      Width           =   1275
   End
   Begin VB.Timer Timer_Aceptarvisor 
      Interval        =   1000
      Left            =   9480
      Top             =   120
   End
   Begin VB.CommandButton cmdaceptarvisor 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FAED&
      Caption         =   "ACEP- VISOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   10200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      Height          =   3975
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   11220
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
         Left            =   9270
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2430
         Width           =   750
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
         Left            =   7770
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2430
         Width           =   1500
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
         Left            =   9270
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   750
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
         Left            =   8520
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   750
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
         Left            =   7770
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   180
         Width           =   750
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
         Left            =   9270
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   750
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
         Left            =   8520
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   930
         Width           =   750
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
         Left            =   7770
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   750
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
         Left            =   9270
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   750
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
         Left            =   8520
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1680
         Width           =   750
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
         Left            =   7770
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   750
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
         Left            =   10020
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   930
         Width           =   1125
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
         Left            =   10020
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1680
         Width           =   1125
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
         Left            =   10020
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1125
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   750
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
         Height          =   675
         Index           =   44
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3195
         Width           =   11055
      End
      Begin VB.TextBox txtTempo 
         Height          =   405
         Left            =   7410
         TabIndex        =   26
         Top             =   3255
         Width           =   3525
      End
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   600
      Index           =   9
      Left            =   9120
      Picture         =   "frmBusquedaValeConsumo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2910
      Width           =   960
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   600
      Index           =   8
      Left            =   9060
      Picture         =   "frmBusquedaValeConsumo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   960
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   600
      Index           =   7
      Left            =   9060
      Picture         =   "frmBusquedaValeConsumo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1710
      Width           =   960
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   600
      Index           =   6
      Left            =   9060
      Picture         =   "frmBusquedaValeConsumo.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1110
      Width           =   960
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   600
      Index           =   5
      Left            =   9060
      Picture         =   "frmBusquedaValeConsumo.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   510
      Width           =   960
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Height          =   600
      Index           =   10
      Left            =   9060
      Picture         =   "frmBusquedaValeConsumo.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3510
      Width           =   960
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6588
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
      ForeColor       =   &H00232323&
      Height          =   405
      Left            =   2070
      TabIndex        =   1
      Top             =   0
      Width           =   6915
   End
   Begin VB.Label txtPredeterm 
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
      ForeColor       =   &H007EC642&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2100
   End
End
Attribute VB_Name = "frmBusquedaValeConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Public RsGrillaVConsumo As Recordset
Dim RsDatosClienteFac As Recordset
Public nPredeterm As Integer
Dim sCriterio As String
Public wAdiciona As Boolean
'Dim xTipoDocumento As String
Dim RsPedido As Recordset
Public tipoBusqueda As String
Public sTipoDocumBusq As String
Public xbermur As String
Public sMonto As Double

Public sTarjetaVCTemp As String
Dim sTemp As String

Private Sub cmdaceptarvisor_Click()
            Timer_Aceptarvisor.Interval = 0
            sCodigo = Calcular("SELECT top 1 tcodigocliente FROM TCLIENTE where tidentidad = (select top 1 rucfactura from INFOVISOR where tcaja='" & sCaja & "')", Cn)
            cmdaceptarvisor.Visible = False
           
            Unload Me
End Sub

Private Sub cmdkey_Click(Index As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    
    Select Case Index
           Case Is = 41 'Escape
                wEnter = False
                Unload Me
                VSApForBusqueda = True
           Case Is = 42 'Backspace
                If Len(Trim(txtResultado.Caption)) > 0 Then
                   txtResultado.Caption = Mid(txtResultado.Caption, 1, Len(txtResultado.Caption) - 1)
                   txtTempo.Text = txtResultado.Caption
                Else
                   txtResultado.Caption = ""
                   txtTempo.Text = txtResultado.Caption
                End If
                txtTempo.SetFocus
                txtTempo.SelStart = Len(txtTempo)
                
           Case Is = 43 'Enter

               Call cmdOpcion_Click(1)

               '  If RsGrillaVConsumo.RecordCount = 0 Then
               '     Exit Sub
               '  End If
               '  wEnter = True
               '  sCodigo = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!codigo)
               '  sDescrip = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!descripcion)
               '    If xbermur = "SI" Then
               '          sMontoBermur = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!MontoDisponible)
               '     End If
               '  If tipoBusqueda = "BusquedaCajaRapida" Then
               '      tenlacebusqueda = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!tEnlace)
               '      tipoBusqueda = ""
               '  Else
               '      If tipoBusqueda = "BusquedaVenta" Then
               '          tenlacebusquedaVenta = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!tEnlace)
               '          tipoBusqueda = ""
               '      End If
               '  End If
               
               '' If MESA247 Then
               ''     If validacionmesa247(sCodigo) = True Then
               ''         Exit Sub
               ''     End If
               '' End If
                
                'Unload Me
                'VSApForBusqueda = False
           Case Is = 44 'Espacio
                txtResultado.Caption = txtResultado.Caption + " "
                txtTempo.Text = txtResultado.Caption
                txtTempo.SetFocus
                txtTempo.SelStart = Len(txtTempo)
           Case Else
                If cmdKey(Index).Caption = "&&" Then
                   txtResultado.Caption = txtResultado.Caption + "&"
                Else
                   txtResultado.Caption = txtResultado.Caption + cmdKey(Index).Caption
                End If
                txtTempo.Text = txtResultado.Caption
                txtTempo.SetFocus
                txtTempo.SelStart = Len(txtTempo) 'lleva el cursor al final
    End Select
    
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   Select Case Index
          Case Is = 0 ' Mostrar Resultado
               If cmdOpcion(0).Caption = "Todas las Und. de Negocios" Then
                  Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lCanal4=1", "lCanal5=1")))) & " Order by Descripcion"
                  cmdOpcion(0).Caption = "Und. Negocios Predeterminado"
               Else
                  Isql = "select * from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "')and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lCanal4=1", "lCanal5=1")))) & " Order by Descripcion"
                  cmdOpcion(0).Caption = "Todas las Und. de Negocios"
               End If
               cmdOpcion(4).FontBold = False
               txtResultado.Caption = ""
               Set RsGrillaVConsumo = Lib.OpenRecordset(Isql, Cn)
               Set grdGrilla.DataSource = RsGrillaVConsumo
          
          Case Is = 1 ' Agregar
               wAdiciona = True
            '    '------VALIDA EMAIL-----
            '    'frmNuevoCliente.ObtenerTipoDocumento xTipoDocumento
            '    frmNuevoCliente.Show vbModal
            '    '-----------------------
            '    RsGrillaVConsumo.Requery
            '    RsGrillaVConsumo.Find ("Codigo ='" & sCodigo & "'")
            '    sDescrip = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!descripcion)
            '    Unload Me
               Dim EsQr As Boolean
               EsQr = False
               Dim textoVale As String
               Dim cadena() As String
               Dim cadena2() As String
               Dim cadena3() As String
               textoVale = txtResultado.Caption
               cadena = Split(Trim(textoVale), "|")
               cadena2 = Split(Trim(textoVale), "]")
               cadena3 = Split(Trim(textoVale), "Ç")
    
              'MsgBox UBound(cadena), vbCritical, sMensaje
               
               If UBound(cadena) > 0 Then
                    EsQr = True
                    textoVale = cadena(0)
                    'MsgBox textoVale, vbCritical, sMensaje
               End If
               If UBound(cadena2) > 0 Then
                    EsQr = True
                    textoVale = cadena2(0)
                    'MsgBox textoVale, vbCritical, sMensaje
               End If
               If UBound(cadena3) > 0 Then
                    EsQr = True
                    textoVale = cadena3(0)
                    'MsgBox textoVale, vbCritical, sMensaje
               End If
                
                Dim RsValeConsumo As ADODB.Recordset
                Isql = "SELECT * FROM VALE_CONSUMO Where lESTADO_RECEPCION = 1 And lESTADO_ENVIO = 1 And lESTADO_USO = 0 And tCOD_VALE = '" & textoVale & "' AND CONVERT(VARCHAR(20),fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23)"

                Set RsValeConsumo = Lib.OpenRecordset(Isql, Cn)

                If RsValeConsumo.RecordCount > 0 Then

                    Dim sCOD_Vale_Temp As String
                    sCOD_Vale_Temp = Calcular("select top 1 tCOD_VALE as Codigo from " & sTarjetaVC_Final & " where tCOD_VALE = '" & textoVale & "' ", Cn)
                    If sCOD_Vale_Temp <> "0" Then
                        MsgBox "El Nro Vale " & sCOD_Vale_Temp & " ya existe en su Lista de Tarjetas a Cancelar.", vbInformation, sMensaje
                        Exit Sub
                    End If

                    Dim sCOD_Vale As String
                    sCOD_Vale = Calcular("select top 1 1 as Codigo from " & sTarjetaVCTemp & " where tCOD_VALE = '" & textoVale & "' ", Cn)

                    If sCOD_Vale = "0" Then
                        nCorrelaVC = Calcular("select max(tIndex) as Codigo from " & sTarjetaVCTemp, Cn)
                        If IsNull(nCorrelaVC) Or nCorrelaVC = "" Then
                            nCorrelaVC = "001"
                        Else
                            nCorrelaVC = Lib.Correlativo(nCorrelaVC, 3)
                        End If
                        Isql = "insert into " & sTarjetaVCTemp & _
                                "( tIndex, tTICKET, tCOD_VALE, tMONEDA, nMONTO_VALE) " & _
                                "Values( '" & nCorrelaVC & "', " _
                                        & "'" & RsValeConsumo!tTICKET & "'," _
                                        & "'" & RsValeConsumo!tCOD_VALE & "'," _
                                        & "'" & RsValeConsumo!tMONEDA & "'," _
                                        & "'" & RsValeConsumo!nMONTO_VALE & "')"
                        Cn.Execute Isql
                        RsGrillaVConsumo.Requery
                        txtResultado.Caption = ""
                         txtTempo.Text = txtResultado.Caption
                        Exit Sub
                    Else
                        MsgBox "Ya existe el mismo nro Vale en su Lista.", vbInformation, sMensaje
                        txtTempo.SetFocus
                        txtResultado.Caption = ""
                         txtTempo.Text = txtResultado.Caption
                        Exit Sub
                    End If

                Else
                    MsgBox "No hay datos para mostrar", vbCritical, sMensaje
                    If EsQr Then
                         txtResultado.Caption = ""
                         txtTempo.Text = txtResultado.Caption
                    End If
                    txtTempo.SetFocus
                    Exit Sub
                End If

          Case Is = 2 ' Eliminar
               If RsGrillaVConsumo.RecordCount <> 0 Then
                    wAdiciona = False
                  
                    Cn.Execute "delete from " & sTarjetaVCTemp & " where tIndex = '" & RsGrillaVConsumo!tIndex & "'"
                    RsGrillaVConsumo.Requery

               Else
                    MsgBox "No hay datos para eliminar", vbCritical, sMensaje
                    txtTempo.SetFocus
                    Exit Sub
               End If
                    
          Case Is = 3 ' Limpiar
               'Set RsGrillaVConsumo = Nothing
               Cn.Execute "delete from " & sTarjetaVCTemp & ""
               RsGrillaVConsumo.Requery

               '  If RsGrillaVConsumo.EOF Or RsGrillaVConsumo.BOF Then
               '      Visor IIf(txtResultado.Caption = "", "", txtResultado.Caption), "No encontrado", nPuerto, "N"
               '  Else
               '      Visor IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!tIdentidad), IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!descripcion), nPuerto, "N"
               '  End If
          Case Is = 4  ' Guardar
               wEnter = True
               lGuardarVC = True

               Dim sCOD_Vale_Temp2 As String
               sCOD_Vale_Temp2 = Calcular("SELECT TOP 1 tCOD_VALE AS Codigo FROM VALE_CONSUMO WHERE tCOD_VALE IN (SELECT tCOD_VALE FROM " & sTarjetaVCTemp & ") AND lESTADO_USO = 1 AND CONVERT(VARCHAR(20),fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) ", Cn)
               If sCOD_Vale_Temp2 <> "0" Then
                  MsgBox "El nro Vale " & sCOD_Vale_Temp2 & " ya se ha usado en otro comprobante.", vbInformation, sMensaje
                  Exit Sub
               End If

               Dim montoVale As String
               montoVale = Calcular("select SUM(isnull(nMONTO_VALE,0)) codigo from " & sTarjetaVCTemp, Cn)
               sDescrip = montoVale
               'sDescrip = "123.00"'sTemp
               Cn.Execute "delete from " & sTarjetaVC & ""
               Isql = "insert into " & sTarjetaVC & _
                         "( tIndex, tTICKET, tCOD_VALE, tMONEDA, nMONTO_VALE) " & _
                         " select tIndex, tTICKET, tCOD_VALE, tMONEDA, nMONTO_VALE from " & sTarjetaVCTemp
               Cn.Execute Isql
               Unload Me
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
        '  Case Is = 11 ' MESA247
        '       Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lCanal4=1", "lCanal5=1")))) & " Order by Descripcion"
        '       txtResultado.Caption = ""
        '       Set RsGrillaVConsumo = Lib.OpenRecordset(Isql, Cn)
        '       Set grdGrilla.DataSource = RsGrillaVConsumo
   End Select
End Sub

Private Sub Form_Activate()
   txtResultado.Caption = sTemp
   txtTempo.Text = txtResultado.Caption
End Sub

Private Sub Form_Load()
    '    If sModulo = "ADICION" Then
    '       frmMozoUsuario.ReseteaTimer
    '    End If
   
    '   If tipoBusqueda = "PAGO" Then
    '    chkSelec.Visible = True
    '   Else
    '    chkSelec.Visible = False
    '   End If

    Centrar Me
    wEnter = False
    sCodigo = ""
    sDescrip = ""
    'PedidoMesa = ""
    MatrizBotones 40, 10, 1, 1, cmdKey
    LlenaKey
    '    If nPuerto > 0 Then
    '       cmdOpcion(3).Visible = True
    '    End If

    '    If sTipo = "Infhotel" Then
    '       Set RsGrillaVConsumo = Lib.OpenRecordset(Isql, CnInfhotel)
    '    Else
    '       Set RsGrillaVConsumo = Lib.OpenRecordset(Isql, Cn)
    '    End If
    
    '    Set grdGrilla.DataSource = RsGrillaVConsumo
    '    Screen.MousePointer = vbDefault

    ' NUEVO -------------------

    sTarjetaVCTemp = dbTemporal(sCaja, 5, "tIndex", "nVarChar(3)", _
                                    "tTICKET", "nVarChar(50)", _
                                    "tCOD_VALE", "nVarChar(20)", _
                                    "tMONEDA", "nVarChar(10)", _
                                    "nMONTO_VALE", "Float")

    Call ConfGrilla(4, grdGrilla, "Ticket", 2, "tTICKET", 1500, 0, 0, "", _
                                    "Vale", 2, "tCOD_VALE", 1500, 0, 0, "", _
                                    "Moneda", 2, "tMONEDA", 1500, 0, 0, "", _
                                    "Monto Vale", 2, "nMONTO_VALE", 1500, 1, 0, "####,###,##0.00")

     Isql = "insert into " & sTarjetaVCTemp & _
               "( tIndex, tTICKET, tCOD_VALE, tMONEDA, nMONTO_VALE) " & _
               " select tIndex, tTICKET, tCOD_VALE, tMONEDA, nMONTO_VALE from " & sTarjetaVC
     Cn.Execute Isql

    Set RsGrillaVConsumo = Lib.OpenRecordset("select * from " & sTarjetaVCTemp, Cn)
    Set grdGrilla.DataSource = RsGrillaVConsumo
    Screen.MousePointer = vbDefault


    If lvisor Then
        Timer_Aceptarvisor.Interval = 1000
    Else
        Timer_Aceptarvisor.Interval = 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   sTipo = ""
   Set frmBusquedaValeConsumo = Nothing
   VSApForBusqueda = True
     txtTempo.SetFocus
End Sub

Private Sub grdGrilla_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
End Sub

Private Sub grdGrilla_DblClick()

    If RsGrillaVConsumo.RecordCount = 0 Then
      Exit Sub
   End If

      'origen de ventas
   If Me.tipoBusqueda = "OrigenVentas" Then
    sCodigo = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!CodOrigenVenta)
    sDescrip = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!descripcion)
   Else
    sCodigo = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!codigo)
    sDescrip = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!descripcion)
   End If

   EjecutarDocumentacion sCodigo, sDescrip
   
End Sub


'Public Function validacionmesa247(codigo As String) As Boolean
'   'If MESA247 Then
'    Isql = "select tCodigoPedido from MPEDIDO where tTurno<>'' and tCaja<>''  and tCodigoPedido= '" + codigo + "'"
'    Set RsPedido = Lib.OpenRecordset(Isql, Cn)
'         If RsPedido.RecordCount > 0 Then
'             MsgBox ("El Pedido ya ha sido Importado")
'             'Isql = "select tCodigoPedido as Codigo,isnull((select (rtrim(ltrim(tApellido))) + ' ' + (rtrim(ltrim(tNombre))) from TDELIVERY where tCodigoDelivery=tClienteDelivery),'') as Descripcion from Mpedido Where tTurno = '' AND tCaja = '' AND isnull(idPedidoMesa,'')<>''"
'             'Set RsGrillaVConsumo = Lib.OpenRecordset(Isql, Cn)
'             'Set grdGrilla.DataSource = RsGrillaVConsumo
'             validacionmesa247 = True
'         Else
'             validacionmesa247 = False
'         End If
''   Else
''    validacionmesa247 = False
''   End If
'End Function

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    Screen.MousePointer = vbHourglass
    If Right(RsGrillaVConsumo.Sort, 3) = "ASC" Then
       RsGrillaVConsumo.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    Else
       RsGrillaVConsumo.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    End If
    RsGrillaVConsumo.Requery
    Screen.MousePointer = vbDefault
    nPredeterm = ColIndex
    txtPredeterm.Caption = grdGrilla.Columns(nPredeterm).Caption
    txtTempo.SetFocus
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   txtTempo.SetFocus
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

Private Sub Form_Paint()
   txtPredeterm.Caption = grdGrilla.Columns(nPredeterm).Caption
End Sub

Private Sub Timer_Aceptarvisor_Timer()

    If Calcular("select top 1 estado as codigo from infovisor   where tcaja='" & sCaja & "'", Cn) = 3 Then
        cmdaceptarvisor.Visible = True
'    RESULTADOVISOR = True
'    OPERADORMENSAJE = 0
'    TimerVisor.Enabled = False
'    Unload Me
'
    Else
    cmdaceptarvisor.Visible = False
        If Calcular("select top 1 estado as codigo from infovisor   where tcaja='" & sCaja & "'", Cn) = 4 Then
'            RESULTADOVISOR = False
'            OPERADORMENSAJE = 1
'            TimerVisor.Enabled = False
'             Cn.Execute "update infovisor set estado=1  where tcaja='" & sCaja & "'"
            Timer_Aceptarvisor.Interval = 0
            cmdaceptarvisor.Visible = False
            MsgBox ("Cancelado por el Cliente!")
            
            Unload Me
        End If
  End If
  
End Sub

Private Sub txtresultado_Change()
'   Filtrar
'    If lvisor And tipoBusqueda = "VISOR" Then
'        txtResultado.Caption = Mid(Trim(txtResultado.Caption), 1, 11)
'        Cn.Execute "UPDATE  infovisor  SET  rucfactura='" & txtResultado.Caption & "' ,tipodocu='" & sTipoDocumBusq & "'  where tcaja='" & sCaja & "'"
'    End If
End Sub

'Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
Private Sub txtTempo_KeyUp(KeyCode As Integer, Shift As Integer)
    txtResultado.Caption = txtTempo.Text

'    If sModulo = "ADICION" Then
'       frmMozoUsuario.ReseteaTimer
'    End If
'    If Shift > 0 Then
'       Select Case KeyCode
'        Case 48
'             txtResultado.Caption = txtResultado.Caption + "="
'        Case 49
'             txtResultado.Caption = txtResultado.Caption + "!"
'        Case 50
'             txtResultado.Caption = txtResultado.Caption + "_"
'        Case 51
'             txtResultado.Caption = txtResultado.Caption + "#"
'        Case 52
'             txtResultado.Caption = txtResultado.Caption + "$"
'        Case 53
'             txtResultado.Caption = txtResultado.Caption + "%"
'        Case 54
'             txtResultado.Caption = txtResultado.Caption + "&"
'        Case 55
'             txtResultado.Caption = txtResultado.Caption + "/"
'        Case 56
'             txtResultado.Caption = txtResultado.Caption + "("
'        Case 57
'             txtResultado.Caption = txtResultado.Caption + ")"
'        'Case Else
'        '     MsgBox KeyCode
'       End Select
'    Else
      Select Case KeyCode
          Case 13 ' ENTER
               'Call cmdkey_Click(43)
               Call cmdOpcion_Click(1)
'           Case 27
'                Call cmdkey_Click(41)
'           Case 8
'                Call cmdkey_Click(42)
'           Case 32
'                txtResultado.Caption = txtResultado.Caption + " "
'           Case 65
'                txtResultado.Caption = txtResultado.Caption + "A"
'           Case 66
'                txtResultado.Caption = txtResultado.Caption + "B"
'           Case 67
'                txtResultado.Caption = txtResultado.Caption + "C"
'           Case 68
'                txtResultado.Caption = txtResultado.Caption + "D"
'           Case 69
'                txtResultado.Caption = txtResultado.Caption + "E"
'           Case 70
'                txtResultado.Caption = txtResultado.Caption + "F"
'           Case 71
'                txtResultado.Caption = txtResultado.Caption + "G"
'           Case 72
'                txtResultado.Caption = txtResultado.Caption + "H"
'           Case 73
'                txtResultado.Caption = txtResultado.Caption + "I"
'           Case 74
'                txtResultado.Caption = txtResultado.Caption + "J"
'           Case 75
'                txtResultado.Caption = txtResultado.Caption + "K"
'           Case 76
'                txtResultado.Caption = txtResultado.Caption + "L"
'           Case 77
'                txtResultado.Caption = txtResultado.Caption + "M"
'           Case 78
'                txtResultado.Caption = txtResultado.Caption + "N"
'           Case 79
'                txtResultado.Caption = txtResultado.Caption + "O"
'           Case 80
'                txtResultado.Caption = txtResultado.Caption + "P"
'           Case 81
'                txtResultado.Caption = txtResultado.Caption + "Q"
'           Case 82
'                txtResultado.Caption = txtResultado.Caption + "R"
'           Case 83
'                txtResultado.Caption = txtResultado.Caption + "S"
'           Case 84
'                txtResultado.Caption = txtResultado.Caption + "T"
'           Case 85
'                txtResultado.Caption = txtResultado.Caption + "U"
'           Case 86
'                txtResultado.Caption = txtResultado.Caption + "V"
'           Case 87
'                txtResultado.Caption = txtResultado.Caption + "W"
'           Case 88
'               txtResultado.Caption = txtResultado.Caption + "X"
'           Case 89
'                txtResultado.Caption = txtResultado.Caption + "Y"
'           Case 90
'                txtResultado.Caption = txtResultado.Caption + "Z"
'           Case 96, 48
'                txtResultado.Caption = txtResultado.Caption + "0"
'           Case 97, 49
'                txtResultado.Caption = txtResultado.Caption + "1"
'           Case 98, 50
'                txtResultado.Caption = txtResultado.Caption + "2"
'           Case 99, 51
'                txtResultado.Caption = txtResultado.Caption + "3"
'           Case 100, 52
'                txtResultado.Caption = txtResultado.Caption + "4"
'           Case 101, 53
'                txtResultado.Caption = txtResultado.Caption + "5"
'           Case 102, 54
'                txtResultado.Caption = txtResultado.Caption + "6"
'           Case 103, 55
'                txtResultado.Caption = txtResultado.Caption + "7"
'           Case 104, 56
'                txtResultado.Caption = txtResultado.Caption + "8"
'           Case 105, 57
'                txtResultado.Caption = txtResultado.Caption + "9"
'           Case 110, 190
'                txtResultado.Caption = txtResultado.Caption + "."
'           Case 106, 188
'                txtResultado.Caption = txtResultado.Caption + ","
'           Case 109, 189
'                txtResultado.Caption = txtResultado.Caption + "-"
'           Case 192
'                txtResultado.Caption = txtResultado.Caption + "Ñ"
' '          Case Else
' '               MsgBox KeyCode
      End Select
'     End If
End Sub

Public Sub Filtrar()
On Error Resume Next
   If Len(Trim(txtResultado.Caption)) > 0 Then
      Select Case VarType(grdGrilla.Columns(nPredeterm).value)
             Case 2 To 6
            ' MsgBox Val(txtResultado)
             sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & "= " & val(txtResultado.Caption)
             Case 7
             sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & " >= #" & txtResultado.Caption & "# and " & Trim(grdGrilla.Columns(nPredeterm).DataField) & " <= #" & txtResultado.Caption & " 23:59#"
            Case 1 ' es nulo
                GoTo nulo
             Case Else
             sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & " Like " & "'*" & txtResultado.Caption & "*'"
      End Select
      Screen.MousePointer = vbHourglass
      With RsGrillaVConsumo
           .Filter = sCriterio
           If .BOF = False Then
              .MoveFirst
           End If
      End With
      Screen.MousePointer = vbDefault
      Exit Sub
      'cmdTexto.Caption = "Total : " & Format(RsGrillaVConsumo.RecordCount, "##,##0") & " de " & Format(nTotal, "##,##0") & " Registros"
nulo:
        sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & "= " & val(txtResultado.Caption)
        RsGrillaVConsumo.Filter = sCriterio
        If RsGrillaVConsumo.BOF = True Then
             sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & " Like " & "'*" & txtResultado.Caption & "*'"
             RsGrillaVConsumo.Filter = sCriterio
             If RsGrillaVConsumo.BOF = True Then
                 sCriterio = Trim(grdGrilla.Columns(nPredeterm).DataField) & " >= #" & txtResultado.Caption & "# and " & Trim(grdGrilla.Columns(nPredeterm).DataField) & " <= #" & txtResultado.Caption & " 23:59#"
                 RsGrillaVConsumo.Filter = sCriterio
             End If
        End If
      Screen.MousePointer = vbDefault
      Exit Sub
   Else
      Screen.MousePointer = vbHourglass
      RsGrillaVConsumo.Filter = adFilterNone
      RsGrillaVConsumo.Requery
      If RsGrillaVConsumo.RecordCount <> 0 Then
         RsGrillaVConsumo.MoveFirst
      End If
      'cmdTexto.Caption = "Total : " & Format(nTotal, "##,##0") & " de " & Format(nTotal, "##,##0") & " Registros"
      Screen.MousePointer = vbDefault
      sCriterio = ""
   End If
End Sub






''---------VALIDA EMAIL------
'Public Sub ObtenerTipoDocumento(ByVal xTipoDoc As Variant)
'   xTipoDocumento = xTipoDoc
'End Sub
''---------------------------


Public Sub EjecutarDocumentacion(ByVal sCodigo As String, ByVal sDescrip As String)

    wEnter = True

    If xbermur = "SI" Then
        sMontoBermur = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!MontoDisponible)
    End If

    If tipoBusqueda = "BusquedaCajaRapida" Then
       tenlacebusqueda = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!tEnlace)
        tipoBusqueda = ""
    Else
        If tipoBusqueda = "BusquedaVenta" Then
            tenlacebusquedaVenta = IIf(RsGrillaVConsumo.EOF = True, "", RsGrillaVConsumo!tEnlace)
            tipoBusqueda = ""
        End If
    End If
    
    Set RsDatosClienteFac = Lib.OpenRecordset("select  isnull(AgenteRetencion,0)as AgenteRetencion,tidentidad from vcliente where Codigo='" & IIf(IsNull(sCodigo), "", sCodigo) & "'", Cn)
    
'    If Not (RsDatosClienteFac.EOF Or RsDatosClienteFac.BOF) Then
'         If sModulo <> "ADICION" Then
'         '9016
'
'          If UCase(xtipodoc) = "F" Then
'            If (RsDatosClienteFac!AgenteRetencion = 0) And xTotal >= val(tMontoRetencion) Then
'                'GCAA 26052021
'                Dim xxDat As String
'                Dim xxWml As New XMLHTTP
'                xxWml.Open "GET", RutaConsultaRucRetenciones & Trim(RsDatosClienteFac!tIdentidad), False
'                xxWml.send
'                If xxWml.Status = 200 Then
'                    xxDat = Trim(xxWml.responseText)
'                    xxDat = Replace(Replace(Replace(xxDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
'                    xxDat = Replace(xxDat, Chr(34), "")
'                    Dim xMATRIZ() As String
'                    xMATRIZ = Split(xxDat, vbNewLine)
'                    xMATRIZ(0) = Replace(xMATRIZ(3), "status:", "")
'                    If (xMATRIZ(0) <> "Ruc:null") Then
'                        Cn.Execute "update TCLIENTE set AgenteRetencion=1  where tcodigocliente='" & sCodigo & "'"
'                    Else
'                         Cn.Execute "update TCLIENTE set AgenteRetencion=0  where tcodigocliente='" & sCodigo & "'"
'                    End If
'                ElseIf xxWml.Status = 500 Then
'                    MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
'                ElseIf xxWml.Status = 12007 Then
'                    MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
'                Else
'                    ' MsgBox "El documento de referencia fue emitida fuera del mes, Usuario sin permiso!!!", vbExclamation, sMensaje
'                    If lActivarAvisoRetenciones = True Then
'                        Set RsDatosClienteFac = Lib.OpenRecordset("select  isnull(AgenteRetencion,0)as AgenteRetencion, tidentidad from vcliente where Codigo='" & IIf(IsNull(sCodigo), "", sCodigo) & "'", Cn)
'                        If (RsDatosClienteFac!AgenteRetencion = 0) And xTotal >= val(tMontoRetencion) Then
'                          '  MsgBox "Verificar si el cliente es Agente de retención, de ser asi, modificar en cliente que es agente de retención!!!", vbExclamation, sMensaje
'                            If MsgBox("Verificar si el cliente es Agente de Retención?, Activar Check de Agente de Retención - Modificar Cliente!", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'                                  Exit Sub
'                            Else
'
'                            End If
'                        End If
'                    End If
''                    MsgBox "Verificar conexion a internet!!!", vbExclamation, sMensaje
'                End If
'            End If
'         End If
'
'            Set xxWml = Nothing
'
'        End If
'    End If

   Unload Me
  VSApForBusqueda = False
fin:
End Sub







