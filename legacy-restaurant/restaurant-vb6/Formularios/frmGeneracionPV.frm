VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmGeneracionPV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Documentos"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   Icon            =   "frmGeneracionPV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   " Recargos "
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
      Height          =   1815
      Left            =   6000
      TabIndex        =   27
      Top             =   6330
      Width           =   4335
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Autorización Recargo"
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
         Index           =   16
         Left            =   180
         TabIndex        =   34
         Top             =   330
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "- 5 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   2940
         TabIndex        =   32
         Top             =   1050
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "+ 5 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   1545
         TabIndex        =   31
         Top             =   1050
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "- 1 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   2940
         TabIndex        =   30
         Top             =   330
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Importe del Recargo"
         Height          =   615
         Index           =   5
         Left            =   180
         TabIndex        =   29
         Top             =   1050
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "+ 1 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   1560
         TabIndex        =   28
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   " Descuentos "
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
      Height          =   1815
      Left            =   6000
      TabIndex        =   21
      Top             =   4440
      Width           =   4335
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Autorización Descuento"
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
         Index           =   14
         Left            =   150
         TabIndex        =   33
         Top             =   330
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "- 5 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   2940
         TabIndex        =   26
         Top             =   1050
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "+ 5 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   1545
         TabIndex        =   25
         Top             =   1050
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "- 1 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2940
         TabIndex        =   24
         Top             =   330
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "+ 1 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1545
         TabIndex        =   23
         Top             =   330
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Importe del Descuento"
         Height          =   615
         Index           =   0
         Left            =   150
         TabIndex        =   22
         Top             =   1050
         Width           =   1275
      End
   End
   Begin VB.Frame fraTipoDocumento 
      Caption         =   " Tipo de Documento "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   60
      TabIndex        =   19
      Top             =   4440
      Width           =   5865
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   0
         Left            =   210
         TabIndex        =   20
         Top             =   330
         Width           =   1320
      End
   End
   Begin VB.Frame fraDato 
      Caption         =   " Datos del Documento "
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
      Height          =   1035
      Left            =   60
      TabIndex        =   6
      Top             =   3360
      Width           =   11625
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Descuento :"
         Height          =   195
         Index           =   2
         Left            =   2370
         TabIndex        =   18
         Top             =   675
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Recargo :"
         Height          =   195
         Index           =   4
         Left            =   5790
         TabIndex        =   17
         Top             =   675
         Width           =   1515
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto Descuento :"
         Height          =   195
         Index           =   1
         Left            =   2685
         TabIndex        =   16
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto Recargo :"
         Height          =   195
         Index           =   3
         Left            =   6105
         TabIndex        =   15
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Neto :"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   315
         Width           =   435
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Index           =   5
         Left            =   9270
         TabIndex        =   13
         Top             =   315
         Width           =   450
      End
      Begin VB.Label txtNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   750
         TabIndex        =   12
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label txtDImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4125
         TabIndex        =   11
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label txtRImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7380
         TabIndex        =   10
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label txtDPorcentaje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4125
         TabIndex        =   9
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label txtRPorcentaje 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7380
         TabIndex        =   8
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9795
         TabIndex        =   7
         Top             =   270
         Width           =   1545
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
      Index           =   15
      Left            =   10470
      Picture         =   "frmGeneracionPV.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7545
      Width           =   1275
   End
   Begin VB.CommandButton cmdEmitir 
      Caption         =   "Emitir"
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
      Left            =   10470
      Picture         =   "frmGeneracionPV.frx":047C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6930
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   615
      Index           =   13
      Left            =   10470
      Picture         =   "frmGeneracionPV.frx":057E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6315
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   615
      Index           =   12
      Left            =   10470
      Picture         =   "frmGeneracionPV.frx":0E48
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5700
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   615
      Index           =   11
      Left            =   10470
      Picture         =   "frmGeneracionPV.frx":1712
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5085
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   615
      Index           =   10
      Left            =   10470
      Picture         =   "frmGeneracionPV.frx":1FDC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4470
      Width           =   1275
   End
   Begin TrueOleDBGrid80.TDBGrid grdDocumento 
      Height          =   3300
      Left            =   30
      TabIndex        =   35
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   5821
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
      Splits(0).ScrollBars=   0
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
      Caption         =   "Documentos por Generar"
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
Attribute VB_Name = "frmGeneracionPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nDescuento As Double
Dim nRecargo As Double
Dim nCantidad As Integer
Dim wCortesia As Double

'Load
Dim RsDocumento As Recordset
Dim RsTotales As Recordset
Dim RsTipoDocumento As Recordset
Dim RsImpresion As Recordset
Dim RsDetalle As Recordset
Dim RsSuma As Recordset
Dim RsT3 As Recordset

Dim sCabecera As String
Dim sDetalle As String
Dim mTipo() As String
Dim mConsumo() As String
Dim nPos As Integer
Dim nNeto As Double
Dim nPrecioOficial As Double
Dim nTotal As Double
Dim xImpuesto1 As Double
Dim xImpuesto2 As Double
Dim xImpuesto3 As Double

Private Sub cmdEmitir_Click()
    grdDocumento.MoveFirst
        
    Dim i As Integer

   'Chequea Consistencia
   For i = 1 To nCantidad
      RsTipoDocumento.MoveFirst
      RsTipoDocumento.Find ("tTipoEmision='" & RsDocumento!tTipodocumento & "'")
      
      If Not RsTipoDocumento.EOF Then
      
      Else
         MsgBox "Error: Documento sin definición", vbCritical, sMensaje
         Exit Sub
      End If
      
      grdDocumento.MoveNext
   Next
        
   RsDocumento.MoveFirst
   For i = 1 To nCantidad
       xCantidad = i
       Facturar RsTipoDocumento, RsDocumento!tTipodocumento, "Division"
   Next i
   Cn.Execute "delete from " & sTemporal
   Screen.MousePointer = vbDefault
   Unload Me

End Sub

Private Sub Form_Load()
   Dim nEmision As Integer
   Dim i As Integer
   Dim col As Integer
   Dim margentop As Integer
   Dim margenleft As Integer
   
   nDescuento = 0
   nRecargo = 0
   Desactiva
   
   'Llena la Cabecera del Documento
   sCabecera = dbTemporal(sCaja, 25, "Item", "Int", _
                                     "tDocumento", "nVarChar(15)", _
                                     "tTipoDocumento", "nVarChar(2)", _
                                     "tTipoPedido", "nVarChar(2)", _
                                     "fFecha", "smalldatetime", _
                                     "tCodigoCliente", "nVarChar(5)", _
                                     "tMoneda", "nVarChar(3)", _
                                     "nNeto", "Float", _
                                     "nRecargo", "Float", _
                                     "nDescuento", "Float", _
                                     "nPrecioOficial", "Float", _
                                     "nPrecioImpuesto1", "Float", "nPrecioImpuesto2", "Float", "nPrecioImpuesto3", "Float", _
                                     "nVenta", "Float", _
                                     "nPropina", "Float", _
                                     "nTotal", "Float", _
                                     "tCortesia", "nVarChar(4)", _
                                     "tEstadoDocumento", "nVarChar(2)", _
                                     "tTurno", "nVarChar(10)", _
                                     "tCaja", "nVarChar(3)", _
                                     "tObservacion", "nVarChar(250)", _
                                     "tUsuario", "nVarChar(15)", _
                                     "fRegistro", "smalldatetime", _
                                     "tConsumo", "nVarChar(250)")
   
   nCantidad = Val(sTemp)
   
   Isql = "select sum(nVenta - nImpuesto1 - nImpuesto2 - nImpuesto3) as nPrecioNeto, sum(nImpuesto1) as nImpuesto1, sum(nImpuesto2) as nImpuesto2, " & _
          "sum(nImpuesto3) as nImpuesto3, sum(nVenta) as nVenta from " & sTemporal & " where Estado='X'"
          
   Set RsTotales = Lib.OpenRecordset(Isql, Cn)
        
   For i = 1 To nCantidad
       Isql = "insert into " & sCabecera & _
              "( item, fFecha, nNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nVenta) " & _
              " values ( " & i & ", " _
                           & "getdate(), " _
                           & RsTotales!nPrecioNeto / nCantidad & ", " _
                           & "0 , 0, " _
                           & RsTotales!nVenta / nCantidad & ", " _
                           & RsTotales!nImpuesto1 / nCantidad & ", " _
                           & RsTotales!nImpuesto2 / nCantidad & ", " _
                           & RsTotales!nImpuesto3 / nCantidad & ", " _
                           & RsTotales!nVenta / nCantidad & ") "
       Cn.Execute Isql
   Next
 
   Isql = "SELECT [" & sCabecera & "].tDocumento, [" & sCabecera & "].Item, [" & sCabecera & "].tCodigoCliente, TCLIENTE.tEmpresa AS Cliente, [" & sCabecera & "].tMoneda, [" & sCabecera & "].tEstadoDocumento, TCLIENTE.tIdentidad AS RUC, TCLIENTE.tDireccion AS Direccion, [" & sCabecera & "].nNeto, [" & sCabecera & "].nPrecioImpuesto1, [" & sCabecera & "].nPrecioImpuesto2, [" & sCabecera & "].fFecha, [" & sCabecera & "].nPrecioImpuesto3, [" & sCabecera & "].nPrecioOficial, [" & sCabecera & "].nRecargo, [" & sCabecera & "].nDescuento, [" & sCabecera & "].nVenta, [" & sCabecera & "].tTipoDocumento, [" & sCabecera & "].tUsuario, vTipoDocumento.Descripcion AS Documento, vCortesia.Descripcion AS Cortesia, [" & sCabecera & "].tConsumo " & _
          "FROM ((TCLIENTE RIGHT JOIN [" & sCabecera & "] ON TCLIENTE.tCodigoCliente = [" & sCabecera & "].tCodigoCliente) LEFT JOIN vTipoDocumento ON [" & sCabecera & "].tTipoDocumento = vTipoDocumento.Codigo) LEFT JOIN vCortesia ON [" & sCabecera & "].tCortesia = vCortesia.Codigo"
   
   Set RsDocumento = Lib.OpenRecordset(Isql, Cn)
   Call ConfGrilla(11, grdDocumento, "No", 2, "Item", 350, 1, 0, "###", _
                                     "Documento", 2, "Documento", 1100, 0, 0, "", _
                                     "Cliente", 2, "Cliente", 2250, 0, 0, "", _
                                     "Neto", 2, "nNeto", 1000, 1, 0, "####,###,##0.00", _
                                     sImpuesto1, 2, "nPrecioImpuesto1", 850, 1, 0, "####,###,##0.00", _
                                     sImpuesto2, 2, "nPrecioImpuesto2", 850, 1, 0, "####,###,##0.00", _
                                     sImpuesto3, 2, "nPrecioImpuesto3", 850, 1, 0, "####,###,##0.00", _
                                     "Total", 2, "nVenta", 1150, 1, 0, "####,###,##0.00", _
                                     "Recargo", 2, "nRecargo", 850, 1, 0, "####,###,##0.00", _
                                     "Descto", 2, "nDescuento", 850, 1, 0, "####,###,##0.00", _
                                     "P.Oficial", 2, "nPrecioOficial", 1150, 1, 0, "####,###,##0.00")

   
   Set grdDocumento.DataSource = RsDocumento
   
   AsignaTotales
      
   'Emision x Impresora (Documentos) sin "Sin Documento"
   Isql = "select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "'"
   Set RsTipoDocumento = Lib.OpenRecordset(Isql, Cn)
   nEmision = RsTipoDocumento.RecordCount
   If nEmision = 0 Then
      MsgBox "No se ha ingresado los Documentos por Caja", vbCritical, "Error Fatal"
      Unload Me
   End If
   
   ReDim mTipo(nEmision, 4)
       
   cmdTipoDocumento(0).Visible = False
   RsTipoDocumento.MoveFirst
    
   col = 0
   margentop = cmdTipoDocumento(0).Top
   margenleft = cmdTipoDocumento(0).Left
   wCortesia = False
   For i = 1 To nEmision
       If RsTipoDocumento!tTipoEmision = "00" Then
          wCortesia = True
       End If
       Load cmdTipoDocumento(i)
       If RsTipoDocumento!tTipoEmision <> "00" Then
          col = col + 1
          cmdTipoDocumento(i).Caption = RsTipoDocumento!Descripcion
          cmdTipoDocumento(i).Left = (col * 50) + ((col - 1) * cmdTipoDocumento(0).Width) + margenleft
          cmdTipoDocumento(i).Top = margentop
          cmdTipoDocumento(i).Visible = True
          
          mTipo(i, 1) = RsTipoDocumento!tTipoEmision
          mTipo(i, 2) = RsTipoDocumento!Prefijo
          mTipo(i, 3) = RsTipoDocumento!Cliente
          mTipo(i, 4) = RsTipoDocumento!tImpresora
       Else
          cmdTipoDocumento(i).Visible = False
       End If
       
       'Cantidad de Columnas
       If i Mod 4 = 0 Then
          col = 0
          margentop = margentop + cmdTipoDocumento(0).Height + 50
       End If
       RsTipoDocumento.MoveNext
   Next i
   ReDim mConsumo(nCantidad, 2)
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Dim nPos As Integer
   Dim Acumulado As Double
   Select Case Index
          Case Is = 0  ' Importe Dscto
               If Val(txtNeto.Caption) > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nPos = grdDocumento.Row
                     nDescuento = Val(sDescrip)
                     
                     nTotal = nPrecioOficial + nRecargo - nDescuento
                     Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
                     Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
                     Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
                     Acumulado = 1 + (Acumulado / 100)
                     
                     xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
                     xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
                     xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)
                                                                                       
                     Isql = "Update " & sCabecera & " set nDescuento=" & nDescuento & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
                                                    
                     Cn.Execute Isql
                     RsDocumento.Requery
                     AsignaTotales
                     grdDocumento.Row = nPos
                  End If
               End If
               
          Case Is = 1  ' +1
               nPos = grdDocumento.Row
               txtDPorcentaje.Caption = Format(IIf(Val(txtDPorcentaje.Caption) >= 100, 100, Val(txtDPorcentaje.Caption) + 1), "##0.00")
               
               nDescuento = Val(txtDPorcentaje.Caption) * nPrecioOficial / 100
               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)
               
               Isql = "Update " & sCabecera & " set nDescuento=" & nDescuento & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos
               
          Case Is = 2  ' -1
               nPos = grdDocumento.Row
               txtDPorcentaje.Caption = Format(IIf(Val(txtDPorcentaje.Caption) <= 0, 0, Val(txtDPorcentaje.Caption) - 1), "##0.00")
               nDescuento = Val(txtDPorcentaje.Caption) * nPrecioOficial / 100

               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)

               Isql = "Update " & sCabecera & " set nDescuento=" & nDescuento & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos

          Case Is = 3  ' +5
               nPos = grdDocumento.Row
               txtDPorcentaje.Caption = Format(IIf(Val(txtDPorcentaje.Caption) >= 95, 100, Val(txtDPorcentaje.Caption) + 5), "##0.00")
               nDescuento = Val(txtDPorcentaje.Caption) * nPrecioOficial / 100

               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)

               Isql = "Update " & sCabecera & " set nDescuento=" & nDescuento & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos

          Case Is = 4  ' -5
               nPos = grdDocumento.Row
               txtDPorcentaje.Caption = Format(IIf(Val(txtDPorcentaje.Caption) <= 5, 0, Val(txtDPorcentaje.Caption) - 5), "##0.00")
               nDescuento = Val(txtDPorcentaje.Caption) * nPrecioOficial / 100

               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)

               Isql = "Update " & sCabecera & " set nDescuento=" & nDescuento & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos
          
          Case Is = 5  ' Importe Recargo
               If Val(txtNeto.Caption) > 0 Then
                  sTipo = ""
                  frmNumPad.Show vbModal
                  If wEnter Then
                     nPos = grdDocumento.Row
                     nRecargo = Val(sDescrip)
                     
                     nTotal = nPrecioOficial + nRecargo - nDescuento
                     Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
                     Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
                     Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
                     Acumulado = 1 + (Acumulado / 100)
                     
                     xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
                     xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
                     xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)
                                                                                                            
                     Isql = "Update " & sCabecera & " set nRecargo=" & nRecargo & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
                     Cn.Execute Isql
                     RsDocumento.Requery
                     AsignaTotales
                     grdDocumento.Row = nPos
                  End If
               End If
               
          Case Is = 6  ' +1
               nPos = grdDocumento.Row
               txtRPorcentaje.Caption = Format(IIf(Val(txtDPorcentaje.Caption) >= 100, 100, Val(txtRPorcentaje.Caption) + 1), "##0.00")
               nRecargo = Val(txtRPorcentaje.Caption) * nPrecioOficial / 100

               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)

               Isql = "Update " & sCabecera & " set nRecargo=" & nRecargo & ", " & _
                                               "nPrecioOficial = " & nPrecioOficial & ", " & _
                                               "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                               "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                               "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                               "nVenta = " & nPrecioOficial + xImpuesto1 + xImpuesto2 + xImpuesto3 & " " & _
                                               "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos
          
          Case Is = 7  ' -1
               nPos = grdDocumento.Row
               txtRPorcentaje.Caption = Format(IIf(Val(txtRPorcentaje.Caption) <= 0, 0, Val(txtRPorcentaje.Caption) - 1), "##0.00")
               nRecargo = Val(txtRPorcentaje.Caption) * nPrecioOficial / 100

               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)

               Isql = "Update " & sCabecera & " set nRecargo=" & nRecargo & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos

          Case Is = 8  ' +5
               nPos = grdDocumento.Row
               txtRPorcentaje.Caption = Format(IIf(Val(txtRPorcentaje.Caption) >= 95, 100, Val(txtRPorcentaje.Caption) + 5), "##0.00")
               nRecargo = Val(txtRPorcentaje.Caption) * nPrecioOficial / 100

               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)

               Isql = "Update " & sCabecera & " set nRecargo=" & nRecargo & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos

          Case Is = 9  ' -5
               nPos = grdDocumento.Row
               txtRPorcentaje.Caption = Format(IIf(Val(txtRPorcentaje.Caption) <= 5, 0, Val(txtRPorcentaje.Caption) - 5), "##0.00")
               nRecargo = Val(txtRPorcentaje.Caption) * nPrecioOficial / 100

               nTotal = nPrecioOficial + nRecargo - nDescuento

               Acumulado = IIf(grdDocumento.Columns(4).Text > 0, Acumulado + nPorcentaje1, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(5).Text > 0, Acumulado + nPorcentaje2, Acumulado)
               Acumulado = IIf(grdDocumento.Columns(6).Text > 0, Acumulado + nPorcentaje3, Acumulado)
               Acumulado = 1 + (Acumulado / 100)
            
               xImpuesto1 = IIf(grdDocumento.Columns(4).Text > 0, nTotal / Acumulado * nPorcentaje1 / 100, 0)
               xImpuesto2 = IIf(grdDocumento.Columns(5).Text > 0, nTotal / Acumulado * nPorcentaje2 / 100, 0)
               xImpuesto3 = IIf(grdDocumento.Columns(6).Text > 0, nTotal / Acumulado * nPorcentaje3 / 100, 0)

               Isql = "Update " & sCabecera & " set nRecargo=" & nRecargo & ", " & _
                                                    "nNeto = " & nTotal - xImpuesto1 - xImpuesto2 - xImpuesto3 & ", " & _
                                                    "nPrecioImpuesto1 = " & xImpuesto1 & ", " & _
                                                    "nPrecioImpuesto2 = " & xImpuesto2 & ", " & _
                                                    "nPrecioImpuesto3 = " & xImpuesto3 & ", " & _
                                                    "nVenta = " & nTotal & " " & _
                                                    "where item = " & grdDocumento.Row + 1
               Cn.Execute Isql
               RsDocumento.Requery
               AsignaTotales
               grdDocumento.Row = nPos
         
          Case Is = 10 ' Primero
               MoverPuntero Primero, grdDocumento
               AsignaDetalle
               AsignaTotales

          Case Is = 11 ' Previo
               MoverPuntero previo, grdDocumento
               AsignaDetalle
               AsignaTotales

          Case Is = 12 ' Siguiente
               MoverPuntero siguiente, grdDocumento
               AsignaDetalle
               AsignaTotales

          Case Is = 13 ' Ultimo
               MoverPuntero Ultimo, grdDocumento
               AsignaDetalle
               AsignaTotales
               
          Case Is = 14 ' Activa Descuento
               If Supervisor = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               Else
                  cmdOpcion(0).Enabled = True
                  cmdOpcion(1).Enabled = True
                  cmdOpcion(2).Enabled = True
                  cmdOpcion(3).Enabled = True
                  cmdOpcion(4).Enabled = True
               End If
                         
          Case Is = 15 ' Salir
               Unload Me

          Case Is = 16 ' Activa Recargo
               If Supervisor = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               Else
                  cmdOpcion(5).Enabled = True
                  cmdOpcion(6).Enabled = True
                  cmdOpcion(7).Enabled = True
                  cmdOpcion(8).Enabled = True
                  cmdOpcion(9).Enabled = True
               End If

          Case Is = 17 ' Cliente
'               sCodigo = ""
'               sDescrip = ""
'               sTipo = "Cliente"
'
'               sTemp = Calcular("select tIdentidad as codigo from TCOMPANIA where tCodigoCliente ='" & sCliente & "'", Cn)
'               Isql = "SELECT * from vCliente where tIdentidad ='" & sTemp & "' and lActivo = 1 Order by Descripcion"
'
'
''               Isql = "select * from vCliente where lActivo=1 order by Descripcion"
'
'               frmBusqueda.Show vbModal
'
'               If wEnter = True Then
'                  nPos = grdDocumento.Row
'                  Isql = "Update " & sCabecera & " set tCodigoCliente='" & sCodigo & "' where item = " & grdDocumento.Row + 1
'                  Cn.Execute Isql
'                  txtRazonS.Caption = sDescrip
'                  RsDocumento.Requery
'
'                  If nCantidad - 1 = nPos Then
'                     grdDocumento.MoveLast
'                  Else
'                     grdDocumento.Row = nPos
'                  End If
'                  AsignaDetalle
'               Else
'                  txtRazonS.Caption = ""
'               End If
               
          Case Is = 18 ' Cortesia
               If Not wCortesia Then
                  MsgBox "Caja sin Cortesía", vbCritical, sMensaje
                  Exit Sub
               End If
          
               frmPassword.Show vbModal
               If wEnter Then
                  If sDescrip <> sPAdmin Then
                     MsgBox "Password Incorrecto", vbCritical, sMensaje
                     Exit Sub
                  End If
               Else
                  Exit Sub
               End If

               sTipo = ""
               Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
               frmBusqueda.Show vbModal
               
               If wEnter = True Then
                  'txtCortesia.Caption = sDescrip
                  nPos = grdDocumento.Row
                  Isql = "Update " & sCabecera & " set tTipoDocumento='00', " & _
                         "tCodigoCliente='', " & _
                         "tCortesia= '" & sCodigo & "' " & _
                         " where item = " & grdDocumento.Row + 1

                  Cn.Execute Isql
                  RsDocumento.Requery
                   
                  If nCantidad - 1 = nPos Then
                     grdDocumento.MoveLast
                  Else
                     grdDocumento.Row = nPos
                  End If
                  AsignaDetalle
               Else
                  'txtCortesia.Caption = ""
               End If
          
   End Select
End Sub

Public Sub cmdTipoDocumento_Click(Index As Integer)
   'Dim nRespuesta As Integer
   'Dim xConsumo As String
   RsTipoDocumento.MoveFirst
   RsTipoDocumento.Move (Index - 1)
   
   'xConsumo = ""
   'nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
   'If nRespuesta = vbYes Then
   '   frmKeyBoard.txtResultado = "POR CONSUMO"
   '   frmKeyBoard.Show vbModal
   '   If sDescrip = "" Or Not wEnter Then
   '      MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
   '      Exit Sub
   '   End If
   '   xConsumo = sDescrip
   'ElseIf nRespuesta = vbCancel Then
   '   Exit Sub
   'End If
      
   nPos = grdDocumento.Row
   Isql = "Update " & sCabecera & " set tTipoDocumento='" & mTipo(Index, 1) & "', tCodigoCliente='', tCortesia='' where item = " & grdDocumento.Row + 1
   Cn.Execute Isql
   RsDocumento.Requery
   If nCantidad - 1 = nPos Then
      grdDocumento.MoveLast
   Else
      grdDocumento.Row = nPos
   End If
   AsignaDetalle
End Sub

Public Sub AsignaTotales()
   nNeto = IIf(IsNull(RsDocumento!nNeto), 0, RsDocumento!nNeto)
   nPrecioOficial = IIf(IsNull(RsDocumento!nPrecioOficial), 0, RsDocumento!nPrecioOficial)
   nDescuento = IIf(IsNull(RsDocumento!nDescuento), "0", RsDocumento!nDescuento)
   nRecargo = IIf(IsNull(RsDocumento!nRecargo), "0", RsDocumento!nRecargo)
   nTotal = IIf(IsNull(RsDocumento!nVenta), "0", RsDocumento!nVenta)
   txtNeto.Caption = Format(nNeto, "###,###,##0.00")
   txtRImporte.Caption = Format(nRecargo, "###,###,##0.00")
   txtDImporte.Caption = Format(nDescuento, "###,###,##0.00")
   txtTotal.Caption = Format(nTotal, "###,###,##0.00")
   
   txtDPorcentaje.Caption = Format(nDescuento * 100 / IIf(nPrecioOficial = 0, 1, nPrecioOficial), "###,###,###,##0.00")
   txtRPorcentaje.Caption = Format(nRecargo * 100 / IIf(nPrecioOficial = 0, 1, nPrecioOficial), "###,###,###,##0.00")
   
End Sub

Public Sub AsignaDetalle()
   'txtRazonS.Caption = IIf(IsNull(RsDocumento!Cliente), "", RsDocumento!Cliente)
   'txtCortesia.Caption = IIf(IsNull(RsDocumento!Cortesia), "", RsDocumento!Cortesia)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sCabecera
   
   Set RsDocumento = Nothing
   Set RsTotales = Nothing
   Set RsTipoDocumento = Nothing
   Set RsDetalle = Nothing
   Set RsSuma = Nothing
   Set frmGeneracionPV = Nothing
End Sub

Private Sub grdDocumento_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   AsignaDetalle
   AsignaTotales
End Sub

Public Sub Desactiva()
   cmdOpcion(0).Enabled = False
   cmdOpcion(1).Enabled = False
   cmdOpcion(2).Enabled = False
   cmdOpcion(3).Enabled = False
   cmdOpcion(4).Enabled = False
   cmdOpcion(5).Enabled = False
   cmdOpcion(6).Enabled = False
   cmdOpcion(7).Enabled = False
   cmdOpcion(8).Enabled = False
   cmdOpcion(9).Enabled = False
End Sub
