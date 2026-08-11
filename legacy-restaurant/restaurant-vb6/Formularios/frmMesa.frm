VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmMesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mesas"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   11880
   FillColor       =   &H8000000B&
   Icon            =   "frmMesa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5715
      TabIndex        =   33
      Top             =   7380
      Width           =   3255
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   2565
         MaxLength       =   3
         TabIndex        =   35
         Text            =   "0"
         Top             =   315
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Juntar Mesas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   7
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad de Mesas Juntas"
         Height          =   420
         Left            =   1395
         TabIndex        =   36
         Top             =   262
         Width           =   1050
      End
   End
   Begin VB.Frame fraLocal 
      Caption         =   " Local "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   9030
      TabIndex        =   30
      Top             =   45
      Width           =   2805
      Begin VB.CommandButton cmdSalon 
         Height          =   555
         Index           =   0
         Left            =   90
         TabIndex        =   32
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdBusca 
         Height          =   555
         Left            =   1485
         Picture         =   "frmMesa.frx":2582
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2085
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   3
      Left            =   8175
      Picture         =   "frmMesa.frx":2B0C
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6780
      Width           =   795
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   2
      Left            =   7380
      Picture         =   "frmMesa.frx":33D6
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6780
      Width           =   795
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   1
      Left            =   6585
      Picture         =   "frmMesa.frx":3CA0
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6780
      Width           =   795
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   0
      Left            =   5790
      Picture         =   "frmMesa.frx":456A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6780
      Width           =   795
   End
   Begin VB.Frame fraMesa 
      Caption         =   " Mesas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5715
      Begin VB.CommandButton cmdMesa 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   540
         Width           =   915
      End
      Begin VB.Image imgFondo 
         Height          =   7815
         Left            =   120
         Stretch         =   -1  'True
         Top             =   330
         Width           =   5505
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Mesa "
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
      Left            =   9030
      TabIndex        =   18
      Top             =   7620
      Width           =   2805
      Begin VB.TextBox txtPersona 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Left            =   1545
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtMesa 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Asientos"
         Height          =   195
         Left            =   2115
         TabIndex        =   22
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "con"
         Height          =   195
         Left            =   1155
         TabIndex        =   21
         Top             =   300
         Width           =   270
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   10410
      TabIndex        =   12
      Top             =   4950
      Width           =   1425
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
         Height          =   495
         Index           =   3
         Left            =   120
         Picture         =   "frmMesa.frx":4E34
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Width           =   1215
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
         Height          =   555
         Index           =   1
         Left            =   90
         Picture         =   "frmMesa.frx":4F26
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1470
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Refrescar"
         Height          =   555
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   855
         Width           =   1275
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
         Height          =   555
         Index           =   0
         Left            =   90
         Picture         =   "frmMesa.frx":5028
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2085
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Sin Mesa"
         Height          =   555
         Index           =   6
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame fraComensal 
      Caption         =   " Comensales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   9030
      TabIndex        =   7
      Top             =   4950
      Width           =   1365
      Begin VB.TextBox txtNino 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "0"
         Top             =   2115
         Width           =   1095
      End
      Begin VB.TextBox txtAdulto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "0"
         Top             =   900
         Width           =   1050
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Adultos"
         Height          =   555
         Index           =   4
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Niños"
         Height          =   585
         Index           =   5
         Left            =   90
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame fraEstado 
      Caption         =   " Estado de las Mesas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   9030
      TabIndex        =   0
      Top             =   2835
      Width           =   2805
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H000080FF&
         Caption         =   "Precuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   8
         Left            =   1440
         MaskColor       =   &H00004080&
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1650
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H00004080&
         Caption         =   "Original"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   7
         Left            =   90
         MaskColor       =   &H00004080&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1650
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H00FF8080&
         Caption         =   "Lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ocupada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H00008000&
         Caption         =   "Reservada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   3
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   690
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H0000C0C0&
         Caption         =   "Sucia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   4
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H000000C0&
         Caption         =   "Bloqueada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   5
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1170
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fuera de Servicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Index           =   6
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1170
         Width           =   1275
      End
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   6690
      Left            =   5760
      TabIndex        =   29
      Top             =   30
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   11800
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
End
Attribute VB_Name = "frmMesa"
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
Dim nCantidad As Integer

Dim i As Integer

Private Sub cmdBusca_Click()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
       
   If sTipo = "M" And wClick Then
     wClick = False
   End If
        
   sTemp = ""
   Isql = "select * from vSalon where lActivo = 1"
   Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                   "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
   frmBusquedaRapida.nPredeterm = 1
   frmBusquedaRapida.Show vbModal
   If wEnter = True Then
      RsSalon.MoveFirst
      RsSalon.Find "tSalon='" & sCodigo & "'"
      fraMesa.Caption = " " & RsSalon!dSalon & " "
      InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
      Desasigna (RsMesa.RecordCount)
      AsignaMesa (sTipo)
   End If
End Sub

Private Sub cmdEstado_Click(Index As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
  
    Dim nPos As Integer
    If sCodigo = "" Then
       MsgBox "Debe escoger una Mesa", vbExclamation, sMensaje
       Exit Sub
    End If
    
    If RsMesa!tEstadoMesa = "02" Then
       Exit Sub
    End If
    
    If Estado = "02" Then
       Exit Sub
    End If
    
    Estado = ""
    Estado = "0" + Trim(str(Index))
    If RsMesa!tEstadoMesa = Estado Then
       MsgBox "El Estado debe ser diferente", vbExclamation, sMensaje
       Exit Sub
    End If
    
    If sTipo = "V" And Estado <> "" Then
       If Calcular("select count(tEstadoMesa) as Codigo From TMESA where tCodigoMesa='" & sCodigo & "' and tEstadoMesa='02'", Cn) > 0 Then
          Refrescar
          MsgBox "La Mesa ya ha sido ocupada", vbCritical, sMensaje
          Exit Sub
       End If
    
       Isql = "Update TMESA set tEstadoMesa='" & Estado & "' where tCodigoMesa ='" & sCodigo & "'"
       Cn.Execute Isql
       nPos = RsMesa.AbsolutePosition
       RsMesa.Requery
       RsMesa.AbsolutePosition = nPos
       Select Case Estado
              Case Is = "01"
                   cmdMesa(nIndex).backColor = vbLista
              Case Is = "02"
                   cmdMesa(nIndex).backColor = vbOcupada
              Case Is = "03"
                   cmdMesa(nIndex).backColor = vbReservada
              Case Is = "04"
                   cmdMesa(nIndex).backColor = vbSucia
              Case Is = "05"
                   cmdMesa(nIndex).backColor = vbBloqueada
              Case Is = "06"
                   cmdMesa(nIndex).backColor = vbFServicio
       End Select
    End If
    
End Sub

Private Sub cmdSalon_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
   If sTipo = "M" And wClick Then
      wClick = False
   End If
   
   RsSalon.MoveFirst
   RsSalon.Move Index - 1
   fraMesa.Caption = " " & RsSalon!dSalon & " "
   InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
   Desasigna (RsMesa.RecordCount)
   AsignaMesa (sTipo)
End Sub

Private Sub cmdMesa_Click(Index As Integer)
     If sModulo = "ADICION" Then
        frmMozoUsuario.ReseteaTimer
     End If

     wClick = True
     RsMesa.MoveFirst
     RsMesa.Move (Index - 1)
     sCodigo = RsMesa!tCodigoMesa
     txtMesa.Text = IIf(IsNull(RsMesa!tResumido), "", RsMesa!tResumido)
     txtPersona.Text = IIf(IsNull(RsMesa!nPersona), "?", RsMesa!nPersona)
     xIndex = nIndex
     nIndex = Index
          
     If sTipo <> "V" Then
        wCambioMesa = True
        If sTipo = "M" Then
           If xIndex <> 0 Then
              If RsSalon!tSalon <> xSalon Then
                 xIndex = 0
                 xSalon = RsSalon!tSalon
                 RsSalon.MoveFirst
                 RsSalon.Find ("tSalon ='" & xSalon & "'")
              End If
              cmdMesa(xIndex).backColor = vbLista
              cmdMesa(nIndex).backColor = vbOcupada
           End If
           If nIndex <> xIndex Then
              cmdMesa(nIndex).Caption = cmdMesa(nIndex).Caption & " " & sTitulo
              cmdMesa(xIndex).Caption = sMesa(xIndex, 1)
              cmdMesa(nIndex).backColor = vbOcupada
           End If
        Else
           If RsSalon!tSalon <> xSalon Then
              xIndex = 0
              xSalon = RsSalon!tSalon
              RsSalon.MoveFirst
              RsSalon.Find ("tSalon ='" & xSalon & "'")
           End If
           cmdMesa(xIndex).backColor = vbLista
           cmdMesa(nIndex).backColor = vbOcupada
           
           txtAdulto.Text = "0"
           txtNino.Text = "0"
        End If
     Else
        If RsMesa!tEstadoMesa = "02" Then
            sPedido = sMesa(Index, 2)
            sTipo = ""
            frmPedido.Show vbModal
        End If
     End If
               
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If

    Select Case Index
           Case Is = 0  ' Ultimo Detalle
                MoverPuntero Primero, grdGrilla
           Case Is = 1  ' Siguiente Detalle
                MoverPuntero previo, grdGrilla
           Case Is = 2  ' Previo Detalle
                MoverPuntero siguiente, grdGrilla
           Case Is = 3  ' Primero Detalle
                MoverPuntero Ultimo, grdGrilla
   End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

  Select Case Index
         Case Is = 0 ' Aceptar
              If sCodigo = "" Then
                 MsgBox "Debe escoger una Mesa", vbExclamation, sMensaje
                 Exit Sub
              End If
              
              If lObligaPax And val(txtAdulto) = 0 Then
                 MsgBox "Debe ingresar el numero de comensales", vbExclamation, sMensaje
                 Exit Sub
              Else
                  Dim nPersonas As Integer
                  nPersonas = Calcular("Select nPersona As Codigo From TMESA where tCodigoMesa = '" & sCodigo & "'", Cn)
                  If nPersonas < val(txtAdulto) Then
                        If MsgBox("El numero de Pax(Adultos) supera al definido para esta mesa," & Chr(13) & "¿Desea Continuar?", vbCritical + vbYesNo, sMensaje) = vbYes Then
                        Else
                           CmdOpcion(4).SetFocus
                           Exit Sub
                        End If
                  End If
              End If
              
              
              If sTipo = "A" And wCambioMesa = False Then
                 MsgBox "Debe Asignar la Mesa", vbExclamation, sMensaje
                 Exit Sub
              End If
              If sTipo <> "V" Then
                 frmVenta.txtAdulto.Text = txtAdulto.Text
                 frmVenta.txtNino.Text = txtNino.Text
                 frmVenta.CmdOpcion(1).Enabled = True
              End If
              If RsMesa.EOF Then
                 wMesa = False
              Else
                 wMesa = True
                 wEnter = True
              End If
              Unload Me
              
         Case Is = 1 ' Cancelar
              sCodigo = ""
              sTemp = ""
              wEnter = False
              Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='')"
              Cn.Execute "delete from TPEDIDOMESA where tCodigoPedido=''"
              Unload Me

         Case Is = 2 ' Refrescar
              Refrescar
         
         Case Is = 3 ' Salir
              Unload Me
              
         Case Is = 4 ' Adulto
         
              sTipo = ""
              If sCodigo = "" Then
                 MsgBox "seleccione una Mesa", vbExclamation, sMensaje
              Else
                 frmNumPad.Show vbModal
                 txtAdulto.Text = IIf(wEnter = True, sDescrip, txtAdulto.Text)
              End If
         
         Case Is = 5 ' Niño
         
              frmNumPad.Show vbModal
              txtNino.Text = IIf(wEnter = True, Format(sDescrip, "#,##0"), txtNino.Text)

         Case Is = 6 ' Sin Mesa
              If lAutorizacionSinMeza = False Then
                    wMesa = True
                    wCambioMesa = True
                    sTemp = sSinMesa
                    sCodigo = ""
                    Call Log_Inforest("Mpedido", "MODULO INFOREST", "", "Caja: " + sCaja, Environ("computername"), "Cambio de Mesa Sin Pedido", "", "CAMBIO DE MESA SIN PEDIDO", sUsuario)
                    Unload Me
              Else
                    'frmPassword.Caption = "Ingrese su nueva clave"
                    Dim xClaveAutoriza As String
                    Dim xUsuarioAutorizapedidosinmesa As String
                    Dim xPerfilUsusario As String
                    
                    xPerfilUsusario = ""
                    xUsuarioAutorizapedidosinmesa = ""
                    xClaveAutoriza = ""
                   
                    frmPassword.Show vbModal
                    If wEnter Then
                        xClaveAutoriza = sDescrip

                        If xClaveAutoriza <> "" Then

                            xUsuarioAutorizapedidosinmesa = VerificarAutorizacionPedidoSinMesa(xClaveAutoriza)

                            If xUsuarioAutorizapedidosinmesa = "" Then
                                MsgBox "Contraseña Incorrecta", vbExclamation, sMensaje
                                Exit Sub
                            Else
                               xPerfilUsusario = Mid(xUsuarioAutorizapedidosinmesa, 1, 1)
                               xUsuarioAutorizapedidosinmesa = Mid(xUsuarioAutorizapedidosinmesa, 2, Len(xUsuarioAutorizapedidosinmesa))
                               wMesa = True
                               wCambioMesa = True
                               sTemp = sSinMesa
                               sCodigo = ""
                               Call Log_Inforest("Mpedido", "MODULO INFOREST", "", "Caja: " + sCaja, Environ("computername"), "Cambio de Mesa Sin Pedido", "", "CAMBIO DE MESA SIN PEDIDO", xUsuarioAutorizapedidosinmesa)
                               Unload Me

                            End If

                        Else
                            Exit Sub
                        End If
                    Else
                       Exit Sub
                    End If
                             
                       
            
                   
              End If
         
         Case Is = 7 'Juntar Mesa
              wEnter = False
              If sCodigo = "" Then
                 MsgBox "seleccione una Mesa", vbExclamation, sMensaje
                 Exit Sub
              End If
              frmJuntaMesa.sSalonActual = xSalon
              frmJuntaMesa.Show vbModal
              If wEnter Then
                 RsMesa.Requery
                 Desasigna (RsMesa.RecordCount)
                 AsignaMesa (sTipo)
                 RsMesa.MoveFirst
                 Dim tAdultoTemp As String
                 Dim tNinoTemp As String
                 Dim tEstatusTemp As Boolean
                 tAdultoTemp = txtAdulto.Text
                 tNinoTemp = txtNino.Text
                 tEstatusTemp = wCambioMesa
                 cmdMesa_Click (nIndex)
                 wCambioMesa = tEstatusTemp
                 txtAdulto.Text = tAdultoTemp
                 txtNino.Text = tNinoTemp
                 Text1.Text = nVar1
              End If
                  
  End Select
End Sub

Private Sub Form_Load()
   Dim i As Integer
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
   wClick = False
   sCodigo = ""
   sDescrip = ""
   sTemp = ""
   xIndex = 0
   wMesa = False
   wCambioMesa = False
      
   Set RsMesa = Lib.OpenRecordset("select * from TMESA where tCodigoMesa <> '000' and lActivo= 1 order by tCodigoMesa", Cn)
   Set RsSalon = Lib.OpenRecordset("select tCodigo as tSalon, tDetallado as dSalon, tiCono from TTABLA where tTABLA ='SALON' and lActivo = 1 ", Cn)
   Set RsPedido = Lib.OpenRecordset("select tCodigoPedido as Codigo, tMesa, fFecha, nPrecuenta from MPEDIDO where tEstadoPedido = '01' order by tMesa", Cn)
   Set RsGrilla = Lib.OpenRecordset("select tObservacion from MPEDIDO where tEstadoPedido = '01' and tTipoPedido <> '04' and not len(rtrim(tObservacion)) = 0 and tCaja = '" & sCaja & "' order by tObservacion", Cn)
   
   'Configuración de Grillas
   Call ConfGrilla(1, grdGrilla, "Mesa", 2, "tObservacion", 2650, 0, 0, "")
   Set grdGrilla.DataSource = RsGrilla
   
   'Salon
   Call MatrizBotones(IIf(RsSalon.RecordCount >= 7, 7, RsSalon.RecordCount), 2, 60, 60, cmdSalon)
   For i = 1 To IIf(RsSalon.RecordCount >= 7, 7, RsSalon.RecordCount)
       cmdSalon(i).Caption = RsSalon!dSalon
       RsSalon.MoveNext
   Next i
      
   If sTipo = "V" Then     'Visualizacion
      ActivaEstado True
      ActivaOpcion False
      CmdOpcion(6).Visible = False
      RsSalon.MoveFirst
      RsSalon.Find ("tSalon='" & sSalon & "'")
      If RsSalon.EOF Then
         RsSalon.MoveFirst
      End If
      
   ElseIf sTipo = "M" Then 'Modifica Pedido
      txtAdulto.Text = frmVenta.txtAdulto.Text
      txtNino.Text = frmVenta.txtNino.Text
      txtPersona.Text = IIf(IsNull(RsMesa!nPersona), "?", RsMesa!nPersona)
      nCantidad = Calcular("select count(tMesa) as Codigo From TPEDIDOMESA where tCodigoPedido='" & sPedido & "'", Cn)
      Text1.Text = Format(nCantidad, "##0")
      
      If lBloquearPedidoSinMesa = True Then
        CmdOpcion(6).Visible = False
      End If
      ActivaEstado (False)
      ActivaOpcion True
      If RsPedido.RecordCount > 0 Then
        RsPedido.MoveFirst
      End If
      RsPedido.Find ("Codigo ='" & sPedido & "'")

      If Not RsPedido.EOF Then
         RsMesa.MoveFirst
         RsMesa.Find ("tCodigoMesa ='" & RsPedido!tMesa & "'")
         If Not RsMesa.EOF Then
            RsSalon.MoveFirst
            RsSalon.Find ("tSalon ='" & RsMesa!tSalon & "'")
            If RsSalon.EOF Then
               RsSalon.MoveFirst
               RsSalon.Find ("tSalon='" & sSalon & "'")
               If RsSalon.EOF Then
                  RsSalon.MoveFirst
               End If
            End If
         Else
            RsSalon.MoveFirst
            RsSalon.Find ("tSalon='" & sSalon & "'")
            If RsSalon.EOF Then
               RsSalon.MoveFirst
            End If
            
         End If
      Else
         RsSalon.MoveFirst
         If RsSalon.EOF Then
            RsSalon.MoveFirst
         End If
      End If
   Else           'Adiciona
      ActivaEstado (False)
      ActivaOpcion True
      RsSalon.MoveFirst
      RsSalon.Find ("tSalon='" & sSalon & "'")
      If RsSalon.EOF Then
         RsSalon.MoveFirst
      End If
            
   End If
   xSalon = RsSalon!tSalon
   
   fraMesa.Caption = " " & RsSalon!dSalon & " "
   InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
   AsignaMesa (sTipo)
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaMesa(Opcion As String)
   Dim i As Integer
   Dim nHora As Long
   RsMesa.Filter = "tSalon = '" & RsSalon!tSalon & "'"
   If RsMesa.RecordCount = 0 Then
      Exit Sub
   End If
   
   RsMesa.MoveFirst
   cmdMesa(0).Visible = False
   ReDim sMesa(RsMesa.RecordCount, 2)
   
   For i = 1 To RsMesa.RecordCount
       'Cantidad de Mesas
       With cmdMesa(i)
            Load cmdMesa(i)
            .Top = RsMesa!tX
            .Left = RsMesa!tY
            .Visible = True
            cmdMesa(i).Caption = Trim(RsMesa!tResumido)
            sMesa(i, 1) = Trim(RsMesa!tResumido)
            
            If sTipo = "V" Then
               cmdMesa(i).Enabled = True
            Else
               cmdMesa(i).Enabled = False
            End If
            
            Select Case RsMesa!tEstadoMesa
                   Case Is = "01" ' Lista
                        cmdMesa(i).Enabled = True
                        cmdMesa(i).backColor = vbLista
                        
                   Case Is = "02" ' Ocupada
                        If RsPedido.RecordCount <> 0 Then
                           RsPedido.MoveFirst
                           RsPedido.Find ("tMesa = '" & RsMesa!tCodigoMesa & "'")
                           If Not RsPedido.EOF Then
                              nHora = DateDiff("n", RsPedido!fFecha, FechaServidor())
                              sMesa(i, 2) = RsPedido!codigo
                              sTitulo = Format(Format(Trim(str(Int(nHora / 60))), "00") & ":" & Format(Trim(str(nHora Mod 60)), "00"), "hh:mm")
                              cmdMesa(i).Caption = "" & cmdMesa(i).Caption & Chr(13) & sTitulo
                                                                                                            
                              If sTipo = "M" Then 'Se separó en dos if para que no se ejecute el load
                                 If RsPedido!codigo = sPedido Then
                                    cmdMesa(i).backColor = &H4080&
                                    'sTemp = RsMesa!tCodigoMesa
                                    sTempMesa = RsMesa!tCodigoMesa
                                    nIndex = i
                                    sCodigo = RsMesa!tCodigoMesa
                                    txtMesa.Text = IIf(IsNull(RsMesa!tResumido), "", RsMesa!tResumido)
                                    sSinMesa = sCodigo
                                    xSalon = RsSalon!tSalon
                                 Else
                                    cmdMesa(i).backColor = IIf(IsNull(RsPedido!nPrecuenta), vbOcupada, vbPrecuentaImp)
                                    'cmdMesa(i).BackColor = vbOcupada
                                 End If
                              ElseIf sTipo = "A" Then
                                 If RsPedido!codigo = sPedido Then
                                    cmdMesa(i).backColor = vbOcupada
                                    cmdMesa(i).Enabled = True
                                    sTemp = RsMesa!tCodigoMesa
                                    nIndex = i
                                 Else
                                    cmdMesa(i).backColor = IIf(IsNull(RsPedido!nPrecuenta), vbOcupada, vbPrecuentaImp)
                                    'cmdMesa(i).BackColor = vbOcupada
                                 End If
                              Else
                                 cmdMesa(i).Enabled = True
                                 cmdMesa(i).backColor = IIf(IsNull(RsPedido!nPrecuenta), vbOcupada, vbPrecuentaImp)
                                 'cmdMesa(i).BackColor = vbOcupada
                              End If
                              
                           Else
                              cmdMesa(i).backColor = IIf(IsNull(RsPedido!nPrecuenta), vbOcupada, vbPrecuentaImp)
                              'cmdMesa(i).BackColor = vbOcupada
                           End If
                        End If
                   Case Is = "03" ' Reservada
                        cmdMesa(i).backColor = vbReservada
                   Case Is = "04" ' Sucia
                        cmdMesa(i).backColor = vbSucia
                        cmdMesa(i).Enabled = True
                   Case Is = "05" ' Bloqueada
                        cmdMesa(i).backColor = vbBloqueada
                   Case Is = "06" ' Fuera de Servicio
                        cmdMesa(i).backColor = vbFServicio
                   Case Is = "07" ' Original
                        cmdMesa(i).backColor = vbOriginal
            End Select
            
       End With
       RsMesa.MoveNext
   Next i
   
   If sTipo = "M" Then
      RsMesa.MoveFirst
      RsMesa.Find ("tCodigoMesa = '" & sCodigo & "'")
   End If
End Sub

Public Sub Desasigna(Cantidad As Integer)
   Dim i As Integer
   For i = 1 To Cantidad
       Unload cmdMesa(i)
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsMesa = Nothing
   Set RsSalon = Nothing
   Set RsPedido = Nothing
   Set frmMesa = Nothing
End Sub

Public Sub ActivaEstado(Activa As Boolean)
   cmdEstado(1).Enabled = Activa
   cmdEstado(2).Enabled = False
   cmdEstado(3).Enabled = Activa
   cmdEstado(4).Enabled = Activa
   cmdEstado(5).Enabled = Activa
   cmdEstado(6).Enabled = Activa
   cmdEstado(7).Enabled = False
   cmdEstado(8).Enabled = False
End Sub

Public Sub ActivaOpcion(Activa As Boolean)
    CmdOpcion(0).Visible = Activa
    CmdOpcion(1).Visible = Activa
    CmdOpcion(3).Visible = Not Activa
End Sub

Private Sub grdGrilla_DblClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
End Sub

Public Sub Refrescar()
    RsMesa.Requery
    Desasigna (RsMesa.RecordCount)
    AsignaMesa (sTipo)
    RsMesa.MoveFirst
    txtMesa.Text = ""
    txtPersona.Text = ""
    sCodigo = ""
End Sub
Public Function VerificarAutorizacionPedidoSinMesa(ByVal tClave As String) As String
    Dim RsUsuario As New ADODB.Recordset
    Dim RsMozo As New ADODB.Recordset
    Dim tCodigoPersona As String
    Dim tDescPersona As String
    Dim ttipoPersona As String 'USUARIO 'MESERO
    
    
    tDescPersona = ""
    ttipoPersona = ""
    
    Set RsUsuario = Lib.OpenRecordset("SELECT tCodigoUsuario, tResumido FROM vGrupoUsuario  WHERE tPassword = '" & Encapsula(tClave) & "' and lActivo = 1 and lopcion33 = 1", Cn)
    
    If RsUsuario.RecordCount > 0 Then
        If Not (RsUsuario.EOF Or RsUsuario.BOF) Then
            RsUsuario.MoveFirst
            tCodigoPersona = RsUsuario.Fields("tCodigoUsuario")
            tDescPersona = "U" + RsUsuario.Fields("tResumido")
            ttipoPersona = "USUARIO"
        End If
    End If
   
    Set RsMozo = Lib.OpenRecordset("SELECT * FROM vmozo WHERE tvalor='" & Encapsula(tClave) & "' and lactivo=1 and lAutorizaEnviosMozo=1", Cn)
    If RsMozo.RecordCount > 0 Then
        If Not (RsMozo.EOF Or RsMozo.BOF) Then
            RsMozo.MoveFirst
            tCodigoPersona = RsMozo.Fields("Codigo")
            tDescPersona = "M" + RsMozo.Fields("tResumido")
            ttipoPersona = Mesero
        End If
    End If
    
    VerificarAutorizacionPedidoSinMesa = tDescPersona
  
End Function
