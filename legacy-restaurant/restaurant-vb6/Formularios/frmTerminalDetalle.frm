VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTerminalDetalle 
   ClientHeight    =   6405
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSucursal 
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
      Left            =   3840
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtDescripcion 
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
      Left            =   3840
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Configuracion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtComercio 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   32
         Text            =   " "
         Top             =   960
         Width           =   3855
      End
      Begin VB.Frame fraGrilla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2880
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   5775
         Begin VB.CommandButton cmdOpcionGrilla 
            Caption         =   "Grabar"
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
            Index           =   3
            Left            =   3120
            Picture         =   "frmTerminalDetalle.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1995
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcionGrilla 
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
            Height          =   615
            Index           =   4
            Left            =   4440
            Picture         =   "frmTerminalDetalle.frx":0532
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1995
            Width           =   1170
         End
         Begin VB.TextBox txtCodigoC 
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
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   360
            Width           =   1170
         End
         Begin VB.CheckBox chkActivoC 
            Alignment       =   1  'Right Justify
            Caption         =   "Activo :"
            Height          =   195
            Left            =   600
            TabIndex        =   10
            Top             =   1200
            Width           =   885
         End
         Begin VB.Frame Frame3 
            Height          =   30
            Left            =   90
            TabIndex        =   9
            Top             =   1800
            Width           =   5565
         End
         Begin MSDataListLib.DataCombo cboCaja 
            Height          =   315
            Left            =   1305
            TabIndex        =   21
            Top             =   780
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            BackColor       =   16777215
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   15
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Caja :"
            Height          =   195
            Index           =   4
            Left            =   765
            TabIndex        =   14
            Top             =   825
            Width           =   405
         End
      End
      Begin VB.TextBox txtCodigo 
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   930
      End
      Begin VB.TextBox txtTerminal 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   6
         Text            =   " "
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdOpcionGrilla 
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
         Height          =   615
         Index           =   0
         Left            =   240
         Picture         =   "frmTerminalDetalle.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4680
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcionGrilla 
         Caption         =   "Modificar"
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
         Left            =   1560
         Picture         =   "frmTerminalDetalle.frx":0B66
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4680
         Width           =   1170
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   1890
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   3334
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
      Begin MSDataListLib.DataCombo cboIntegracion 
         Height          =   315
         Left            =   3885
         TabIndex        =   35
         Top             =   450
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Integracion :"
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   34
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Comercio :"
         Height          =   195
         Index           =   6
         Left            =   690
         TabIndex        =   33
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal :"
         Height          =   195
         Index           =   5
         Left            =   735
         TabIndex        =   30
         Top             =   1335
         Width           =   705
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion :"
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   20
         Top             =   2055
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   18
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Terminal :"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   17
         Top             =   1695
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   8565
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5655
      Width           =   8625
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
         Index           =   3
         Left            =   7245
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4890
         TabIndex        =   2
         Top             =   60
         Width           =   4950
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTerminalDetalle.frx":0C68
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTerminalDetalle.frx":11AA
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTerminalDetalle.frx":16EC
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4395
            Picture         =   "frmTerminalDetalle.frx":1C2E
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3915
            Picture         =   "frmTerminalDetalle.frx":2170
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3435
            Picture         =   "frmTerminalDetalle.frx":26B2
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
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
            Left            =   1500
            TabIndex        =   29
            Top             =   150
            Width           =   1830
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
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
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image1 
      Height          =   5610
      Left            =   0
      Picture         =   "frmTerminalDetalle.frx":2BF4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2205
   End
   Begin VB.Image Image 
      Height          =   1905
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "frmTerminalDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsGrilla As Recordset
Dim RsOperador As Recordset
Dim RsOperador2 As Recordset

Dim wAgrega  As Boolean
Dim tCajaUpdate As String


Sub Asignar()
    With frmTerminal.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtComercio = IIf(IsNull(!tComercio), "", !tComercio)
        txtSucursal = IIf(IsNull(!tSucursal), "", !tSucursal)
        txtTerminal = IIf(IsNull(!tTerminal), "", !tTerminal)
        
        txtDescripcion = IIf(IsNull(!tDescripcion), "", !tDescripcion)
        cboIntegracion.BoundText = IIf(IsNull(!tIntegracion), "", !tIntegracion)
        
        Set RsGrilla = Lib.OpenRecordset("SELECT TCT.tTerminal,TCT.tCaja,TC.tDescripcion,TCT.lActivo FROM TCAJATERMINAL TCT INNER JOIN TCAJA TC ON TCT.tCaja = TC.TCaja WHERE TCT.tTerminal='" & txtCodigo & "'", Cn)
        Set grdGrilla.DataSource = RsGrilla
        
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTerminal.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTerminal.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTerminal.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTerminal.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTerminal.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTerminal.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmTerminal.RsCabecera.AbsolutePosition & " de " & frmTerminal.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               txtTerminal.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
          
               'Chequea Datos
               If txtComercio.Text = "" Then MsgBox "Ingrese la Descripción del Comercio", vbExclamation, sMensaje: txtComercio.SetFocus: Exit Sub
               If txtSucursal.Text = "" Then MsgBox "Ingrese la Descripción de la Sucursal", vbExclamation, sMensaje: txtSucursal.SetFocus: Exit Sub
               If txtTerminal.Text = "" Then MsgBox "Ingrese la Descripción del Terminal", vbExclamation, sMensaje: txtTerminal.SetFocus: Exit Sub
               If cboIntegracion.Text = "" Then MsgBox "Ingrese el tipo de Integracion", vbExclamation, sMensaje: cboIntegracion.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("SELECT max(Codigo) as codigo FROM TTERMINAL", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Or nCorrela = "0" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTERMINAL( " & _
                         "Codigo, tIntegracion,tComercio, tSucursal,tTerminal,tDescripcion) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & cboIntegracion.BoundText & "', " & _
                                " '" & txtComercio.Text & "', " & _
                                " '" & txtSucursal.Text & "', " & _
                                " '" & txtTerminal.Text & "', " & _
                                " '" & txtDescripcion.Text & "') "
                                       
           
                      Cn.Execute Isql
                      
                      frmTerminal.RsCabecera.Requery
                      If frmTerminal.RsCabecera.RecordCount > 0 Then
                        frmTerminal.RsCabecera.Sort = "Codigo ASC"
                        frmTerminal.RsCabecera.MoveLast
                      End If
                      
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(frmTerminal.RsCabecera.RecordCount = 0, 0, frmTerminal.RsCabecera.AbsolutePosition) & " de " & frmTerminal.RsCabecera.RecordCount
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      
              Else
                 'Cambiar el SQL
                 Isql = "update TTERMINAL set " & _
                        "tComercio ='" & txtComercio.Text & "', " & _
                        "tSucursal ='" & txtSucursal.Text & "', " & _
                        "tTerminal ='" & txtTerminal.Text & "', " & _
                        "tDescripcion ='" & txtDescripcion.Text & "' " & _
                        " where Codigo = '" & txtCodigo.Text & "'"
                      
                  Cn.Execute Isql
                  
                  frmTerminal.RsCabecera.Requery
                  If frmTerminal.RsCabecera.RecordCount = 0 Then
                     nPos = frmTerminal.RsCabecera.Bookmark
                     frmTerminal.RsCabecera.Filter = adFilterNone
                  End If
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
              End If
         
          Case Is = 2 ' Eliminar
               'If frmTerminal.RsCabecera.RecordCount = 0 Then
               '   Exit Sub
               'End If
               'Cambia el MsgBox
               'If MsgBox("Seguro de Eliminar el Motivo de Reserva " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
               '   Exit Sub
               'End If
               
               'Cambia el Delete
               'Cn.Execute "delete from TTABLA where tTabla = 'TERMINAL' and tCodigo = '" & txtCodigo & "'"
               'frmTerminal.RsCabecera.Requery
               'If frmTerminal.RsCabecera.RecordCount <> 0 Then
               '   frmTerminal.RsCabecera.MoveLast
               '   Asignar
               '   cmdTexto.Caption = "Registro " & IIf(frmTerminal.RsCabecera.RecordCount = 0, 0, frmTerminal.RsCabecera.AbsolutePosition) & " de " & frmTerminal.RsCabecera.RecordCount
               '
               'Else
               '   ActivarBotones False
               '   Blanquear Me
               '   Sw = True
               'End If
          
          Case Is = 3 ' Salir
               Unload Me
   End Select
End Sub

Private Sub cmdOpcionGrilla_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
          
               'Cambiar los Controles
               With RsGrilla
                    cboCaja.Text = ""
                    cboCaja.BoundText = ""
                    chkActivoC.Value = 1
                End With
                
                Isql = "SELECT tCaja,tDescripcion FROM TCAJA WHERE tCaja NOT IN (SELECT tCaja FROM TCAJATERMINAL WHERE tTerminal = '" & txtCodigo.Text & "') AND lActivo=1"
                Set RsOperador = Lib.OpenRecordset(Isql, Cn)
                
                If Not RsOperador.RecordCount = 0 Then
                     cboCaja.Enabled = True
                     SubDetalle False
                     wAgrega = True
                    
                     With cboCaja
                          Set .RowSource = RsOperador
                             .DataField = "tDescripcion"
                             .ListField = "tDescripcion"
                             .BoundColumn = "tCaja"
                    End With
                End If
               
          
          Case Is = 1 ' Modificar
          
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               
               SubDetalle False
               wAgrega = False
               
               With cboCaja
                     Isql = "SELECT tCaja,tDescripcion AS Descripcion FROM TCAJA WHERE lActivo=1"
                     Set RsOperador = Lib.OpenRecordset(Isql, Cn)
                     Set .RowSource = RsOperador
                        .DataField = "tDescripcion"
                        .ListField = "Descripcion"
                        .BoundColumn = "tCaja"
               End With
               cboCaja.Enabled = False
               SubAsignar
              
               
               
          Case Is = 2 ' Eliminar
               'If RsGrilla.RecordCount = 0 Then
               '  Exit Sub
               'End If
               
               'Cambia el MsgBox
               'SubAsignar
'               If MsgBox("Seguro de Eliminar el Sub Tipo " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
'                  Exit Sub
'               End If
'
'               If Calcular("select count(tCodigoCliente) as Codigo From TCOMPANIA where tSubTipoCtaCte='" & txtCodigo.Text & "'", Cn) > 0 Then
'                  MsgBox "Se tiene movimientos con este SubTipo, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
'                  Exit Sub
'               End If
'
'               'Cambia el Delete
'               Cn.Execute "delete from TTABLA where tCodigo ='" & txtCodigo.Text & "' and TTABLA='SUBTIPOCTACTE'"
'               frmTipoCtaCte.RsCabecera.Requery
'               RsGrilla.Requery
'
'               If RsGrilla.RecordCount <> 0 Then
'                  RsGrilla.MoveLast
'               End If
          
          Case Is = 3 ' Grabar
          
               If cboCaja.Text = "" Then MsgBox "Seleccione la caja", vbExclamation, sMensaje: cboCaja.SetFocus: Exit Sub
               If Calcular("select COUNT(*) as codigo from TCAJA where tCaja = '" & cboCaja.BoundText & "' and lActivaIntegracionNiubiz = 1", Cn) = 0 And Calcular("select COUNT(*) as codigo from TCAJA where tCaja = '" & cboCaja.BoundText & "' and lActivaIntegracionIzipay = 1", Cn) = 0 Then MsgBox "La caja seleccionada no tiene habilitada ninguna integracion.", vbExclamation, sMensaje: Exit Sub
               
               If chkActivoC.Value Then
               
'                    Isql = "SELECT top 1 TT.tTerminal FROM TCAJATERMINAL TC INNER JOIN TTERMINAL TT ON TC.tTerminal = TT.Codigo where TT.tIntegracion = '" & cboIntegracion.BoundText & "' and TC.tCaja = '" & txtCodigoC.Text & "' and TC.lActivo = 1"
'
'                    Set RsOperador = Lib.OpenRecordset(Isql, Cn)
'
'                    If RsOperador.RecordCount = 1 Then
'                        MsgBox "Esta caja ya está activa en la terminal '" & RsOperador!tTerminal & "'", vbInformation, sMensaje
'                        Exit Sub
'                    End If
                    
                    
                Else
                    Isql = "select top 1 tTurno from MTURNO where tCaja = '" & cboCaja.BoundText & "' and lCierre = 0"
                    Set RsOperador = Lib.OpenRecordset(Isql, Cn)
                    
                    If RsOperador.RecordCount = 1 Then
                    
                        Isql = "select tTurno from DPAGOPOST where IdTerminal='" & txtCodigo.Text & "' and tTurno = '" & RsOperador!tTurno & "'"
                        Set RsOperador = Lib.OpenRecordset(Isql, Cn)
                        
                        If RsOperador.RecordCount > 0 Then
                            
                            Isql = "select TOP 1 * from MCIERREPOST where IdTerminal='" & txtCodigo.Text & "' and tTurno = '" & RsOperador!tTurno & "'"
                            Set RsOperador = Lib.OpenRecordset(Isql, Cn)
                            
                            If RsOperador.RecordCount = 0 And Not wAgrega Then
                                MsgBox "No es posible desactivar esta configuración, ya que se han detectado operaciones asociadas." & vbCrLf & _
                                "Para continuar, debe ejecutar un cierre del terminal.", vbInformation, sMensaje
                                Exit Sub
                            End If
                        End If
                    End If
                    
               End If
               
               If wAgrega Then
                                      
                  Isql = "insert into TCAJATERMINAL( " & _
                         "tTerminal, tCaja, lActivo) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & cboCaja.BoundText & "', " & _
                                " '" & chkActivoC.Value & "') "
           
                    Cn.Execute Isql
                    
                    Asignar
                    RsGrilla.Requery
                    RsOperador.Requery
                    
                    If RsGrilla.RecordCount > 0 Then
                        RsGrilla.MoveLast
                    End If
                    
                    
                    'MsgBox "Registro Guardado", vbInformation, sMensaje
                Else
                   Isql = "update TCAJATERMINAL set " & _
                        "tCaja ='" & cboCaja.BoundText & "', " & _
                        "lActivo ='" & chkActivoC.Value & "' " & _
                        " where tTerminal = '" & txtCodigo.Text & "' and tCaja = '" & txtCodigoC.Text & "'"
                      
                    Cn.Execute Isql
                    RsGrilla.Requery
                    If RsGrilla.RecordCount > 0 Then
                          RsGrilla.MoveLast
                    End If
                  'MsgBox "Registro Modificado", vbInformation, sMensaje
                End If
               SubDetalle True
          
          Case Is = 4 ' Cancelar
               cboCaja.Text = ""
               SubDetalle True
              
   End Select
End Sub

Public Sub SubDetalle(Activa As Boolean)

   fraGrilla.Visible = Not Activa
   ActivarBotones Activa
   
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   
   cmdOpcionGrilla(0).Enabled = Activa
   cmdOpcionGrilla(1).Enabled = Activa
         
   'Controles de la cabecera
   txtTerminal.Enabled = Activa
   txtDescripcion.Enabled = Activa
   
End Sub

Private Sub cboCaja_Change()
    txtCodigoC.Text = cboCaja.BoundText
End Sub


Sub SubAsignar()
    'Cambiar la Asignación del Detalle
    With RsGrilla
         cboCaja.Text = IIf(IsNull(!tCaja), "", !tCaja)
         cboCaja.BoundText = IIf(IsNull(!tCaja), "", !tCaja)
         txtCodigoC.Text = IIf(IsNull(!tCaja), "", !tCaja)
         chkActivoC.Value = IIf(IsNull(!lActivo), 1, IIf(!lActivo = True, 1, 0))
    End With
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    fraGrilla.Visible = False
    
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Terminal "
    fraDetalle.Caption = Me.Caption
    
    
    Isql = "SELECT TCODIGO,tValor FROM TTABLA WHERE TTABLA = 'INTEGRACION'"
    Set RsOperador = Lib.OpenRecordset(Isql, Cn)
    
    If Not RsOperador.RecordCount = 0 Then
         With cboIntegracion
              Set .RowSource = RsOperador
                 .DataField = "tValor"
                 .ListField = "tValor"
                 .BoundColumn = "TCODIGO"
        End With
    End If
                
                
    
    Call ConfGrilla(3, grdGrilla, "Código", 2, "tCaja", 1200, 2, 0, "", _
                                  "Descripcion", 2, "tDescripcion", 2300, 0, 0, "", _
                                  "Activo", 2, "lActivo", 1000, 2, 4, "")
              
    
    Set RsGrilla = Lib.OpenRecordset("SELECT tCaja,lActivo FROM TCAJATERMINAL WHERE tTerminal='" & txtCodigo.Text & "'", Cn)
    Set grdGrilla.DataSource = RsGrilla
                     
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       cmdOpcion(1).Enabled = True
       
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
       cboIntegracion.Enabled = False
    End If
    
    cmdTexto.Caption = "Registro " & frmTerminal.RsCabecera.AbsolutePosition & " de " & frmTerminal.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTerminalDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    
    cmdOpcion(1).Enabled = Activa
    'cmdOpcion(3).Enabled = Activa
    
    cmdOpcionGrilla(0).Enabled = Activa
    cmdOpcionGrilla(1).Enabled = Activa
End Sub

