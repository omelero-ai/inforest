VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmTipoCtaCteDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5295
   ClientLeft      =   -570
   ClientTop       =   1875
   ClientWidth     =   9705
   Icon            =   "frmTipoCtaCteDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9705
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
      Height          =   4515
      Left            =   2430
      TabIndex        =   25
      Top             =   0
      Width           =   7215
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
         Height          =   2760
         Left            =   45
         TabIndex        =   26
         Top             =   1710
         Width           =   6975
         Begin VB.TextBox txtCuentaContable 
            Height          =   285
            Left            =   4680
            TabIndex        =   37
            Top             =   270
            Width           =   2055
         End
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
            Left            =   4410
            Picture         =   "frmTipoCtaCteDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   15
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
            Left            =   5640
            Picture         =   "frmTipoCtaCteDetalle.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1995
            Width           =   1170
         End
         Begin VB.TextBox txtDetalladoS 
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
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   9
            Text            =   " "
            Top             =   660
            Width           =   4890
         End
         Begin VB.TextBox txtResumidoS 
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
            Left            =   1845
            MaxLength       =   24
            TabIndex        =   10
            Text            =   " "
            Top             =   1050
            Width           =   2595
         End
         Begin VB.TextBox txtCodigoS 
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
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   270
            Width           =   1170
         End
         Begin VB.CheckBox chkActivoS 
            Alignment       =   1  'Right Justify
            Caption         =   "Activo :"
            Height          =   195
            Left            =   1170
            TabIndex        =   11
            Top             =   1440
            Width           =   885
         End
         Begin VB.Frame Frame3 
            Height          =   30
            Left            =   90
            TabIndex        =   27
            Top             =   1800
            Width           =   6765
         End
         Begin VB.Label lblCuentaContable 
            Caption         =   "Cuenta Contable: "
            Height          =   255
            Left            =   3360
            TabIndex        =   36
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Index           =   3
            Left            =   1155
            TabIndex        =   35
            Top             =   315
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Resumida :"
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   34
            Top             =   1095
            Width           =   1680
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Detallada :"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   33
            Top             =   705
            Width           =   1650
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtResumido 
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
         Left            =   1965
         MaxLength       =   24
         TabIndex        =   2
         Text            =   " "
         Top             =   1040
         Width           =   2595
      End
      Begin VB.TextBox txtDetallado 
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
         Left            =   1965
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   640
         Width           =   5070
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1275
         TabIndex        =   3
         Top             =   1440
         Width           =   885
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
         Left            =   180
         Picture         =   "frmTipoCtaCteDetalle.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3645
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
         Left            =   1410
         Picture         =   "frmTipoCtaCteDetalle.frx":0FA8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3645
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcionGrilla 
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
         Height          =   615
         Index           =   2
         Left            =   2640
         Picture         =   "frmTipoCtaCteDetalle.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3645
         Width           =   1170
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   1770
         Left            =   45
         TabIndex        =   32
         Top             =   1755
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   3122
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1260
         TabIndex        =   30
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   29
         Top             =   1065
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   28
         Top             =   675
         Width           =   1650
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9645
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4545
      Width           =   9705
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   645
         Left            =   30
         ScaleHeight     =   585
         ScaleWidth      =   4875
         TabIndex        =   18
         Top             =   60
         Width           =   4935
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3435
            Picture         =   "frmTipoCtaCteDetalle.frx":11AC
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3915
            Picture         =   "frmTipoCtaCteDetalle.frx":16EE
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4395
            Picture         =   "frmTipoCtaCteDetalle.frx":1C30
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoCtaCteDetalle.frx":2172
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoCtaCteDetalle.frx":26B4
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoCtaCteDetalle.frx":2BF6
            Style           =   1  'Graphical
            TabIndex        =   19
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
            TabIndex        =   31
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
         Left            =   6135
         Picture         =   "frmTipoCtaCteDetalle.frx":3138
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
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
         Left            =   4965
         Picture         =   "frmTipoCtaCteDetalle.frx":366A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
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
         Height          =   615
         Index           =   2
         Left            =   7305
         Picture         =   "frmTipoCtaCteDetalle.frx":3B9C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1170
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
         Index           =   3
         Left            =   8475
         Picture         =   "frmTipoCtaCteDetalle.frx":3C9E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   4410
      Left            =   30
      Picture         =   "frmTipoCtaCteDetalle.frx":3D90
      Stretch         =   -1  'True
      Top             =   90
      Width           =   2295
   End
End
Attribute VB_Name = "frmTipoCtaCteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrilla As Recordset
Dim nCorrela As String
Dim wAgrega  As Boolean
Dim nOrden   As Integer

Sub Asignar()
    With frmTipoCtaCte.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
    
    'Cambiar el Filtro
    RsGrilla.Filter = "tTipoCtaCte ='" & txtCodigo.Text & "'"
    If RsGrilla.RecordCount > 0 Then
       SubAsignar
    Else
       txtCodigoS.Text = ""
       txtDetalladoS.Text = ""
       txtResumidoS.Text = ""
       chkActivoS.Value = 0
    End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoCtaCte.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoCtaCte.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoCtaCte.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoCtaCte.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoCtaCte.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoCtaCte.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmTipoCtaCte.RsCabecera.AbsolutePosition & " de " & frmTipoCtaCte.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.Value = 1
                
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
               cmdOpcionGrilla(0).Enabled = False
               cmdOpcionGrilla(1).Enabled = False
               cmdOpcionGrilla(2).Enabled = False
               RsGrilla.Filter = "Codigo ='" & txtCodigo.Text & "'"
               
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               
               If Sw Then
                  'Asignar El Campo de Codificación
                  nCorrela = Calcular("select max(Codigo) as Codigo from vTipoCtaCte", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     txtCodigo = "01"
                  Else
                     txtCodigo = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, lActivo) " & _
                         "values ('TIPOCTACTE', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       chkActivo.Value & ") "
                  
                  Cn.Execute Isql
                  frmTipoCtaCte.RsCabecera.Requery
                  frmTipoCtaCte.RsCabecera.MoveLast
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmTipoCtaCte.RsCabecera.RecordCount = 0, 0, frmTipoCtaCte.RsCabecera.AbsolutePosition) & " de " & frmTipoCtaCte.RsCabecera.RecordCount
           
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "lActivo =" & chkActivo.Value & _
                         " where tTABLA = 'TIPOCTACTE' and tCodigo='" & txtCodigo.Text & "'"
                   Cn.Execute Isql
                   nPos = frmTipoCtaCte.RsCabecera.Bookmark
                   frmTipoCtaCte.RsCabecera.Requery
                   If frmTipoCtaCte.RsCabecera.RecordCount = 0 Then
                      frmTipoCtaCte.RsCabecera.Filter = adFilterNone
                   End If
                   frmTipoCtaCte.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If

               cmdOpcionGrilla(0).Enabled = True
               cmdOpcionGrilla(1).Enabled = True
               cmdOpcionGrilla(2).Enabled = True
                                                
          Case Is = 2 ' Eliminar
               If frmTipoCtaCte.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
            
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Tipo de Cuenta Corriente " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                                    
               If Calcular("select count(tCodigoCliente) as Codigo From TCOMPANIA where tTipoCtaCte='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este Tipo, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
                  
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tCodigo ='" & txtCodigo.Text & "' and tTABLA='TIPOCTACTE'"
               Cn.Execute "delete from TTABLA where tValor ='" & txtCodigo.Text & "' and tTABLA='SUBTIPOCTACTE'"
               frmTipoCtaCte.RsCabecera.Requery
               RsGrilla.Requery
               
               If frmTipoCtaCte.RsCabecera.RecordCount <> 0 Then
                  frmTipoCtaCte.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoCtaCte.RsCabecera.RecordCount = 0, 0, frmTipoCtaCte.RsCabecera.AbsolutePosition) & " de " & frmTipoCtaCte.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
                  cmdOpcionGrilla(0).Enabled = False
                  cmdOpcionGrilla(1).Enabled = False
                  cmdOpcionGrilla(2).Enabled = False
               End If
                                   
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub cmdOpcionGrilla_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
          
               'Cambiar los Controles
               With RsGrilla
                    'Cuadro de Texto
                    txtCodigoS.Text = ""
                    txtDetalladoS.Text = ""
                    txtResumidoS.Text = ""
                    'Check Box
                    chkActivoS.Value = 1
                End With
                SubDetalle False
                wAgrega = True
                txtDetalladoS.SetFocus
          
          Case Is = 1 ' Modificar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetalle False
               wAgrega = False
               SubAsignar
               txtDetalladoS.SetFocus
               
          Case Is = 2 ' Eliminar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               SubAsignar
               If MsgBox("Seguro de Eliminar el Sub Tipo " & txtCodigoS & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If Calcular("select count(tCodigoCliente) as Codigo From TCOMPANIA where tSubTipoCtaCte='" & txtCodigoS.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este SubTipo, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tCodigo ='" & txtCodigoS.Text & "' and TTABLA='SUBTIPOCTACTE'"
               frmTipoCtaCte.RsCabecera.Requery
               RsGrilla.Requery
               
               If RsGrilla.RecordCount <> 0 Then
                  RsGrilla.MoveLast
               End If
          
          Case Is = 3 ' Grabar
               If txtDetalladoS.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetalladoS.SetFocus: Exit Sub
               If txtResumidoS.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumidoS.SetFocus: Exit Sub
               
               If wAgrega Then
                  'Cambiar el SQL Agregar Detalle

                  nCorrela = Mid(Calcular("select max(tCodigo) as Codigo from TTABLA where tValor ='" & txtCodigo.Text & "' and TTABLA='SUBTIPOCTACTE'", Cn), 3, 2)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     txtCodigoS = Trim(txtCodigo.Text) & "01"
                  Else
                     txtCodigoS = Trim(txtCodigo.Text) & Lib.Correlativo(nCorrela, 2)
                  End If
                                      
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor,nValor, lActivo) " & _
                         "values (  'SUBTIPOCTACTE', " & _
                                  "'" & txtCodigoS.Text & "', " & _
                                  " '" & txtDetalladoS.Text & "', " & _
                                  " '" & txtResumidoS.Text & "', " & _
                                  " '" & txtCodigo.Text & "', '" & txtCuentaContable.Text & "'" & _
                                         chkActivoS.Value & ") "
                   Cn.Execute Isql
                   'Cambiar el Filtro de la Grilla
                   RsGrilla.Filter = "tTipoCtaCte ='" & txtCodigo.Text & "'"
                   RsGrilla.Requery
                   RsGrilla.MoveLast
                   MsgBox "Registro Agregado", vbInformation, sMensaje
                Else
                   ''Cambiar el SQL Modificar Detalle
                   Isql = "update TTABLA set " & _
                          "tDetallado ='" & txtDetalladoS.Text & "', " & _
                          "tResumido ='" & txtResumidoS.Text & "', " & _
                          "nValor = '" & txtCuentaContable.Text & "'," & _
                          "lActivo =" & chkActivoS.Value & " " & _
                          "where tCodigo = '" & txtCodigoS.Text & "' and tTABLA = 'SUBTIPOCTACTE'"
                    
                    Cn.Execute Isql
                    nPos = RsGrilla.AbsolutePosition
                    RsGrilla.Requery
                    RsGrilla.AbsolutePosition = nPos
                    MsgBox "Registro Modificado", vbInformation, sMensaje
                End If
                SubDetalle True
                txtCodigoS.Text = ""
          
          Case Is = 4 ' Cancelar
               SubDetalle True
               txtCodigoS.Text = ""
              
   End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    fraGrilla.Visible = False

    'Ingrese el Titulo
    Me.Caption = " Mantenimiento de Tipos de Cuentas Corrientes "
    fraDetalle.Caption = Me.Caption
            
    'Ingrese el SubTitulo
    grdGrilla.Caption = " Mantenimiento de Sub Tipos de Cuentas Corrientes "
    fraGrilla.Caption = grdGrilla.Caption
           
    'Ingresar la Vista de la Grilla
    'Isql = "select * from vSubTipoCtaCte " CSHM
    Isql = "select vstc.*,TT.nValor from vSubTipoCtaCte vstc INNER JOIN TTABLA TT on vstc.Codigo = TT.TCODIGO"
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(3, grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                  "Descripción", 2, "Descripcion", 4100, 0, 0, "", _
                                  "Activo", 2, "lActivo", 1000, 2, 4, "")
    Set grdGrilla.DataSource = RsGrilla
    If Sw = True Then
       Blanquear Me
       chkActivo.Value = 1
       'Cambiar Filtro de la Grilla
       RsGrilla.Filter = "Codigo ='" & txtCodigo.Text & "'"
       cmdOpcion(0).Enabled = False
       cmdOpcion(2).Enabled = False
       cmdOpcionGrilla(0).Enabled = False
       cmdOpcionGrilla(1).Enabled = False
       cmdOpcionGrilla(2).Enabled = False
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & frmTipoCtaCte.RsCabecera.AbsolutePosition & " de " & frmTipoCtaCte.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmTipoCtaCteDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
End Sub

Private Sub grdGrilla_DblClick()
   cmdOpcionGrilla_Click (1)
End Sub

Public Sub SubDetalle(Activa As Boolean)
   fraGrilla.Visible = Not Activa
   ActivarBotones Activa
   
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   
   cmdOpcionGrilla(0).Enabled = Activa
   cmdOpcionGrilla(1).Enabled = Activa
   cmdOpcionGrilla(2).Enabled = Activa
         
   'Controles de la cabecera
   txtDetallado.Enabled = Activa
   txtResumido.Enabled = Activa
   chkActivo.Enabled = Activa
   
End Sub

Sub SubAsignar()
    'Cambiar la Asignación del Detalle
    With RsGrilla
         'Cuadro de Texto
         txtCodigoS.Text = IIf(IsNull(!codigo), "", !codigo)
         txtDetalladoS.Text = IIf(IsNull(!Descripcion), "", !Descripcion)
         txtResumidoS.Text = IIf(IsNull(!tResumido), "", !tResumido)
         txtCuentaContable.Text = IIf(IsNull(!nValor), "", !nValor)
         chkActivoS.Value = IIf(IsNull(!lActivo), 1, IIf(!lActivo = True, 1, 0))
    End With
End Sub

