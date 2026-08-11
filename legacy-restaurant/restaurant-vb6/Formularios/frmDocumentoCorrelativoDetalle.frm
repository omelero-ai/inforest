VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDocumentoCorrelativoDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7050
   ClientLeft      =   -660
   ClientTop       =   1815
   ClientWidth     =   11910
   Icon            =   "frmDocumentoCorrelativoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11910
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
      Height          =   5775
      Left            =   1560
      TabIndex        =   9
      Top             =   0
      Width           =   9975
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   2670
         Left            =   90
         TabIndex        =   33
         Top             =   930
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   4710
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
         Caption         =   "SubTitulo"
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
      Begin VB.Label lblReimpresion 
         BackColor       =   &H0080C0FF&
         Caption         =   "Total de reimpresiones:"
         Height          =   255
         Left            =   3960
         TabIndex        =   36
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label lblLey1 
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
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1260
         TabIndex        =   35
         Top             =   3960
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ley1 :"
         Height          =   195
         Index           =   11
         Left            =   750
         TabIndex        =   34
         Top             =   4005
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Recargos :"
         Height          =   195
         Index           =   3
         Left            =   4065
         TabIndex        =   31
         Top             =   4020
         Width           =   780
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descuentos :"
         Height          =   195
         Index           =   4
         Left            =   3900
         TabIndex        =   30
         Top             =   4335
         Width           =   945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Index           =   5
         Left            =   4395
         TabIndex        =   29
         Top             =   5265
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Precio Oficial :"
         Height          =   195
         Index           =   2
         Left            =   3825
         TabIndex        =   28
         Top             =   3705
         Width           =   1020
      End
      Begin VB.Label txtFecha 
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
         Left            =   1155
         TabIndex        =   27
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label txtCliente 
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
         Left            =   1155
         TabIndex        =   26
         Top             =   585
         Width           =   8610
      End
      Begin VB.Label txtRecargo 
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
         Left            =   4980
         TabIndex        =   25
         Top             =   3975
         Width           =   1545
      End
      Begin VB.Label txtDescuento 
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
         Left            =   4980
         TabIndex        =   24
         Top             =   4290
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
         Left            =   4995
         TabIndex        =   23
         Top             =   5220
         Width           =   1545
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
         Left            =   4980
         TabIndex        =   22
         Top             =   4590
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Neto :"
         Height          =   195
         Index           =   6
         Left            =   4410
         TabIndex        =   21
         Top             =   4635
         Width           =   435
      End
      Begin VB.Label txtImpuesto 
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
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   4980
         TabIndex        =   20
         Top             =   4905
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Impuestos / Otros :"
         Height          =   195
         Index           =   7
         Left            =   3495
         TabIndex        =   19
         Top             =   4950
         Width           =   1350
      End
      Begin VB.Label txtImpuesto1 
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
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1260
         TabIndex        =   18
         Top             =   4290
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Impuesto1 :"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   17
         Top             =   4335
         Width           =   825
      End
      Begin VB.Label txtImpuesto2 
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
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1260
         TabIndex        =   16
         Top             =   4590
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Impuesto2 :"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   15
         Top             =   4635
         Width           =   825
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Impuesto3 :"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   14
         Top             =   4950
         Width           =   825
      End
      Begin VB.Label txtImpuesto3 
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
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   1260
         TabIndex        =   13
         Top             =   4905
         Width           =   1545
      End
      Begin VB.Label txtOficial 
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
         Left            =   4980
         TabIndex        =   12
         Top             =   3660
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   11
         Top             =   630
         Width           =   570
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   0
         Left            =   495
         TabIndex        =   10
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11850
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6300
      Width           =   11910
      Begin VB.CommandButton cmdSalir 
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
         Left            =   10260
         Picture         =   "frmDocumentoCorrelativoDetalle.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   10080
         TabIndex        =   1
         Top             =   30
         Width           =   10140
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   8445
            Picture         =   "frmDocumentoCorrelativoDetalle.frx":00FE
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   8985
            Picture         =   "frmDocumentoCorrelativoDetalle.frx":0640
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   9525
            Picture         =   "frmDocumentoCorrelativoDetalle.frx":0B82
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmDocumentoCorrelativoDetalle.frx":10C4
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmDocumentoCorrelativoDetalle.frx":1606
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmDocumentoCorrelativoDetalle.frx":1B48
            Style           =   1  'Graphical
            TabIndex        =   2
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
            Height          =   255
            Left            =   3480
            TabIndex        =   32
            Top             =   180
            Width           =   3345
         End
      End
   End
   Begin VB.Image Image 
      Height          =   5670
      Left            =   30
      Picture         =   "frmDocumentoCorrelativoDetalle.frx":208A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1545
   End
End
Attribute VB_Name = "frmDocumentoCorrelativoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrilla As Recordset

Sub Asignar()
    'Asigna la Cabecera
    If sModulo = "PVCORP" Then
       Isql = "select * from vDocumentoCorrelativoDetalle where tDocumento ='" & frmDocumentoCorrelativo.grdGrilla.Columns(1).value & "'"
    Else
       Isql = "select * from vDocumentoCorrelativoDetalle where tDocumento ='" & frmDocumentoCorrelativo.grdGrilla.Columns(2).value & "'"
    End If
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
    Set grdGrilla.DataSource = RsGrilla
    
    Dim vley1 As Double
    vley1 = Calcular("select isnull(nLey1,0) as codigo from mdocumento where tDocumento ='" & frmDocumentoCorrelativo.grdGrilla.Columns(2).value & "'", Cn)
    
    With RsGrilla
        'Cuadro de Texto
        sDocumento = !tDocumento
        txtFecha.Caption = IIf(IsNull(!fRegistro), "", !fRegistro)
        txtCliente.Caption = IIf(IsNull(!Cliente), "", !Cliente)
        txtNeto.Caption = Format(IIf(IsNull(!nNeto), 0, !nNeto), "###,###,##0.00")
        txtRecargo.Caption = Format(IIf(IsNull(!nRecargo), 0, !nRecargo), "###,###,##0.00")
        txtDescuento.Caption = Format(IIf(IsNull(!nDescuento), 0, !nDescuento), "###,###,##0.00")
        txtOficial.Caption = Format(IIf(IsNull(!nPrecioOficial), 0, !nPrecioOficial), "###,###,##0.00")
        txtImpuesto.Caption = 0
        lblLey1.Caption = Format(IIf(IsNull(vley1), 0, vley1), "###,###,##0.00")
        
        If Not IsNull(!nprecioImpuesto1) And !nprecioImpuesto1 > 0 Then
           txtImpuesto1.Caption = Format(!nprecioImpuesto1, "###,###,##0.00")
           Label(8).Caption = sImpuesto1
           txtImpuesto.Caption = CDbl(txtImpuesto.Caption) + !nprecioImpuesto1
           txtImpuesto1.Visible = True
           Label(8).Visible = True
        Else
           txtImpuesto1.Caption = Format(0, "###,###,##0.00")
           txtImpuesto1.Visible = False
           Label(8).Visible = False
        End If
        
        If Not IsNull(!nprecioImpuesto2) And !nprecioImpuesto2 > 0 Then
           txtImpuesto2.Caption = Format(!nprecioImpuesto2, "###,###,##0.00")
           Label(9).Caption = sImpuesto2
           txtImpuesto.Caption = CDbl(txtImpuesto.Caption) + !nprecioImpuesto2
           txtImpuesto2.Visible = True
           Label(9).Visible = True
        Else
           txtImpuesto2.Caption = Format(0, "###,###,##0.00")
           txtImpuesto2.Visible = False
           Label(9).Visible = False
        End If
        
        If Not IsNull(!nprecioImpuesto3) And !nprecioImpuesto3 > 0 Then
           txtImpuesto3.Caption = Format(!nprecioImpuesto3, "###,###,##0.00")
           Label(10).Caption = sImpuesto3
           txtImpuesto.Caption = CDbl(txtImpuesto.Caption) + !nprecioImpuesto3
           txtImpuesto3.Visible = True
           Label(10).Visible = True
        Else
           txtImpuesto3.Caption = Format(0, "###,###,##0.00")
           txtImpuesto3.Visible = False
           Label(10).Visible = False
        End If
        
        txtTotal.Caption = Format(IIf(IsNull(!nVenta), 0, !nVenta), "###,###,##0.00")
        fraDetalle.Caption = " Documento " & Format(sDocumento, "@-@@@@@-@@@@@@@@@")
        
        Dim CantDoc As Double
        Dim DocUser As String
        CantDoc = Calcular("select isnull(nReimpresion,0) as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
        DocUser = Calcular("select isnull(UsuarioReimpresion,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
        
        Me.lblReimpresion.Caption = "Total de Reimpresiones: " & CantDoc & ", Ultimo usuario de reimpresion: " & DocUser
        
    End With
    txtImpuesto.Caption = Format(CDbl(txtImpuesto.Caption), "###,###,##0.00")
    cmdTexto.Caption = "Registro " & frmDocumentoCorrelativo.RsCabecera.AbsolutePosition & " de " & frmDocumentoCorrelativo.RsCabecera.RecordCount
    'Exit Sub
On Error GoTo fin:
 If txtImpuesto2.Caption = "" Then
        If (CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) > CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)) Then
            txtImpuesto1.Caption = CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) - ((CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) - CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)))
        ElseIf (CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) < CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)) Then
            txtImpuesto1.Caption = Format(CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) - (CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) + CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)), "###,###,##0.00")
        End If
    Else
        If (CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) > CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)) Then
            txtImpuesto2.Caption = CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) - ((CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) - CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)))
        ElseIf (CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) < CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)) Then
            txtImpuesto2.Caption = Format(CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) - (CDbl(IIf(txtImpuesto1.Caption = "", 0, txtImpuesto1.Caption)) + CDbl(IIf(txtImpuesto2.Caption = "", 0, txtImpuesto2.Caption)) + CDbl(IIf(txtImpuesto3.Caption = "", 0, txtImpuesto3.Caption))) + CDbl(IIf(txtImpuesto.Caption = "", 0, txtImpuesto.Caption)), "###,###,##0.00")
        End If
    End If
    
   txtImpuesto.Caption = Format(CDbl(txtImpuesto.Caption) + vley1, "###,###,##0.00")
fin:

End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmDocumentoCorrelativo.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmDocumentoCorrelativo.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmDocumentoCorrelativo.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmDocumentoCorrelativo.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmDocumentoCorrelativo.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmDocumentoCorrelativo.grdGrilla
    End Select
   Asignar
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el Titulo
    Me.Caption = " Visualización de Documentos "
    
    'Ingrese el SubTitulo
    grdGrilla.Caption = " Detalle del Documento "
    Isql = "select * from vDocumentoCorrelativoDetalle where tDocumento ='" & sDocumento & "'"
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
    
    If sModulo = "PVCORP" Then
       Call ConfGrilla(7, grdGrilla, "Item", 2, "tItem", 600, 2, 0, "", _
                                     "Producto", 2, "Producto", 2200, 0, 0, "", _
                                     "P.Venta", 2, "nPrecioVenta", 1050, 1, 0, "###,###,##0.00", _
                                     "Cantidad", 2, "nCantidad", 1050, 1, 0, "##,##0.00", _
                                     "SubTotal", 2, "Venta", 1050, 1, 0, "###,###,##0.00", _
                                     "Pedido", 2, "tCodigoPedido", 1100, 0, 0, "", _
                                     "Guia Remision", 2, "tObservacion", 1600, 0, 0, "")
       
    Else
       Call ConfGrilla(9, grdGrilla, "Item", 2, "tItem", 500, 2, 0, "", _
                                     "Producto", 2, "Producto", 2200, 0, 0, "", _
                                     IIf(sImpuesto1 = "", "no def", sImpuesto1), 2, "nImpuesto1", 800, 1, 0, "###,###,##0.00", _
                                     IIf(sImpuesto2 = "", "no def", sImpuesto2), 2, "nImpuesto2", 800, 1, 0, "###,###,##0.00", _
                                     IIf(sImpuesto3 = "", "no def", sImpuesto3), 2, "nImpuesto3", 800, 1, 0, "###,###,##0.00", _
                                     "P.Venta", 2, "nPrecioVenta", 1000, 1, 0, "###,###,##0.00", _
                                     "Cantidad", 2, "nCantidad", 800, 1, 0, "##,##0.00", _
                                     "SubTotal", 2, "Venta", 1100, 1, 0, "###,###,##0.00", _
                                     "Cod.Pedido", 2, "tCodigoPedido", 1100, 2, 0, "")
    End If
    
    Set grdGrilla.DataSource = RsGrilla
    
    If sLey1 <> "" Then
        Label(11).Visible = True
        lblLey1.Visible = True
        Label(11).Caption = sLey1 & " :"
    End If
    
    Asignar
    cmdTexto.Caption = "Registro " & frmDocumentoCorrelativo.RsCabecera.AbsolutePosition & " de " & frmDocumentoCorrelativo.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsGrilla = Nothing
    Set frmDocumentoCorrelativoDetalle = Nothing
End Sub


