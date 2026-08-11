VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmReciboEgreso 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8610
   ClientLeft      =   -3525
   ClientTop       =   -1125
   ClientWidth     =   13785
   Icon            =   "frmReciboEgreso.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8581.01
   ScaleMode       =   0  'User
   ScaleWidth      =   13785
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1365
      ScaleWidth      =   13725
      TabIndex        =   19
      Top             =   7185
      Width           =   13785
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5805
         TabIndex        =   23
         Top             =   30
         Width           =   5865
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmReciboEgreso.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmReciboEgreso.frx":084C
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmReciboEgreso.frx":0D8E
            Style           =   1  'Graphical
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5250
            Picture         =   "frmReciboEgreso.frx":12D0
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4710
            Picture         =   "frmReciboEgreso.frx":1812
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4170
            Picture         =   "frmReciboEgreso.frx":1D54
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
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
            Left            =   1650
            TabIndex        =   24
            Top             =   150
            Width           =   2505
         End
      End
      Begin VB.CommandButton cmdOpcion 
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
         Left            =   5895
         Picture         =   "frmReciboEgreso.frx":2296
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Agrega"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Left            =   6480
         Picture         =   "frmReciboEgreso.frx":27C8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Modifica"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdEmite 
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
         Left            =   7065
         Picture         =   "frmReciboEgreso.frx":28CA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Emite"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
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
         Index           =   6
         Left            =   8820
         Picture         =   "frmReciboEgreso.frx":2DFC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   5880
         Picture         =   "frmReciboEgreso.frx":2EEE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Filtrar"
         Top             =   720
         Width           =   555
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas "
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
         Left            =   165
         TabIndex        =   20
         Top             =   650
         Width           =   5700
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   3240
            TabIndex        =   14
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   53018625
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   450
            TabIndex        =   12
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   53018625
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpHorIni 
            Height          =   345
            Left            =   1920
            TabIndex        =   13
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   53018627
            UpDown          =   -1  'True
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpHorFin 
            Height          =   345
            Left            =   4680
            TabIndex        =   15
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   53018627
            UpDown          =   -1  'True
            CurrentDate     =   37541.9993055556
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Del "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   315
            Width           =   405
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " al "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   21
            Top             =   270
            Width           =   225
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   555
         Index           =   5
         Left            =   8235
         Picture         =   "frmReciboEgreso.frx":2FF0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Reportes"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   7650
         Picture         =   "frmReciboEgreso.frx":3522
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exportar HTML"
         Top             =   90
         Width           =   555
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   7125
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   13695
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   6870
         Left            =   90
         TabIndex        =   17
         Top             =   180
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12118
         _LayoutType     =   4
         _RowHeight      =   21
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).FooterText=   "Footer Text"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "True/False"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).ScrollBars=   3
         Splits(0).AllowColMove=   -1  'True
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerStyle=   3
         Splits(0).DividerColor=   32768
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos(0).NeedTotalPageCount=   -1  'True
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Correlativo"
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DataView        =   2
         GroupByCaption  =   "Arrastre aqui la columna a agrupar"
         DeadAreaBackColor=   -2147483637
         ScrollTrack     =   -1  'True
         RowDividerColor =   12632256
         RowSubDividerColor=   -2147483648
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   16711935
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.bgcolor=&H80000018&,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=3,.fontname=Calibri"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=63,.parent=1,.namedParent=33"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=65,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=80,.parent=6,.bgcolor=&HE7FAB6&,.fgcolor=&H80000009&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=79,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=90,.parent=63"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=64,.alignment=0"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=65"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0,.bgcolor=&H80000018&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(38)  =   ":id=33,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=33,.fontname=Calibri"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H80000003&"
         _StyleDefs(42)  =   ":id=34,.fgcolor=&H800000&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
         _StyleDefs(43)  =   ":id=34,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=34,.fontname=Arial Black"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
         _StyleDefs(49)  =   ":id=36,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(50)  =   ":id=36,.fontname=Calibri"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&HFFFF&"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H800000&,.locked=0,.bold=-1"
         _StyleDefs(55)  =   ":id=38,.fontsize=2025,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=38,.fontname=Calibri"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(59)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(60)  =   ":id=39,.fontname=Calibri"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33,.bgcolor=&H80000005&,.borderColor=&H80000002&,.bold=0"
         _StyleDefs(63)  =   ":id=40,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(64)  =   ":id=40,.fontname=Calibri"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34,.bgcolor=&H80000003&,.fgcolor=&H80000009&,.bold=-1"
         _StyleDefs(67)  =   ":id=41,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=41,.fontname=Calibri"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33,.bgcolor=&H80000018&,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(71)  =   ":id=42,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(72)  =   ":id=42,.fontname=Calibri"
         _StyleDefs(73)  =   "Named:id=13:CabeceraTitulo"
         _StyleDefs(74)  =   ":id=13,.parent=34,.bold=-1,.fontsize=1125,.italic=0,.underline=0"
         _StyleDefs(75)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=13,.fontname=Arial Black"
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   315
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cmdSave1 
         Left            =   1305
         Top             =   405
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmReciboEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim RsReporte As Recordset
Dim Reporte As New dsrReciboEgreso
Dim ReporteCompra As New dsrReciboEgresoCompra
Dim nColumna As Integer
Dim col             As TrueOleDBGrid80.Column
Dim cols            As TrueOleDBGrid80.Columns
Dim lAsc            As Boolean

Private Sub cmdEmite_Click()
    With grdGrilla.PrintInfo
        'Cabecera
        .SettingsOrientation = 2
        .PageHeaderFont.Italic = True
        .RepeatColumnHeaders = True
        .PreviewMaximize = True
        .SettingsMarginLeft = 400
        .SettingsMarginRight = 400
        .SettingsMarginTop = 400
        .SettingsMarginBottom = 400
        .VariableRowHeight = True
        .PageFooter = "\tPage: \p"
        .PreviewCaption = "Emite de " & Me.Caption
        .PrintPreview
    End With
End Sub

Private Sub cmdExporta_Click()
    cmdSave.Filter = "Hyper Text Markup Language|*.html"
    cmdSave.ShowSave
    If cmdSave.FileName = "" Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    grdGrilla.ExportToFile cmdSave.FileName, False, False
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 2000 > 0, Me.Height - 2000, 0)
   fraGrilla.Width = IIf(Me.Width - 300 > 0, Me.Width - 300, 0)
   
   grdGrilla.Height = IIf(fraGrilla.Height - 300 > 0, fraGrilla.Height - 300, 0)
   grdGrilla.Width = IIf(fraGrilla.Width - 300 > 0, fraGrilla.Width - 300, 0)
End Sub




Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    nColumna = ColIndex
    lAsc = Not lAsc
    If lAsc Then
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    Else
       RsCabecera.Sort = grdGrilla.Columns(ColIndex).DataField & " DESC"
    End If
End Sub

Private Sub grdGrilla_FilterChange()
    On Error GoTo errHandler
    Set cols = grdGrilla.Columns
    Dim c As Integer
    c = grdGrilla.col
    grdGrilla.HoldFields
    RsCabecera.Filter = getFilter(cols)
    grdGrilla.col = c
    grdGrilla.EditActive = True
    Exit Sub

errHandler:
    MsgBox "Error: Caracter no valido para esta columna" & vbCrLf & Err.Description, vbCritical + vbOKOnly, sMensaje
    
    Call cmdClearFilter_Click(cols)
    RsCabecera.Filter = adFilterNone
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
           Case Is = 0  'Nuevo
                If lSolicitaClaveEgreso Then
                    If Supervisor("34") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                       
                        Exit Sub
                    End If
                End If
                Sw = True
                'Cambiar el Nombre del Formulario Detalle
                frmReciboEgresoDetalle.Show vbModal
           
           Case Is = 1  'Modifica
                If RsCabecera.RecordCount > 0 Then
                
                If lSolicitaClaveEgreso Then
                    If Supervisor("34") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                    End If
                End If
                
                    Sw = False
                    'Cambiar el Nombre del Formulario Detalle
                    frmReciboEgresoDetalle.Show vbModal
                Else
                    MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                End If
           
           Case Is = 5  'Emite
                If RsCabecera.RecordCount <= 0 Then MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema": Exit Sub
                
                frmFiltroRecibo.Show vbModal
                
                If sTipo = "Excel" Then
                    
                    Isql = " exec [spRep_ReciboEgreso] '" & IIf(lMCPV, sUsuario, "") & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "',1"
                    Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                    If RsReporte.RecordCount <= 0 Then MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema": Exit Sub
                    
                    RsReporte.MoveFirst
                    ExportaExcel RsReporte
                    
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf sTipo = "ExcelCompra" Then
                    Isql = " exec [spRep_ReciboEgreso] '" & IIf(lMCPV, sUsuario, "") & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "',2,'" & sAlmacenMDB & "'"
                    Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                    If RsReporte.RecordCount <= 0 Then MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema": Exit Sub
                    
                    RsReporte.MoveFirst
                    ExportaExcel RsReporte
                    
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf sTipo = "PantallaCompra" Then
                    Isql = " exec [spRep_ReciboEgreso] '" & IIf(lMCPV, sUsuario, "") & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "',2,'" & sAlmacenMDB & "'"
                    Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                    If RsReporte.RecordCount <= 0 Then MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema": Exit Sub
                    
                     'Set RsReporte = RsCabecera.Clone
                     'RsReporte.Sort = grdGrilla.Columns(nColumna).DataField & " ASC"
                    
                     ReporteCompra.DiscardSavedData
                     ReporteCompra.Database.SetDataSource RsReporte
                     ReporteCompra.Text8.SetText "Reporte de Recibos de Egreso"
                     ReporteCompra.Text16.SetText sRazonSocial
                     ReporteCompra.Text15.SetText localConectado
                     frmEmite.CRViewer.ReportSource = ReporteCompra
                     
                    frmEmite.CRViewer.DisplayGroupTree = False
                    ReporteCompra.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    ReporteCompra.PaperOrientation = crLandscape
                    frmEmite.CRViewer.ViewReport
                    frmEmite.Show vbModal
                     
                End If
                
                If Not wEnter Then
                   Exit Sub
                End If
                
                Set RsReporte = RsCabecera.Clone
                RsReporte.Sort = grdGrilla.Columns(nColumna).DataField & " ASC"
               
                Reporte.DiscardSavedData
                Reporte.Database.SetDataSource RsReporte
                Reporte.Text8.SetText "Reporte de Recibos de Egreso"
                Reporte.Text16.SetText sRazonSocial
                Reporte.Text15.SetText localConectado
                frmEmite.CRViewer.ReportSource = Reporte
                
                Screen.MousePointer = vbHourglass
                Select Case sTipo
                       Case "Pantalla"
                            frmEmite.CRViewer.DisplayGroupTree = False
                            Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            Reporte.PaperOrientation = crLandscape
                            frmEmite.CRViewer.ViewReport
                            frmEmite.Show vbModal
                            
                       Case "Impresora"
                            Screen.MousePointer = vbDefault
                            Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            Reporte.PaperOrientation = crPortrait
                            Reporte.PrintOut
                       Case "Excel"
'                            Screen.MousePointer = vbDefault
'                            Reporte.ExportOptions.FormatType = 21
'                            Reporte.ExportOptions.DestinationType = 1
'                            cmdSave1.Filter = "Libro de Microsoft Excel|*.xls"
'                            cmdSave1.ShowSave
'                            If cmdSave1.FileName = "" Then
'                               Exit Sub
'                            End If
'                            Reporte.ExportOptions.DiskFileName = cmdSave1.FileName
'                            Reporte.Export False
                End Select
                 Screen.MousePointer = vbDefault
                 
           Case Is = 6  'Salir
                Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    cmdOpcion(0).Enabled = IIf(wInicio, True, False)
    cmdOpcion(1).Enabled = IIf(wInicio, True, False)
    
    ' Ingrese el Titulo
    Me.Caption = " Mantenimiento de Recibos de Egreso "
    grdGrilla.Caption = Me.Caption

    Sw = False
    Centrar Me
    nColumna = 0
    dtpFecIni.value = Date - nDias
    dtpFecFin.value = Date
    
    ' Ingrese la Vista
'    If lMCPV Then
'       Isql = "select tRecibo, tTurno, fFecha, tMoneda, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro, Estado, Moneda , TipoEgreso,ttipoegreso,TipoIdentificacion,IdentificacionResumida,Identificacion,RazonSocial,NComprobante,TipoComprobante,CentroCosto,CCostoResumida, str(nMonto,10,2) as xMonto, convert(varchar, fFecha, 103) as xFecha, dReferencia, tModoEgreso from vEgreso " & _
'              "where tUsuario='" & sUsuario & "' and fregistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fregistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "'"
'    Else
'       Isql = "select tRecibo, tTurno, fFecha, tMoneda, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro, Estado, Moneda, TipoEgreso,ttipoegreso,TipoIdentificacion,IdentificacionResumida,Identificacion,RazonSocial,NComprobante,TipoComprobante,CentroCosto,CCostoResumida, str(nMonto,10,2) as xMonto, convert(varchar, fFecha, 103) as xFecha, dReferencia, tModoEgreso from vEgreso " & _
'              "where fregistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fregistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "'"
'    End If
    
    Isql = " exec [spRep_ReciboEgreso] '" & IIf(lMCPV, sUsuario, "") & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "',1"
    
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'Configuración de la Grilla
    Call ConfGrilla(12, grdGrilla, "Recibo", 2, "tRecibo", 1200, 0, 0, "", _
                                  "Fecha", 2, "xFecha", 1200, 0, 0, "dd/mm/yyyy", _
                                  "Tipo Egreso", 2, "tipoEgreso", 1600, 0, 0, "", _
                                  "Mon", 2, "Moneda", 550, 2, 0, "", _
                                  "Monto", 2, "xMonto", 1100, 1, 0, "###,###,##0.00", _
                                  "Descripción", 2, "tDescripcion", 3000, 0, 0, "", _
                                  "Autoriza", 2, "tAutoriza", 1600, 0, 0, "", _
                                  "Estado", 2, "Estado", 1200, 0, 0, "", _
                                  "Doc. Ref", 2, "dReferencia", 1800, 0, 0, "", _
                                  "Forma", 2, "tModoEgreso", 1800, 0, 0, "", _
                                  "Turno", 2, "tTurno", 1200, 0, 0, "", _
                                  "Observacion", 2, "tObservacion", 5000, 0, 0, "")
                                    
    Set grdGrilla.DataSource = RsCabecera
    
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsReporte = Nothing
    Set RsCabecera = Nothing
    Set frmReciboEgreso = Nothing
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub
Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub cmdProcesa_Click()
   Screen.MousePointer = vbHourglass

'
'    If lMCPV Then
'       Isql = "select tRecibo, tTurno, fFecha, tMoneda, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro, Estado, Moneda , TipoEgreso,ttipoegreso,TipoIdentificacion,IdentificacionResumida,Identificacion,RazonSocial,NComprobante,TipoComprobante,CentroCosto,CCostoResumida, str(nMonto,10,2) as xMonto, convert(varchar, fFecha, 103) as xFecha, dReferencia from vEgreso " & _
'              "where tUsuario='" & sUsuario & "' and fregistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fregistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "'"
'    Else
'       Isql = "select tRecibo, tTurno, fFecha, tMoneda, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro, Estado, Moneda, TipoEgreso,ttipoegreso,TipoIdentificacion,IdentificacionResumida,Identificacion,RazonSocial,NComprobante,TipoComprobante,CentroCosto,CCostoResumida, str(nMonto,10,2) as xMonto, convert(varchar, fFecha, 103) as xFecha, dReferencia from vEgreso " & _
'              "where fregistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and fregistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "'"
'    End If
    
 Isql = " exec [spRep_ReciboEgreso] '" & IIf(lMCPV, sUsuario, "") & "','" & Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "','" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(Me.dtpHorFin.value, "HH:mm") & "',1"
    
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   Set grdGrilla.DataSource = RsCabecera
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
End Sub

