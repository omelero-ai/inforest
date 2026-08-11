VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmDespachador 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8400
   ClientLeft      =   2535
   ClientTop       =   1725
   ClientWidth     =   14160
   Icon            =   "frmDespachador.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   14220
   ScaleWidth      =   28680
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   13920
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   28620
      TabIndex        =   2
      Top             =   13470
      Width           =   28680
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   7485
         Picture         =   "frmDespachador.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Exportar HTML"
         Top             =   90
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
         Left            =   9765
         TabIndex        =   19
         Top             =   0
         Width           =   3525
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   2070
            TabIndex        =   20
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
            Format          =   54460417
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   405
            TabIndex        =   21
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
            Format          =   54460417
            CurrentDate     =   37539
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
            Left            =   1755
            TabIndex        =   23
            Top             =   270
            Width           =   225
            WordWrap        =   -1  'True
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
            Left            =   -45
            TabIndex        =   22
            Top             =   315
            Width           =   405
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdProcesa 
         Height          =   555
         Left            =   13320
         Picture         =   "frmDespachador.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Filtrar"
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
         Left            =   9150
         Picture         =   "frmDespachador.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salir"
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
         Left            =   6930
         Picture         =   "frmDespachador.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Emite"
         Top             =   90
         Width           =   555
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5085
         TabIndex        =   8
         Top             =   30
         Width           =   5145
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3450
            Picture         =   "frmDespachador.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3990
            Picture         =   "frmDespachador.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4530
            Picture         =   "frmDespachador.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmDespachador.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmDespachador.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmDespachador.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   9
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
            TabIndex        =   15
            Top             =   150
            Width           =   1740
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
         Index           =   10
         Left            =   5265
         Picture         =   "frmDespachador.frx":3026
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ver Detalle"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "No Emp"
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
         Index           =   9
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Desasigna Empacador"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Emp"
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
         Index           =   8
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Asigna Empacador"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "No Mot"
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
         Left            =   8595
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desasigna Motorizado"
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Mot"
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
         Left            =   6375
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Asigna Motorizado"
         Top             =   90
         Width           =   555
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   7620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13875
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   7365
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   13680
         _ExtentX        =   24130
         _ExtentY        =   12991
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
         Left            =   450
         Top             =   495
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmDespachador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim RsReporte As Recordset
Dim RsCodigoPCD As ADODB.Recordset
Dim nColumna As Integer
Dim xPedido As String
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

Private Sub cmdProcesa_Click()
   Screen.MousePointer = vbHourglass
    Isql = "select * from vDespachador " & _
           "where tTipoPedido='02' and tEstadoPedido='02' and isnull(fLlegada,0)=0 and fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' order by lempacador,ffecha asc"
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   Set grdGrilla.DataSource = RsCabecera
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
   Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
   fraGrilla.Height = IIf(Me.Height - 1300 > 0, Me.Height - 1300, 0)
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
    
         Case Is = 0  'Asignar
                
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                Dim nMaximoAsignacionMotorizado As Double
                Dim nCantidadPedidosYaAsignados As Integer
                Dim nCantidadPedidosYaAsignados1 As Integer
                Dim nDiaActual As Integer
                Dim nTarifaDiaria As Double
                Dim nMontoAsignado As Double
                Dim nMaximo As Boolean
                Dim MaxMotorizado As Double
                
                nMaximoAsignacionMotorizado = Calcular("select isnull(nAsignacionMotorizado,0) as Codigo FROM TPARAMETRO", Cn)
                MaxMotorizado = Calcular("select case when isnull(tMaxMotorizado,0)='' then 0 else  isnull(tMaxMotorizado,0) end   as Codigo FROM TPARAMETRO", Cn)
                
                If IIf(IsNull(RsCabecera!Empacador), "", Trim(RsCabecera!Empacador)) <> "" Then
                   If IsNull(RsCabecera!fAsignacion) Then
                
                   sTemp = ""
                   Isql = "select * from vMotorizado where lActivo = 1 and Codigo <> '0000' order by Descripcion"
                   Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 2, "Codigo", 1800, 2, 0, "", _
                                                                   "Motorizado", 2, "Descripcion", 6400, 0, 0, "")
                    
                   frmBusquedaRapida.nPredeterm = 1
                   frmBusquedaRapida.Show vbModal
    
                   If wEnter Then
                   
                            If nMaximoAsignacionMotorizado > 0 Then
                                
                                nCantidadPedidosYaAsignados = Calcular("SELECT COUNT(ISNULL(nTarifaMotorizado,0)) AS Codigo FROM dbo.vMotorizado INNER JOIN dbo.MPEDIDO ON dbo.vMotorizado.Codigo = dbo.MPEDIDO.tMotorizado WHERE  ISNULL(dbo.MPEDIDO.nTarifaExtra,0)=0 AND (CONVERT(nvarchar, dbo.MPEDIDO.fAsignacion, 112) = '" & Format(obtieneFechaServidor, "yyyyMMdd") & "') GROUP BY dbo.MPEDIDO.tMotorizado HAVING (dbo.MPEDIDO.tMotorizado = '" & sCodigo & "')  ", Cn)
                                nCantidadPedidosYaAsignados1 = Calcular("SELECT COUNT(ISNULL(nTarifaMotorizado,0)) AS Codigo FROM dbo.vMotorizado INNER JOIN dbo.MPEDIDO ON dbo.vMotorizado.Codigo = dbo.MPEDIDO.tMotorizadoN WHERE ISNULL(dbo.MPEDIDO.nTarifaExtraN,0)=0 AND (CONVERT(nvarchar, dbo.MPEDIDO.fAsignacion, 112) = '" & Format(obtieneFechaServidor, "yyyyMMdd") & "') GROUP BY dbo.MPEDIDO.tMotorizadoN HAVING (dbo.MPEDIDO.tMotorizadoN = '" & sCodigo & "')  ", Cn)
                                
                                Select Case tTarifaActualMotorizado
                                    Case "Tarifa Lun - Sab"
                                                nTarifaDiaria = Calcular("select isnull(ntarifalv,0) as codigo  from tmotorizadodatos where tcodigo='" & sCodigo & "'", Cn)

                                    Case "Tarifa Dom"
                                                nTarifaDiaria = Calcular("select isnull(ntarifasd,0) as codigo  from tmotorizadodatos where tcodigo='" & sCodigo & "'", Cn)

                                    Case "Tarifa Especial"
                                    
                                                nTarifaDiaria = Calcular("select isnull(ntarifaes,0) as codigo  from tmotorizadodatos where tcodigo='" & sCodigo & "'", Cn)
                                    
                                    Case Else
                                                nTarifaDiaria = 0
                                End Select
                          
                                
                                
                                nMontoAsignado = Round((nTarifaDiaria * (nCantidadPedidosYaAsignados + nCantidadPedidosYaAsignados1)) + nTarifaDiaria, 2)
                                
                                        ' Validar Tarifa Extra
                                        If nMaximoAsignacionMotorizado < nMontoAsignado Then
                                        
                                            If MsgBox("Con la Asignacion de este Pedido, sobre pasa el Monto Máximo de Asignación Por Motorizado, Desea Continuar?", vbQuestion + vbYesNo) = vbYes Then

                                                   nMaximo = True
                                                
                                            Else
                                                Exit Sub
                    
                                            End If
                                                                                            
                                        End If
                                        
                            End If
                        If MaxMotorizado > 0 Then
                            Dim con As Long
                            
                            Isql = "select count(*) as Codigo from vDespachador " & _
                            "where tTipoPedido='02' and tEstadoPedido='02' and isnull(fLlegada,0)=0 and  tmotorizado='" & sCodigo & "'"
                            con = Calcular(Isql, Cn)
                            If MaxMotorizado < con + 1 Then
                                MsgBox "Ha alcanzado el tope Maximo de Asignacion de motorizado: " & MaxMotorizado & " Cada Motorizado", vbInformation, sMensaje
                                Exit Sub
                            End If
                        End If

                      'Tarifa Oficial
                      Cn.Execute "update TTABLA set ntamano=1,lreplica=1 where tTABLA='MOTORIZADO' and tCodigo='" & sCodigo & "'"
                      Cn.Execute "update MPEDIDO set nTarifaMotorizado=" & IIf(IsNull(nTarifaDiaria), 0, nTarifaDiaria) & ", nTarifaExtra=0, fSalida=getdate(), fAsignacion = getdate(), tMotorizado='" & sCodigo & "' where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                      
                      'Tarifa Extra
                      If nMaximo = True Then
                         Cn.Execute "update MPEDIDO set nTarifaExtra=1 where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                      End If
                      
                      
                      
                      If CD = True Or WEBAP = True Then
                            
                            Dim xTiporecepcion As Integer
                            
                            xTiporecepcion = Calcular("Select tTiporecepcion as codigo from MPEDIDO Where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'", Cn)
                            
                            If xTiporecepcion = 1 Then
                                    Isql = "Select tCodigoPedidoCD From MPEDIDO WHERE tCodigoPedido = '" & RsCabecera!tCodigoPedido & "'"
                                    Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
                                    If Not RsCodigoPCD.EOF Then
                                      If Not IsNull(RsCodigoPCD!tCodigoPedidoCD) Then
                                          Call ModifcarEstadoDeliveryCabecera(RsCodigoPCD!tCodigoPedidoCD, "8", sCodigo, sDescrip)
                                      End If
                                    End If
                            ElseIf xTiporecepcion = 2 Then
                                'Cn.Execute "update MPEDIDO set tCodigoPedidoWebAp=1 where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
'                                    Isql = "Select tCodigoPedidoWebAp From MPEDIDO WHERE tCodigoPedido = '" & RsCabecera!tCodigoPedido & "'"
'                                    Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
'                                    If Not RsCodigoPCD.EOF Then
'                                      If Not IsNull(RsCodigoPCD!tCodigoPedidoWebAp) Then
'                                          Call ModificarEstadoPedidoWebAp(RsCodigoPCD!tCodigoPedidoWebAp)
'                                      End If
'                                    End If
                            End If
                      End If


                                                                                        
                      xPedido = RsCabecera!tCodigoPedido
                      RsCabecera.Requery
                      RsCabecera.MoveFirst
                      RsCabecera.Find "tCodigoPedido='" & xPedido & "'"

                   End If
                      
                      
                   Else
                      MsgBox "Error: El Pedido ya fue asignado", vbCritical, sMensaje
                   End If
                Else
                   MsgBox "Error: Debe asignar primero el empacador", vbCritical, sMensaje
                End If
                
         Case Is = 1  'Desasignar

                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                
                If Not IsNull(RsCabecera!fAsignacion) Then
                   If MsgBox("Seguro de Desasignar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                   
                      'nPos = RsCabecera.Bookmark
                      
                      If Calcular("select count(tMotorizado) as codigo from MPEDIDO where tTipoPedido='02' and tEstadoPedido='02' and isnull(fLlegada,0)=0 and tMotorizado='" & RsCabecera!tMotorizado & "' and isnull(fAsignacion,0)<>0", Cn) = 1 Then
                         Cn.Execute "update TTABLA set ntamano = 0 ,lreplica=1 where tTABLA='MOTORIZADO' and tCodigo='" & RsCabecera!tMotorizado & "'"
                      End If
                      
                      Cn.Execute "update MPEDIDO set  nTarifaMotorizado=null, nTarifaExtra = null , fAsignacion = null, fSalida=null, tMotorizado='0000' where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                      
                      xPedido = RsCabecera!tCodigoPedido
                      RsCabecera.Requery
                      RsCabecera.MoveFirst
                      RsCabecera.Find "tCodigoPedido='" & xPedido & "'"
                      'RsCabecera.Bookmark = nPos
                      
                      'Oscar Ortega-Central Delivery
                      If CD Then
                         Isql = "Select tCodigoPedidoCD From MPEDIDO WHERE tCodigoPedido = '" + RsCabecera!tCodigoPedido + "'"
                         'Dim RsCodigoPCD As ADODB.Recordset
                         Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
                         If Not RsCodigoPCD.EOF Then
                            If Not IsNull(RsCodigoPCD!tCodigoPedidoCD) Then
                               Call ModifcarEstadoDeliveryCabecera(RsCodigoPCD!tCodigoPedidoCD, "8", "", "")
                            End If
                         End If
                      End If
                      ''''''''''''''''''''''''''''''
                   End If
                Else
                   MsgBox "Error: El Pedido no esta asignado", vbCritical, sMensaje
                End If
           
         Case Is = 5  'Refrescar
                Screen.MousePointer = vbHourglass
'                Cn.Execute "update MPEDIDO set lEstado=0 where isnull(fAsignacion,0)=0"
                RsCabecera.Requery
                If Not RsCabecera.EOF Then
                   RsCabecera.MoveLast
                End If
                Screen.MousePointer = vbDefault
                
           Case Is = 6  'Salir
'                Cn.Execute "update MPEDIDO set lEstado=0 where isnull(fAsignacion,0)=0"
                Unload Me
                           
           Case Is = 8  'Asignar Empacador
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If RsCabecera!Empacador <> "" Then
                   MsgBox "Error: El Pedido ya fue Empacado", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                sTemp = ""
                Isql = "select * from vEmpacador where lActivo = 1 Order by Descripcion"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                   'nPos = RsCabecera.Bookmark
                   Cn.Execute "update MPEDIDO set tEmpacador='" & sCodigo & "', fEmpacador=getdate() where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                   xPedido = RsCabecera!tCodigoPedido
                   RsCabecera.Requery
                   RsCabecera.MoveFirst
                   RsCabecera.Find "tCodigoPedido='" & xPedido & "'"
                   'RsCabecera.Bookmark = nPos
                End If
                
           Case Is = 9  'Desasignar Empacador
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If (IsNull(RsCabecera!Empacador) Or RsCabecera!Empacador = "") Then
                   MsgBox "Error: El Pedido no esta empaquetado", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                If Not IsNull(RsCabecera!Motorizado) And RsCabecera!Motorizado <> "" And RsCabecera!tMotorizado <> "0000" Then
                   MsgBox "Error: El Pedido ya esta en camino", vbExclamation, sMensaje
                   Exit Sub
                End If
                If MsgBox("Seguro de Desasignar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                   'nPos = RsCabecera.Bookmark
                   Cn.Execute "update MPEDIDO set tEmpacador='', fEmpacador=null  where tCodigoPedido='" & RsCabecera!tCodigoPedido & "'"
                   xPedido = RsCabecera!tCodigoPedido
                   RsCabecera.Requery
                   RsCabecera.MoveFirst
                   RsCabecera.Find "tCodigoPedido='" & xPedido & "'"
                   
                   'RsCabecera.Requery
                   'RsCabecera.Bookmark = nPos
                End If
                
           Case Is = 10  'Detalle
           If Not (RsCabecera.EOF Or RsCabecera.BOF) Then
                sPedido = RsCabecera!tCodigoPedido
                frmPedido.cmdOpcion(5).Enabled = False
                frmPedido.cmdOpcion(6).Enabled = False
                frmPedido.cmdOpcion(7).Enabled = False
                frmPedido.Show
        End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    dtpFecIni.Value = FechaServidor() - nDias
    dtpFecFin.Value = FechaServidor()

    ' Ingrese el Titulo
    Me.Caption = " Despachador "
    grdGrilla.Caption = Me.Caption
    nColumna = 2
    
    ' Ingrese la Vista
    'EDLCQ
'    Isql = "select * from vDespachador " & _
'           "where tTipoPedido in (select Codigo from vTipoPedido where lActivo =1 and lCanalDelivery=1 ) and CodigoOrigenVenta in (select CodOrigenVenta from vOrigenVenta where lVisualizaDespachador =1 and Activo =1)  and tEstadoPedido='02' and isnull(fLlegada,0)=0 and fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' order by lempacador,ffecha asc"
'    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
     Isql = "select * from vDespachador " & _
           "where tTipoPedido in (select Codigo from vTipoPedido where lActivo =1 and lCanalDelivery=1 ) and tEstadoPedido='02' and isnull(fLlegada,0)=0 and fFecha >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " 00:00' and fFecha <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " 23:59' order by lempacador,ffecha asc"
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    'Configuración de la Grilla
    Call ConfGrilla(11, grdGrilla, "Pedido", 2, "tCodigoPedido", 1200, 2, 0, "", _
                                   "Fecha", 2, "fFecha", 1200, 0, 0, "dd/MM HH:mm", _
                                   "Usuario", 2, "tUsuario", 1000, 0, 0, "", _
                                   "Caja", 2, "tCaja", 800, 2, 0, "", _
                                   "Cliente", 2, "Cliente", 2600, 0, 0, "", _
                                   "Telefono", 2, "tTelefono", 1500, 0, 0, "", _
                                   "Direccion", 2, "tDireccion", 3500, 0, 0, "", _
                                   "Zona", 2, "Referencia", 1500, 0, 0, "", _
                                   "Empacador", 2, "Empacador", 1500, 0, 0, "", _
                                   "Motorizado", 2, "Motorizado", 1500, 0, 0, "", _
                                   "H.Asigna", 2, "fAsignacion", 1200, 0, 0, "dd/MM HH:mm")
                                       
    Set grdGrilla.DataSource = RsCabecera
    cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmDespachador = Nothing
End Sub

Private Sub grdGrilla_DblClick()
   cmdOpcion_Click (10)
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub

Private Sub txtCriterio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOpcion_Click (2)
   End If
End Sub

Private Sub ModifcarEstadoDeliveryCabecera(CodigoPedido As String, tEstadoDelivery As String, Optional sCodigoMotorizado As String, Optional sNombremotorizado As String)
    Dim CnCD As Connection
    'Configuración
    
    Set CnCD = New Connection
    CnCD.Provider = "SQLOLEDB"
    CnCD.CursorLocation = adUseServer
    CnCD.ConnectionString = "User ID=" & sUserName & _
                         ";password=" & sUserPassword & _
                         ";Data Source=" & sRutaCD & _
                         ";Initial Catalog=" & sMDBCD
    CnCD.CommandTimeout = 250
    CnCD.Open
    
    'CnCD.Execute "usp_CD_Modificar_EstadoDelivery_Cabecera " + tEstadoDelivery + ", '" + CodigoPedido + "'"
    CnCD.Execute "UPDATE CD_MPEDIDO SET TMOTORIZADO ='" & sCodigoMotorizado & "' ,tnombremotorizado='" & Mid(sNombremotorizado, 1, 50) & "'  WHERE TCODIGOPEDIDO='" & CodigoPedido & "'"
 
    'Asignacion Fecha Entrega
    Dim lEntrega As Boolean
    lEntrega = Calcular("Select ISNULL(lHoraEntregaDelivery,0) As codigo From TPARAMETRO", Cn)
    
    If lEntrega = True Then
           'CnCD.Execute "UPDATE CD_MPEDIDO SET fEntrega = getDate()  WHERE TCODIGOPEDIDO='" & CodigoPedido & "'"
           CnCD.Execute "usp_CD_Modificar_EstadoDelivery_Cabecera '" & 8 & "', '" & CodigoPedido & "', ''"
    End If
    
    CnCD.Close
 
End Sub


'Private Sub ModificarEstadoPedidoWebAp(CodigoPedido As String)
'    Dim CnWAP As Connection
'
'    Set CnWAP = New Connection
'    CnWAP.Provider = "SQLOLEDB"
'    CnWAP.CursorLocation = adUseServer
'    CnWAP.ConnectionString = "User ID=" & sUserName & _
'                         ";password=" & sUserPassword & _
'                         ";Data Source=" & sRutaWAP & _
'                         ";Initial Catalog=" & sMDBWAP
'    CnWAP.CommandTimeout = 250
'    CnWAP.Open
'
'    CnWAP.Execute "APP_ACTUALIZAR_MPEDIDO_ESTADO '" & CodigoPedido & "','7'"
'    CnWAP.Close
'
'End Sub

Private Sub Timer1_Timer()
On Error GoTo fin
    If RsCabecera.RecordCount > 0 And RsCabecera.AbsolutePosition > 0 Then
        Dim s As Integer
        s = RsCabecera.AbsolutePosition
        RsCabecera.Requery
        RsCabecera.AbsolutePosition = s
    Else
        RsCabecera.Requery
    End If
    Exit Sub
fin:
    RsCabecera.Requery
End Sub
