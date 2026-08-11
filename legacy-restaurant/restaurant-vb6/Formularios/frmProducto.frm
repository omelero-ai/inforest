VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmProducto 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8100
   ClientLeft      =   2535
   ClientTop       =   1725
   ClientWidth     =   14115
   Icon            =   "frmProducto.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   14115
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   14055
      TabIndex        =   1
      Top             =   7395
      Width           =   14115
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Patio Comidas"
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
         Left            =   9840
         Picture         =   "frmProducto.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Enviar Menu a PatioComidas"
         Top             =   50
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Uber"
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
         Left            =   8760
         Picture         =   "frmProducto.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Enviar Menu a Uber Eats"
         Top             =   50
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   5325
         TabIndex        =   14
         Top             =   30
         Width           =   5385
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   540
            Picture         =   "frmProducto.frx":0F56
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   1080
            Picture         =   "frmProducto.frx":1498
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmProducto.frx":19DA
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4800
            Picture         =   "frmProducto.frx":1F1C
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4260
            Picture         =   "frmProducto.frx":245E
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3720
            Picture         =   "frmProducto.frx":29A0
            Style           =   1  'Graphical
            TabIndex        =   15
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
            TabIndex        =   21
            Top             =   150
            Width           =   2025
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
         Left            =   5350
         Picture         =   "frmProducto.frx":2EE2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Agrega"
         Top             =   50
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
         Left            =   5900
         Picture         =   "frmProducto.frx":3414
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Modifica"
         Top             =   50
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
         Left            =   6450
         Picture         =   "frmProducto.frx":3516
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Emite"
         Top             =   50
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
         Left            =   8100
         Picture         =   "frmProducto.frx":3A48
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir"
         Top             =   50
         Width           =   555
      End
      Begin VB.CommandButton cmdProcesa 
         Enabled         =   0   'False
         Height          =   555
         Left            =   14160
         Picture         =   "frmProducto.frx":3B3A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Filtrar"
         Top             =   120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame fraFechas 
         Caption         =   " Rango de Fechas "
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
         Height          =   645
         Left            =   11160
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   3660
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   345
            Left            =   2205
            TabIndex        =   5
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
            Format          =   58720257
            CurrentDate     =   37539
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   345
            Left            =   450
            TabIndex        =   6
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
            Format          =   58720257
            CurrentDate     =   37539
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
            TabIndex        =   8
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
            Left            =   1845
            TabIndex        =   7
            Top             =   270
            Width           =   225
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   555
         Index           =   5
         Left            =   7560
         Picture         =   "frmProducto.frx":3C3C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Reportes"
         Top             =   50
         Width           =   555
      End
      Begin VB.CommandButton cmdExporta 
         Height          =   555
         Left            =   7000
         Picture         =   "frmProducto.frx":416E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exportar HTML"
         Top             =   50
         Width           =   555
      End
   End
   Begin VB.Frame fraGrilla 
      Height          =   7440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   7140
         Left            =   90
         TabIndex        =   22
         Top             =   180
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   12594
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
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H400040&,.borderColor=&H808000&"
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
         Left            =   90
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera   As Recordset
Dim RsReporte       As Recordset
Dim rsSAP           As Recordset
Dim nColumna        As Integer
Dim Reporte         As New dsrPlato
Dim ReporteC        As New dsrPlatoCombo
Dim Reporte2        As New dsrPrecios

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
Private Sub EnviaMenuPatioComidas()
On Error GoTo fin
    Screen.MousePointer = vbHourglass
    Dim URL As String, CodTienda As String
    Dim xRazSoc As String, xEst As String, xCon As String, xDir As String
    Dim xRazSocX As Long, xEstX As Long, xConX As Long, xDirX As Long
    Dim xRazSocY As Long, xEstY As Long, xConY As Long, xDirY As Long
    Dim xWml As New XMLHTTP
 

    URL = Calcular("select isnull(tUrlApiWeb,'http://www.microsoft.com/api') as codigo from tparametro", Cn)
    CodTienda = Calcular("select isnull(tCodigoTiendaApiWeb,'1') as codigo from tparametro", Cn)
    xWml.Open "GET", URL & "/PatioComida/GetEnviarMenuPatioComidas?token=" & CodTienda, False
    xWml.send
    If xWml.Status = 200 Or xWml.Status = 204 Then
        MsgBox xWml.Status & " " & xWml.StatusText & ": Proceso correcto", vbInformation, sMensaje
    ElseIf xWml.Status = 500 Then
        MsgBox xWml.Status & xWml.StatusText, vbInformation, sMensaje
    ElseIf xWml.Status = 12007 Then
        MsgBox xWml.Status & xWml.StatusText, vbInformation, sMensaje
    Else
        MsgBox xWml.Status & xWml.StatusText, vbInformation, sMensaje
    End If
    
    Set xWml = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    MsgBox "Mensaje: " & error
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdOpcion_Click(Index As Integer)
frmProductoDetalle.SAPVALOR = False ' -- solo para opcion sap
Dim xGrupo As String
Dim xRGrupoAcceso As Recordset
xGrupo = Calcular("Select tGrupoUsuario As Codigo From TUSUARIO where tResumido = '" & sUsuario & "'", Cn)
    If xGrupo <> "00" Then
            Isql = "Select * From TGRUPOACCESO WHERE tGrupoUsuario = '" & xGrupo & "'"
            Set xRGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
                '10305020 MODIFICAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10305020"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(1).Enabled = False
                End If
                  '10305020 AGREGAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10305010"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(0).Enabled = False
                End If
    End If
    Select Case Index
           Case Is = 0  'Nuevo
                Sw = True
                'Cambiar el Nombre del Formulario Detalle
                frmProductoDetalle.Show vbModal
                Sw = True
                
           Case Is = 1  'Modifica
                If cmdOpcion(1).Enabled = False Then
                    MsgBox ("Sin Acceso")
                Else
                    If RsCabecera.RecordCount > 0 Then
                        Sw = False
                        'Cambiar el Nombre del Formulario Detalle
                        frmProductoDetalle.Show vbModal
                    Else
                        MsgBox "No Existe Datos Ingresados", vbExclamation, sMensaje
                    End If
                End If
           Case Is = 2 ' Envio de menu a Uber
                If (EjecutaMotorIntegraciones("MENU", sUsuario, 3, 1)) Then
                    MsgBox "El proceso se realizo correctamente, Favor de verificar!!!", vbInformation
                Else
                    MsgBox "El proceso no se realizo con exito!!!" + vbNewLine + "- Revisar la salida a internet" + vbNewLine + "- Contactarse con Infhotel Servicios informaticos SAC!!!", vbCritical
                End If
                
           Case Is = 3 ' Envio de menu a patio comidafv
            If MsgBox("Esta seguro de enviar el menu a patio de comidas?", vbYesNo) = vbYes Then
                 Call EnviaMenuPatioComidas
            End If
            
           Case Is = 5  'Emite
                frmFiltroProducto.Show vbModal
                If Not wEnter Then
                   Exit Sub
                End If
                
                If sDescrip = "Combo" Then
                    Isql = "usp_InforestAdm_ObtenerProductosReporteLocal2 '" & sTipoCanalReporte & "' ,1,'" & sCodigo & "'"
                    Set RsReporte = Lib.OpenRecordset(Isql, Cn)
'                    If sCodigo <> "" Then
'                        sCodigo = sCodigo & " and Combo <> ''"
'                    Else
'                        sCodigo = "Combo <> ''"
'                    End If
'                    If RsReporte.RecordCount > 0 Then
'                       RsReporte.Filter = sCodigo
'                    End If
                Else
                    If sDescrip = "Precios" Then
                    
                       If LTrim(sCodigo) <> "" Then
                          sCodigo = "where " & sCodigo
                        Else ' GCAA 30122020
                            sCodigo = ""
                       End If
                       
                       Select Case pais ' ok
                        Case "001" 'Bolivia
                                    Isql = "select TipoProducto, SubGrupo, Descripcion, Moneda, " & _
                                           "case when lLocal=0 then 0 else nprecioventa - ( nPrecioVenta * nFactorLocal ) end as nNetoSalon, case when lLocal=0 then 0 else nPrecioVenta * nImpuesto1 / 100 end as nImpuesto1, case when lLocal=0 then 0 else nPrecioVenta * nImpuesto2 / 100 end as nImpuesto2, case when lLocal=0 then 0 else nPrecioVenta * nImpuesto3 / 100 end as nImpuesto3, case when lLocal=0 then 0 else nPrecioVenta end as nPrecioVenta, " & _
                                           "case when lDelivery=0 then 0 else nPrecioDelivery - (nPrecioDelivery * nFactorDelivery) end as nNetoDelivery, case when lDelivery=0 then 0 else nPrecioDelivery * nImpuesto4 / 100 end as nImpuesto4, case when lDelivery=0 then 0 else nPrecioDelivery * nImpuesto5 / 100 end as nImpuesto5, case when lDelivery=0 then 0 else nPrecioDelivery  * nImpuesto6 / 100 end as nImpuesto6, case when lDelivery=0 then 0 else nPrecioDelivery end as nPrecioDelivery, " & _
                                           "case when lLlevar=0 then 0 else nPreciollevar - (nPreciollevar * nFactorllevar) end as nNetoLlevar, case when lLlevar=0 then 0 else nPreciollevar * nImpuesto7 / 100 end as nImpuesto7, case when lLlevar=0 then 0 else nPreciollevar * nImpuesto8 / 100 end as nImpuesto8, case when lLlevar=0 then 0 else nPreciollevar * nImpuesto9 / 100 end as nImpuesto9, case when lLlevar=0 then 0 else nPreciollevar end as nPreciollevar, " & _
                                           "case when lCanal4=0 then 0 else nPrecioCanal4 - (nPrecioCanal4 * nFactorCanal4) end as nNetoCanal4, case when lCanal4=0 then 0 else nPrecioCanal4 * nImpuesto10 / 100 end as nImpuesto10, case when lCanal4=0 then 0 else nPrecioCanal4 * nImpuesto11 / 100 end as nImpuesto11, case when lCanal4=0 then 0 else nPrecioCanal4 * nImpuesto12 / 100 end as nImpuesto12, case when lCanal4=0 then 0 else nPrecioCanal4 end as nPrecioCanal4, " & _
                                           "case when lCanal5=0 then 0 else nPrecioCanal5 - (nPrecioCanal5 * nFactorCanal5) end as nNetoCanal5, case when lCanal5=0 then 0 else nPrecioCanal5 * nImpuesto12 / 100 end as nImpuesto13, case when lCanal5=0 then 0 else nPrecioCanal5 * nImpuesto14 / 100 end as nImpuesto14, case when lCanal5=0 then 0 else nPrecioCanal5  * nImpuesto15 / 100 end as nImpuesto15, case when lCanal5=0 then 0 else nPrecioCanal5 end as nPrecioCanal5 " & _
                                           "From (select tCodigoProducto, " & _
                                           "case when lImpuesto1=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto1, " & "case when lImpuesto2=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto2, " & _
                                           "case when lImpuesto3=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto3, " & "case when lImpuesto4=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto4, " & _
                                           "case when lImpuesto5=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto5, " & "case when lImpuesto6=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto6, " & _
                                           "case when lImpuesto7=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto7, " & "case when lImpuesto8=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto8, " & _
                                           "case when lImpuesto9=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto9, " & "case when lImpuesto10=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto10, " & _
                                           "case when lImpuesto11=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto11, " & "case when lImpuesto12=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto12, " & _
                                           "case when lImpuesto13=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto13, " & "case when lImpuesto14=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto14, " & _
                                           "case when lImpuesto15=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto15, " & _
                                           " ((case when lImpuesto1=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto2 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto3 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorLocal, " & _
                                           " ((case when lImpuesto4=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto5 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto6 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorDelivery," & _
                                           " ((case when lImpuesto7=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto8 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto9 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorLlevar, " & _
                                           " ((case when lImpuesto10=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto11 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto12 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorCanal4," & _
                                           " ((case when lImpuesto13=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto14 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto15 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorCanal5 " & _
                                           "from TPRODUCTO) IMPUESTOS INNER JOIN dbo.vProducto ON dbo.vProducto.Codigo = IMPUESTOS.tCodigoProducto " & sCodigo
                                           
                              Case Else 'Peru, Ecuador
                                    Isql = "select TipoProducto, SubGrupo, Descripcion, Moneda, " & _
                                           "case when lLocal=0 then 0 else nPrecioVenta / nFactorLocal end as nNetoSalon, case when lLocal=0 then 0 else nPrecioVenta / nFactorLocal * nImpuesto1 / 100 end as nImpuesto1, case when lLocal=0 then 0 else nPrecioVenta / nFactorLocal * nImpuesto2 / 100 end as nImpuesto2, case when lLocal=0 then 0 else nPrecioVenta / nFactorLocal * nImpuesto3 / 100 end as nImpuesto3, case when lLocal=0 then 0 else nPrecioVenta end as nPrecioVenta, " & _
                                           "case when lDelivery=0 then 0 else nPrecioDelivery / nFactorDelivery end as nNetoDelivery, case when lDelivery=0 then 0 else nPrecioDelivery / nFactorDelivery * nImpuesto4 / 100 end as nImpuesto4, case when lDelivery=0 then 0 else nPrecioDelivery / nFactorDelivery * nImpuesto5 / 100 end as nImpuesto5, case when lDelivery=0 then 0 else nPrecioDelivery / nFactorDelivery * nImpuesto6 / 100 end as nImpuesto6, case when lDelivery=0 then 0 else nPrecioDelivery end as nPrecioDelivery, " & _
                                           "case when lLlevar=0 then 0 else nPreciollevar / nFactorllevar end as nNetoLlevar, case when lLlevar=0 then 0 else nPreciollevar / nFactorllevar * nImpuesto7 / 100 end as nImpuesto7, case when lLlevar=0 then 0 else nPreciollevar / nFactorllevar * nImpuesto8 / 100 end as nImpuesto8, case when lLlevar=0 then 0 else nPreciollevar / nFactorllevar * nImpuesto9 / 100 end as nImpuesto9, case when lLlevar=0 then 0 else nPreciollevar end as nPreciollevar, " & _
                                           "case when lCanal4=0 then 0 else nPrecioCanal4 / nFactorCanal4 end as nNetoCanal4, case when lCanal4=0 then 0 else nPrecioCanal4 / nFactorCanal4 * nImpuesto10 / 100 end as nImpuesto10, case when lCanal4=0 then 0 else nPrecioCanal4 / nFactorCanal4 * nImpuesto11 / 100 end as nImpuesto11, case when lCanal4=0 then 0 else nPrecioCanal4 / nFactorCanal4 * nImpuesto12 / 100 end as nImpuesto12, case when lCanal4=0 then 0 else nPrecioCanal4 end as nPrecioCanal4, " & _
                                           "case when lCanal5=0 then 0 else nPrecioCanal5 / nFactorCanal5 end as nNetoCanal5, case when lCanal5=0 then 0 else nPrecioCanal5 / nFactorCanal5 * nImpuesto12 / 100 end as nImpuesto13, case when lCanal5=0 then 0 else nPrecioCanal5 / nFactorCanal5 * nImpuesto14 / 100 end as nImpuesto14, case when lCanal5=0 then 0 else nPrecioCanal5 / nFactorCanal5 * nImpuesto15 / 100 end as nImpuesto15, case when lCanal5=0 then 0 else nPrecioCanal5 end as nPrecioCanal5 " & _
                                           "From (select tCodigoProducto, " & _
                                           "case when lImpuesto1=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto1, " & "case when lImpuesto2=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto2, " & _
                                           "case when lImpuesto3=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto3, " & "case when lImpuesto4=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto4, " & _
                                           "case when lImpuesto5=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto5, " & "case when lImpuesto6=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto6, " & _
                                           "case when lImpuesto7=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto7, " & "case when lImpuesto8=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto8, " & _
                                           "case when lImpuesto9=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto9, " & "case when lImpuesto10=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto10, " & _
                                           "case when lImpuesto11=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto11, " & "case when lImpuesto12=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto12, " & _
                                           "case when lImpuesto13=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end as nImpuesto13, " & "case when lImpuesto14=0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end as nImpuesto14, " & _
                                           "case when lImpuesto15=0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end as nImpuesto15, " & _
                                           "1 + ((case when lImpuesto1=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto2 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto3 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorLocal, " & _
                                           "1 + ((case when lImpuesto4=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto5 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto6 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorDelivery," & _
                                           "1 + ((case when lImpuesto7=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto8 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto9 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorLlevar, " & _
                                           "1 + ((case when lImpuesto10=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto11 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto12 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorCanal4," & _
                                           "1 + ((case when lImpuesto13=0 then 0 else " & Format(nPorcentaje1, "####.0000") & " end + case when lImpuesto14 = 0 then 0 else " & Format(nPorcentaje2, "####.0000") & " end + case when lImpuesto15 = 0 then 0 else " & Format(nPorcentaje3, "####.0000") & " end) / 100 ) as nFactorCanal5 " & _
                                           "from TPRODUCTO) IMPUESTOS INNER JOIN dbo.vProducto ON dbo.vProducto.Codigo = IMPUESTOS.tCodigoProducto " & sCodigo
                       End Select
                       
                       Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                    Else
                    'GCAA - 2 / 23072020
                        Isql = "usp_InforestAdm_ObtenerProductosReporteLocal2 '" & sTipoCanalReporte & "', 0 , '" & sCodigo & "' " ' & iif(
                        Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                        If sCodigo <> "" Then
                            sCodigo = sCodigo & " and Combo = ''"
                        Else
                            sCodigo = "Combo = ''"
                        End If
                        '-- gcaa-2 // 23072020
                        'If RsReporte.RecordCount > 0 Then
                        '   RsReporte.Filter = sCodigo
                        'End If
                    End If
                End If
                If sDescrip = "Precios" Then
                    Reporte2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Reporte2.DiscardSavedData
                    Reporte2.Database.SetDataSource RsReporte
                    Reporte2.Text15.SetText sRazonSocial
                    Reporte2.Text15.SetText localConectado
                    
                    If sImpuesto1 <> "" Then
                       Reporte2.Text2.SetText sImpuesto1
                       Reporte2.Text19.SetText sImpuesto1
                       Reporte2.Text20.SetText sImpuesto1
                       Reporte2.Text21.SetText sImpuesto1
                       Reporte2.Text22.SetText sImpuesto1
                    Else
                       Reporte2.Text2.SetText "n/d"
                       Reporte2.Text19.SetText "n/d"
                       Reporte2.Text20.SetText "n/d"
                       Reporte2.Text21.SetText "n/d"
                       Reporte2.Text22.SetText "n/d"
                    End If
                    
                    If sImpuesto2 <> "" Then
                       Reporte2.Text3.SetText sImpuesto2
                       Reporte2.Text23.SetText sImpuesto2
                       Reporte2.Text25.SetText sImpuesto2
                       Reporte2.Text26.SetText sImpuesto2
                       Reporte2.Text27.SetText sImpuesto2
                    Else
                       Reporte2.Text3.SetText "n/d"
                       Reporte2.Text23.SetText "n/d"
                       Reporte2.Text25.SetText "n/d"
                       Reporte2.Text26.SetText "n/d"
                       Reporte2.Text27.SetText "n/d"
                    End If
                    
                    If sImpuesto3 <> "" Then
                       Reporte2.Text5.SetText sImpuesto3
                       Reporte2.Text28.SetText sImpuesto3
                       Reporte2.Text29.SetText sImpuesto3
                       Reporte2.Text30.SetText sImpuesto3
                       Reporte2.Text31.SetText sImpuesto3
                    Else
                       Reporte2.Text5.SetText "n/d"
                       Reporte2.Text28.SetText "n/d"
                       Reporte2.Text29.SetText "n/d"
                       Reporte2.Text30.SetText "n/d"
                       Reporte2.Text31.SetText "n/d"
                    End If
                    If sBoton1 <> "" Then Reporte2.Text32.SetText sBoton1 Else Reporte2.Text32.SetText "n/d"
                    If sBoton2 <> "" Then Reporte2.Text33.SetText sBoton2 Else Reporte2.Text33.SetText "n/d"
                    If sBoton3 <> "" Then Reporte2.Text34.SetText sBoton3 Else Reporte2.Text34.SetText "n/d"
                    If sBoton4 <> "" Then Reporte2.Text36.SetText sBoton4 Else Reporte2.Text36.SetText "n/d"
                    If sBoton5 <> "" Then Reporte2.Text37.SetText sBoton5 Else Reporte2.Text37.SetText "n/d"
                    
                    Reporte2.PaperOrientation = crLandscape
                Else
                    If sDescrip = "Combo" Then
                        ReporteC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ReporteC.DiscardSavedData
                        ReporteC.Database.SetDataSource RsReporte
                        ReporteC.Text47.SetText "Canal de Venta: " & sTipoCanalNombreReporte
                        ReporteC.Text15.SetText sRazonSocial
                        ReporteC.Text35.SetText localConectado
                        ReporteC.PaperOrientation = crLandscape
                    Else
                        Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Reporte.DiscardSavedData
                        Reporte.Database.SetDataSource RsReporte
                        Reporte.Text47.SetText "Canal de Venta: " & sTipoCanalNombreReporte
                        Reporte.Text15.SetText sRazonSocial
                        Reporte.Text35.SetText localConectado
                        Reporte.PaperOrientation = crLandscape
                    End If
                End If
                
                Select Case sTipo
                       Case "Pantalla"
                            If sDescrip = "Precios" Then
                                frmEmite.CRViewer.ReportSource = Reporte2
                                frmEmite.CRViewer.DisplayGroupTree = False
                                frmEmite.CRViewer.ViewReport
                                frmEmite.Show vbModal
                            Else
                                If sDescrip = "Combo" Then
                                    frmEmite.CRViewer.ReportSource = ReporteC
                                Else
                                    frmEmite.CRViewer.ReportSource = Reporte
                                End If
                                frmEmite.CRViewer.DisplayGroupTree = False
                                frmEmite.CRViewer.ViewReport
                                frmEmite.Show vbModal
                            End If
                       Case "Impresora"
                            Screen.MousePointer = vbDefault
                            If sDescrip = "Precios" Then
                               Reporte2.PrintOut
                            Else
                               If sDescrip = "Combo" Then
                                    ReporteC.PrintOut
                               Else
                                    Reporte.PrintOut
                               End If
                            End If
                       Case "Excel"
                            Screen.MousePointer = vbDefault
                            If sDescrip = "Precios" Then
                                Reporte2.ExportOptions.FormatType = 21
                                Reporte2.ExportOptions.DestinationType = 1
                                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                                cmdSave.ShowSave
                                If cmdSave.FileName = "" Then
                                   Exit Sub
                                End If
                                Reporte2.ExportOptions.DiskFileName = cmdSave.FileName
                                Reporte2.Export False
                            Else
                                If sDescrip = "Combo" Then
                                    ReporteC.ExportOptions.FormatType = 21
                                    ReporteC.ExportOptions.DestinationType = 1
                                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                                    cmdSave.ShowSave
                                    If cmdSave.FileName = "" Then
                                       Exit Sub
                                    End If
                                    ReporteC.ExportOptions.DiskFileName = cmdSave.FileName
                                    ReporteC.Export False
                                Else
                                    Reporte.ExportOptions.FormatType = 21
                                    Reporte.ExportOptions.DestinationType = 1
                                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                                    cmdSave.ShowSave
                                    If cmdSave.FileName = "" Then
                                       Exit Sub
                                    End If
                                    Reporte.ExportOptions.DiskFileName = cmdSave.FileName
                                    Reporte.Export False
                                End If
                            End If
                End Select
                      
           Case Is = 6  'Salir
                Unload Me
           
    End Select
End Sub

Private Sub Form_Load()
Dim xGrupo As String
Dim xRGrupoAcceso As Recordset
xGrupo = Calcular("Select tGrupoUsuario As Codigo From TUSUARIO where tResumido = '" & sUsuario & "'", Cn)
    If xGrupo <> "00" Then
            Isql = "Select * From TGRUPOACCESO WHERE tGrupoUsuario = '" & xGrupo & "'"
            Set xRGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
            '10175010 AGREGAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10305010"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(0).Enabled = False
                End If
            
            '10175020 MODIFICAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=10305020"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(1).Enabled = False
                End If
    End If

    Screen.MousePointer = vbHourglass
    
    ' Ingrese el Titulo
    Me.Caption = " Mantenimiento de Productos y Precios "
    grdGrilla.Caption = Me.Caption
    Centrar Me
    nColumna = 0
    
    Dim xImp1 As Double
    Dim xImp2 As Double
    Dim xImp3 As Double
    
    xImp1 = (nPorcentaje1 / 100)
    xImp2 = (nPorcentaje2 / 100)
    xImp3 = (nPorcentaje3 / 100)
    
    'Ingrese la Vista
'     Isql = "select *, ' ' AS Combo, (case nPrecioVenta when 0 then 0 else nPrecioVenta / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when 1 then " & xImp3 & " else 0 end)) end) as nPrecioBase,  str(nPrecioVenta,10,2) as xPrecioVenta,  convert(varchar, fRegistro, 103)  as xRegistro,str(isnull(nInsumo,0)+isnull(nGasto,0)+isnull(nManoObra,0),10,2) as xCosto, str((case nPrecioVenta when 0 then 0 else nPrecioVenta / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + " & _
'            "case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end),10,2) as xPrecioBase, str((isnull(nInsumo,0)+isnull(nGasto,0)+isnull(nManoObra,0)) / (case nPrecioVenta when 0 then 0 else nPrecioVenta / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end)*100,10,2) as xFoodCost From vProducto "

     Isql = "select *, ' ' AS Combo, (case Round(nPrecioVenta,2) when 0 then 0 else nPrecioVenta / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when 1 then " & xImp3 & " else 0 end)) end) as nPrecioBase,  str(nPrecioVenta,10,2) as xPrecioVenta,  convert(varchar, fRegistro, 103)  as xRegistro,str(isnull(nInsumo,0)+isnull(nGasto,0)+isnull(nManoObra,0),10,2) as xCosto, str((case Round(nPrecioVenta,2) when 0 then 0 else nPrecioVenta / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + " & _
            "case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end),10,2) as xPrecioBase, str((isnull(nInsumo,0)+isnull(nGasto,0)+isnull(nManoObra,0)) / (case Round(nPrecioVenta,2) when 0 then 1 else nPrecioVenta / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end)*100,10,2) as xFoodCost, + " & _
            "(case Round(nPrecioDelivery,2) when 0 then 0 else nPrecioDelivery / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when 1 then " & xImp3 & " else 0 end)) end) as nPrecioBase2,  str(nPrecioDelivery,10,2) as xPrecioVenta2,str(isnull(nInsumo2,0)+isnull(nGasto2,0)+isnull(nManoObra2,0),10,2) as xCosto2, str((case Round(nPrecioDelivery,2) when 0 then 0 else nPrecioDelivery / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + " & _
            "case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end),10,2) as xPrecioBase2, str((isnull(nInsumo2,0)+isnull(nGasto2,0)+isnull(nManoObra2,0)) / (case Round(nPrecioDelivery,2) when 0 then 1 else nPrecioDelivery / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end)*100,10,2) as xFoodCost2, + " & _
            "(case Round(nPrecioLlevar,2) when 0 then 0 else nPrecioLlevar / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when 1 then " & xImp3 & " else 0 end)) end) as nPrecioBase3,  str(nPrecioLlevar,10,2) as xPrecioVenta3,str(isnull(nInsumo3,0)+isnull(nGasto3,0)+isnull(nManoObra3,0),10,2) as xCosto3, str((case Round(nPrecioLlevar,2) when 0 then 0 else nPrecioLlevar / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + " & _
            "case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end),10,2) as xPrecioBase3, str((isnull(nInsumo3,0)+isnull(nGasto3,0)+isnull(nManoObra3,0)) / (case Round(nPrecioLlevar,2) when 0 then 1 else nPrecioLlevar / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end)*100,10,2) as xFoodCost3, + " & _
            "(case Round(nPrecioCanal4,2) when 0 then 0 else nPrecioCanal4 / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when 1 then " & xImp3 & " else 0 end)) end) as nPrecioBase4,  str(nPrecioCanal4,10,2) as xPrecioVenta4,str(isnull(nInsumo4,0)+isnull(nGasto4,0)+isnull(nManoObra4,0),10,2) as xCosto4, str((case Round(nPrecioCanal4,2) when 0 then 0 else nPrecioCanal4 / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + " & _
            "case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end),10,2) as xPrecioBase4, str((isnull(nInsumo4,0)+isnull(nGasto4,0)+isnull(nManoObra4,0)) / (case Round(nPrecioCanal4,2) when 0 then 1 else nPrecioCanal4 / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end)*100,10,2) as xFoodCost4, + " & _
            "(case Round(nPrecioCanal5,2) when 0 then 0 else nPrecioCanal5 / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when 1 then " & xImp3 & " else 0 end)) end) as nPrecioBase5,  str(nPrecioCanal5,10,2) as xPrecioVenta5,str(isnull(nInsumo5,0)+isnull(nGasto5,0)+isnull(nManoObra5,0),10,2) as xCosto5, str((case Round(nPrecioCanal5,2) when 0 then 0 else nPrecioCanal5 / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + " & _
            "case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end),10,2) as xPrecioBase5, str((isnull(nInsumo5,0)+isnull(nGasto5,0)+isnull(nManoObra5,0)) / (case Round(nPrecioCanal5,2) when 0 then 1 else nPrecioVenta / ( 1.00 + (case lImpuesto1 when 1 then " & xImp1 & " else 0 end + case lImpuesto2 when 1 then " & xImp2 & " else 0 end + case lImpuesto3 when " & xImp3 & " then 0 else 0 end)) end)*100,10,2) as xFoodCost5 From vProducto "

    ' order by Grupo, SubGrupo, Descripcion
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
 
    
    'Configuración de la Grilla
    Call ConfGrilla(27, grdGrilla, "Codigo", 2, "Codigo", 1000, 2, 0, "", _
                                  "Grupo", 2, "Grupo", 2400, 0, 0, "", _
                                  "Sub Grupo", 2, "SubGrupo", 2500, 0, 0, "", _
                                  "Tipo Producto", 2, "TipoProducto", 2400, 0, 0, "", _
                                  "Descripción", 2, "Descripcion", 3700, 0, 0, "", _
                                  "CAN1 Precio", 2, "xPrecioVenta", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN1 Neto", 2, "xPrecioBase", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN1 Costo", 2, "xCosto", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN1 FC", 2, "xFoodCost", 1400, 1, 0, "###,###,##0.00", _
                                  "Fecha", 2, "xRegistro", 1400, 0, 0, "dd/MM/yyyy", _
                                  "Activo", 2, "lactivo", 1100, 2, 4, "", _
                                  "CAN2 Precio", 2, "xPrecioVenta2", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN2 Neto", 2, "xPrecioBase2", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN2 Costo", 2, "xCosto2", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN2 FC", 2, "xFoodCost2", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN3 Precio", 2, "xPrecioVenta3", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN3 Neto", 2, "xPrecioBase3", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN3 Costo", 2, "xCosto3", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN3 FC", 2, "xFoodCost3", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN4 Precio", 2, "xPrecioVenta4", 1400, 1, 0, "###,###,##0.00", "CAN4 Neto", 2, "xPrecioBase4", 1400, 1, 0, "###,###,##0.00", "CAN4 Costo", 2, "xCosto4", 1400, 1, 0, "###,###,##0.00", "CAN4 FC", 2, "xFoodCost4", 1400, 1, 0, "###,###,##0.00", _
                                  "CAN5 Precio", 2, "xPrecioVenta5", 1400, 1, 0, "###,###,##0.00", "CAN5 Neto", 2, "xPrecioBase5", 1400, 1, 0, "###,###,##0.00", "CAN5 Costo", 2, "xCosto5", 1400, 1, 0, "###,###,##0.00", "CAN5 FC", 2, "xFoodCost5", 1400, 1, 0, "###,###,##0.00")

                                  
    Set grdGrilla.DataSource = RsCabecera
'    If lIntPatio Then
'        cmdOpcion(3).Visible = True
'    End If
    '--- Activación del timer sap
'    If lSAP Then
'        TimerSAP.Enabled = True
'        cmdOpcion(0).Enabled = False
'    Else
'        TimerSAP.Enabled = False
'        cmdOpcion(0).Enabled = True
'     End If
    '--------------------------
    cmdTexto.Caption = "Registro " & RsCabecera.AbsolutePosition & " de " & RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmImpresora = Nothing
End Sub

Private Sub grdGrilla_DblClick()
    cmdOpcion_Click (1)
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   cmdTexto.Caption = "Registro " & IIf(RsCabecera.RecordCount = 0, 0, RsCabecera.AbsolutePosition) & " de " & RsCabecera.RecordCount
End Sub

'Private Sub TimerSAP_Timer()
'       If lSAP Then
'            'Isql = "select * from [" & sServidorSAp & "]." & sBdSAP & ".[dbo].[X_INF_ProductoVenta] "
'            Isql = "exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
'
'            Set rsSAP = Lib.OpenRecordset(Isql, Cn)
'            If rsSAP.RecordCount > 0 Then
'                Me.cmdSAP.Caption = "(" + CStr(rsSAP.RecordCount) + " Prod. Nuevos)"
'                Me.cmdSAP.Visible = True
'            Else
'                Me.cmdSAP.Visible = False
'            End If
'        End If
'
'        If lSAP And Me.cmdSAP.Visible = True Then
'            If Me.cmdSAP.backColor = &HFF00& Then
'               Me.cmdSAP.backColor = &H8000000F
'
'            Else
'               Me.cmdSAP.backColor = &HFF00&
'            End If
'
'        End If
'End Sub

'Private Sub cmdSAP_Click()
'  On Error GoTo fin
'
'    'Isql = "select Inf_CodigoExterno as Codigo ,  Inf_DescripcionDetallada as Descripcion from [" & sServidorSAp & "]." & sBdSAP & ".[dbo].[X_INF_ProductoVenta] where Inf_CodigoLocal = '" & sCodSap & "'"
'     Isql = "exec sp_VinculacionSAP '" & sServidorSAp & "','" & sBdSAP & "','" & sCodSap & "','','',2"
'
'
'    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1000, 2, 0, "", _
'                                                    "Descripción", 2, "Descripcion", 3800, 0, 0, "")
'       frmBusquedaRapida.Show vbModal
'       Sw = True
'       If VSApForBusqueda = False Then
'        frmProductoDetalle.SAPVALOR = True
'        frmProductoDetalle.Show vbModal
'       End If
'       VSApForBusqueda = False
'    Exit Sub
'
'fin:
'MsgBox "error verificar"
'End Sub
