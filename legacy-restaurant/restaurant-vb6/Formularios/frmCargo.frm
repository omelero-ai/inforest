VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmCargo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "frmCargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtObservacion 
      Height          =   795
      Left            =   2205
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   5670
      Width           =   6855
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Facturable"
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
      Index           =   18
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6465
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Propina US$"
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
      Index           =   17
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6465
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Propina S/."
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
      Index           =   16
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6465
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Filtrar Nombre Rva"
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
      Left            =   9150
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5130
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Reserva Cerrada"
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
      Index           =   5
      Left            =   45
      TabIndex        =   3
      Top             =   4980
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
      Index           =   2
      Left            =   7785
      Picture         =   "frmCargo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6465
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Contado InfHotel"
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
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5625
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Filtrar Habitación"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4020
      Width           =   1275
   End
   Begin VB.TextBox txtResultado 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   10395
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "No Filtro"
      Height          =   555
      Index           =   9
      Left            =   9150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5775
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Filtrar Pasajero"
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
      Index           =   7
      Left            =   9150
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4570
      Width           =   1275
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
      Left            =   6390
      Picture         =   "frmCargo.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6465
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   15
      Left            =   9150
      Picture         =   "frmCargo.frx":050E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3270
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   14
      Left            =   9150
      Picture         =   "frmCargo.frx":0DD8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2715
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   13
      Left            =   9150
      Picture         =   "frmCargo.frx":16A2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   12
      Left            =   9150
      Picture         =   "frmCargo.frx":1F6C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1605
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   11
      Left            =   9150
      Picture         =   "frmCargo.frx":2836
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1044
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Height          =   555
      Index           =   10
      Left            =   9150
      Picture         =   "frmCargo.frx":3100
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   495
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Habitación"
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
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3690
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Reserva"
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
      Index           =   4
      Left            =   45
      TabIndex        =   2
      Top             =   4335
      Width           =   1275
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   5130
      Left            =   2205
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   495
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   9049
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
      ScrollTrack     =   -1  'True
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
   Begin VB.Frame fraHabitacion 
      Caption         =   " Habitación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   0
      TabIndex        =   21
      Top             =   495
      Width           =   2130
      Begin VB.TextBox txtCia 
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   22
         Top             =   1545
         Width           =   1950
      End
      Begin VB.TextBox txtReserva 
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   31
         Top             =   2475
         Width           =   1950
      End
      Begin VB.TextBox txtNacionalidad 
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   23
         Top             =   585
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reserva :"
         Height          =   195
         Left            =   90
         TabIndex        =   32
         Top             =   2190
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nacionalidad :"
         Height          =   195
         Left            =   90
         TabIndex        =   25
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Compañía"
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   1260
         Width           =   1950
      End
   End
   Begin VB.Frame fraReserva 
      Caption         =   " Reserva "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   0
      TabIndex        =   26
      Top             =   495
      Width           =   2130
      Begin VB.TextBox txtFOut 
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   28
         Top             =   1545
         Width           =   1950
      End
      Begin VB.TextBox txtFIn 
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   27
         Top             =   585
         Width           =   1950
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrada :"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Salida :"
         Height          =   195
         Left            =   90
         TabIndex        =   30
         Top             =   1260
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrilla As Recordset
Dim RsReserva As Recordset
Dim sCriterio As String
Dim sTipoCargo As String
Dim sFiltro As String


Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Enviar Infhotel
               If cmdOpcion(18).Caption = "Facturable" Then
                  sTipoComanda = "01"
               Else
                  sTipoComanda = "02"
               End If
               sDescrip = "Infhotel"
               sReserva = ""
               sHabitacion = ""
               sPasajero = ""
               
               wEnter = True
               Unload Me
               
          Case Is = 1 ' Cancelar
               sDescrip = ""
               sPropina = ""
               sMonPropina = ""
               sTipoComanda = ""
               wEnter = False
               Unload Me
                    
          Case Is = 2 ' Aceptar
               wEnter = True
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               
               sDescrip = sTipoCargo
               If sTipoCargo = "Habitacion" Then
                  sReserva = IIf(RsGrilla.EOF = True, "", RsGrilla!Reserva)
                  sHabitacion = IIf(RsGrilla.EOF = True, "", RsGrilla!Habitacion)
                  sPasajero = Mid(Apostrofe(IIf(RsGrilla.EOF = True, "", RsGrilla!Pasajero)), 1, 45)
                  sFichaPasajero = IIf(RsGrilla.EOF = True, "", RsGrilla!tCodigoPasajero)
               Else
                  sReserva = IIf(RsGrilla.EOF = True, "", RsGrilla!Reserva)
                  sHabitacion = ""
                  sPasajero = ""
                  sFichaPasajero = ""
               End If
               Unload Me
                              
          Case Is = 3 ' Habitacion
               Call confGrillaHabitacion
            
          Case Is = 4 ' Reserva
               sFiltro = "AND vReserva.fSalida>='" & Format(FechaServidor(), "yyyy/MM/dd") & "'"
               Call confGrillaReserva
                                           
          Case Is = 5 ' Cerrada
               sFiltro = "AND vReserva.fSalida<='" & Format(FechaServidor(), "yyyy/MM/dd") & "'"
               Call confGrillaReserva
          
          Case Is = 6 'Filtrar Habitación
               frmKeyBoard.Show vbModal
               If wEnter And Len(Trim(sDescrip)) > 0 Then
                  If sTipoCargo = "Habitacion" Then
                     sCriterio = "Habitacion like *" & Trim(sDescrip) & "*"
                  Else
                     sCriterio = "Reserva like *" & Trim(sDescrip) & "*"
                  End If
                  Screen.MousePointer = vbHourglass
                  RsGrilla.Filter = sCriterio
                  If RsGrilla.RecordCount <= 0 Then
                     RsGrilla.Filter = adFilterNone
                     MsgBox "Criterio No Encontrado", vbExclamation, sMensaje
                  Else
                    RsGrilla.MoveFirst
                    If sTipoCargo = "Habitacion" Then
                       txtResultado.Text = RsGrilla!Pasajero
                    Else
                       txtResultado.Text = RsGrilla!Grupo
                    End If
                  End If
                End If
                wEnter = False
                Screen.MousePointer = vbDefault
          
          Case Is = 7 'Filtrar Pasajero
               frmKeyBoard.Show vbModal
               If wEnter And Len(Trim(sDescrip)) > 0 Then
                  If sTipoCargo = "Habitacion" Then
                     sCriterio = "Pasajero like *" & Trim(sDescrip) & "*"
                  Else
                     sCriterio = "Compania like *" & Trim(sDescrip) & "*"
                  End If
                  
                  Screen.MousePointer = vbHourglass
                  RsGrilla.Filter = sCriterio
                  If RsGrilla.RecordCount <= 0 Then
                     RsGrilla.Filter = adFilterNone
                     MsgBox "Criterio No Encontrado", vbExclamation, sMensaje
                  Else
                    RsGrilla.MoveFirst
                    If sTipoCargo = "Habitacion" Then
                       txtResultado.Text = RsGrilla!Pasajero
                    Else
                       txtResultado.Text = RsGrilla!Grupo
                    End If

                  End If
              End If
                wEnter = False
                Screen.MousePointer = vbDefault
          
          Case Is = 8 'Filtrar Grupo
               frmKeyBoard.Show vbModal
               If wEnter And Len(Trim(sDescrip)) > 0 Then
                  sCriterio = "Grupo like *" & Trim(sDescrip) & "*"
                  Screen.MousePointer = vbHourglass
                  RsGrilla.Filter = sCriterio
                  If RsGrilla.RecordCount <= 0 Then
                     RsGrilla.Filter = adFilterNone
                     MsgBox "Criterio No Encontrado", vbExclamation, sMensaje
                  Else
                    RsGrilla.MoveFirst
                    txtResultado.Text = RsGrilla!Grupo
                  End If
                End If
                wEnter = False
                Screen.MousePointer = vbDefault
                    
          Case Is = 9 ' Todos
               Screen.MousePointer = vbHourglass
               RsGrilla.Filter = adFilterNone
               txtResultado.Text = ""
               Set grdGrilla.DataSource = RsGrilla
               Screen.MousePointer = vbDefault
               
               If sTipoCargo = "Habitacion" Then
                  txtResultado.Text = RsGrilla!Pasajero
               Else
                  txtResultado.Text = RsGrilla!Grupo
               End If
          
          Case Is = 10  ' Primero
               MoverPuntero Primero, grdGrilla
               
          Case Is = 11  ' PG Up
               MoverPuntero pgup, grdGrilla
               
          Case Is = 12  ' Previo
               MoverPuntero previo, grdGrilla
               
          Case Is = 13  ' Siguiente
               MoverPuntero siguiente, grdGrilla
               
          Case Is = 14  ' PG Dn
               MoverPuntero pgdn, grdGrilla
               
          Case Is = 15 ' Ultimo
               MoverPuntero Ultimo, grdGrilla
        
          Case Is = 16 'Propina S/.
            sTipo = ""
            sDescrip = ""
            frmNumPad.Show vbModal
            sPropina = sDescrip
            sMonPropina = "01"
            wEnter = False
            
          Case Is = 17 'Propina U$
            sTipo = ""
            sDescrip = ""
            frmNumPad.Show vbModal
            sPropina = sDescrip
            sMonPropina = "02"
            wEnter = False
          
          Case Is = 18
            cmdOpcion(18).Caption = IIf(cmdOpcion(18).Caption = "Facturable", "Incluido", "Facturable")
            If cmdOpcion(18).Caption = "Facturable" Then
                sTipoComanda = "01"
            Else
                sTipoComanda = "02"
            End If
   End Select

End Sub

Private Sub Form_Load()
   Centrar Me
   wEnter = False
   frmVenta.lIncluido = False
   sTipoComanda = "01"
   sPropina = ""
   sMonPropina = ""
   
   If sReserva <> "" Then
      sFiltro = "AND vReserva.fSalida>='" & Format(FechaServidor(), "yyyy/MM/dd") & "'"
      Call confGrillaReserva
   Else
      Call confGrillaHabitacion
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsGrilla = Nothing
   Set frmCargo = Nothing
End Sub

Private Sub grdGrilla_HeadClick(ByVal ColIndex As Integer)
    Screen.MousePointer = vbHourglass
    RsGrilla.Sort = grdGrilla.Columns(ColIndex).DataField & " ASC"
    RsGrilla.Requery
    Screen.MousePointer = vbDefault
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If RsGrilla.EOF = True And RsGrilla.RecordCount <= 0 Then
      If sTipoCargo = "Habitacion" Then
         txtResultado = ""
         txtNacionalidad = ""
         txtCia = ""
         txtReserva = ""
      Else
         txtFIn = ""
         txtFOut = ""
      End If
      txtObservacion.Text = ""
   Else
      If sTipoCargo = "Habitacion" Then
         txtResultado = RsGrilla!Pasajero
         txtNacionalidad = IIf(IsNull(RsGrilla!Nacionalidad), "", RsGrilla!Nacionalidad)
         txtCia = IIf(IsNull(RsGrilla!compania), "", RsGrilla!compania)
         txtReserva = IIf(IsNull(RsGrilla!Reserva), "", RsGrilla!Reserva)
      Else
         txtResultado = RsGrilla!compania
         txtFIn = IIf(IsNull(RsGrilla!fLlegada), "", RsGrilla!fLlegada)
         txtFOut = IIf(IsNull(RsGrilla!fSalida), "", RsGrilla!fSalida)
      End If
      txtObservacion.Text = IIf(IsNull(RsGrilla!Observacion), "", RsGrilla!Observacion)
   End If
End Sub

Public Sub confGrillaHabitacion()
    sTipoCargo = "Habitacion"
    fraReserva.Visible = False
    fraHabitacion.Visible = True
    
    cmdOpcion(6).Caption = "Filtrar Habitación"
    cmdOpcion(7).Caption = "Filtra Pasajero"
    cmdOpcion(8).Visible = False
    
    InicializaGrilla grdGrilla
    Call ConfGrilla(2, grdGrilla, "Habitación", 2, "Habitacion", 1500, 2, 0, "", _
                                  "Pasajero", 2, "Pasajero", 4750, 0, 0, "")

    Isql = "SELECT  dbo.MPASAJERO.tNumeroHabitacion AS Habitacion, " & _
            "ISNULL(dbo.MPASAJERO.tPaterno, N'') + N' ' + ISNULL(dbo.MPASAJERO.tMaterno, N'') + N' ' + ISNULL(dbo.MPASAJERO.tNombre, N'') AS Pasajero, " & _
            "dbo.MPASAJERO.tCodigoReserva AS Reserva, " & _
            "dbo.TCOMPANIA.tRazonSocialCompania AS Compania, " & _
            "dbo.MRESERVA.tGrupo AS NombreReserva, " & _
            "dbo.TPAIS.tDescripcionNacionalidad AS Nacionalidad, " & _
            "dbo.THABITACIONHOTEL.lBloqueado, tCodigoPasajero, " & _
            "dbo.MRESERVA.mObservacion + + dbo.MPASAJERO.mObservacion as Observacion  " & _
            "From dbo.THABITACIONHOTEL " & _
            "RIGHT OUTER JOIN dbo.MPASAJERO ON dbo.THABITACIONHOTEL.tHotel = dbo.MPASAJERO.tHotel AND dbo.THABITACIONHOTEL.tNumeroHabitacion = dbo.MPASAJERO.tNumeroHabitacion " & _
            "LEFT OUTER JOIN  dbo.MRESERVA " & _
            "LEFT OUTER JOIN  dbo.TCOMPANIA ON dbo.MRESERVA.tCodigoCompania = dbo.TCOMPANIA.tCodigoCompania ON dbo.MPASAJERO.tCodigoReserva = dbo.MRESERVA.tCodigoReserva " & _
            "LEFT OUTER JOIN  dbo.TPAIS ON dbo.MPASAJERO.tNacionalidad = dbo.TPAIS.tCodigoPais " & _
            "Where (IsNull(dbo.MPASAJERO.tNumeroHabitacion, '0') <> '0') " & _
            "AND (ISNULL(dbo.MRESERVA.fRecepcion, 0) <> 0) AND (dbo.MPASAJERO.lCheckOut = 0) " & _
            "AND (dbo.MPASAJERO.tEstadoFicha = 'O') AND dbo.THABITACIONHOTEL.lBloqueado<>1 " & _
            "ORDER BY habitacion,Pasajero, dbo.MRESERVA.tCodigoReserva, dbo.MPASAJERO.tCodigoPasajero"
    Set RsGrilla = Lib.OpenRecordset(Isql, CnInfhotel)
    
    If sHabitacion <> "" Then
       RsGrilla.Filter = "Habitacion like *" & sHabitacion & "*"
    End If
    
    Set grdGrilla.DataSource = RsGrilla
End Sub

Public Sub confGrillaReserva()
    sTipoCargo = "Reserva"
    fraReserva.Visible = True
    fraHabitacion.Visible = False
    cmdOpcion(6).Caption = "Filtrar Reserva"
    cmdOpcion(7).Caption = "Filtra Compañia"
    cmdOpcion(8).Visible = True
    InicializaGrilla grdGrilla
    Call ConfGrilla(4, grdGrilla, "Reserva", 2, "Reserva", 850, 2, 0, "", _
                                  "Compañía", 2, "Compania", 2200, 0, 0, "", _
                                  "Nombre Reserva", 2, "Grupo", 2200, 0, 0, "", _
                                  "Tipo", 2, "tTipoReserva", 1000, 0, 0, "")
                                                                
    Isql = "SELECT vReserva.tCodigoReserva As Reserva,tgrupo As grupo,trazoncomercialcompania As Compania,fLlegada ,fSalida, (CASE WHEN vRESERVA.tTipoReserva = '1' THEN 'Habitación' ELSE 'Evento' END) as tTipoReserva,  " & _
           "(CASE WHEN vRESERVA.tTipoReserva = '3' THEN tObservacion + ' ' + mObservacion ELSE mObservacion END) as Observacion " & _
           "FROM dbo.vReserva LEFT OUTER JOIN dbo.MEVENTOSRESERVA ON dbo.vReserva.tHotel = dbo.MEVENTOSRESERVA.tHotel AND " & _
           "dbo.vReserva.tCodigoReserva = dbo.MEVENTOSRESERVA.tCodigoReserva LEFT OUTER JOIN dbo.vTipoReserva ON dbo.vReserva.tTipoReserva = dbo.vTipoReserva.tCodigo " & _
           "Where " & _
           "IsNull(vReserva.fRecepcion,0)<>0 " & _
           "AND vReserva.Estado<>'ANULADA' AND tCodigoReferencia<>'99999' " & _
           "AND trazoncomercialcompania<>'' " & _
           "AND vReserva.fLlegada<=GETDATE() "
    Isql = Isql & sFiltro
    Set RsGrilla = Lib.OpenRecordset(Isql, CnInfhotel)
    If sReserva <> "" Then
       RsGrilla.Filter = "Reserva like *" & sReserva & "*"
    End If
        
    Set grdGrilla.DataSource = RsGrilla
End Sub
