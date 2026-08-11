VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mesas"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   13440
   FillColor       =   &H8000000B&
   Icon            =   "frmIr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   13440
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Mesas / Observaciones"
      TabPicture(0)   =   "frmIr.frx":2582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdGrilla"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOpcion(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraMesa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraLocal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraLeyenda"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNavegar(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdNavegar(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdNavegar(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdNavegar(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdOpcion(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdOpcion(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Clientes Frecuentes / Socios"
      TabPicture(1)   =   "frmIr.frx":259E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdOpcion(10)"
      Tab(1).Control(1)=   "cmdOpcion(7)"
      Tab(1).Control(2)=   "cmdOpcion(9)"
      Tab(1).Control(3)=   "cmdOpcion(11)"
      Tab(1).Control(4)=   "cmdOpcion(4)"
      Tab(1).Control(5)=   "cmdOpcion(3)"
      Tab(1).Control(6)=   "Frame1"
      Tab(1).Control(7)=   "grdGrillaPedido"
      Tab(1).Control(8)=   "Label(12)"
      Tab(1).Control(9)=   "txtResultado"
      Tab(1).Control(10)=   "txtPredeterm"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton cmdOpcion 
         Height          =   600
         Index           =   10
         Left            =   -64440
         Picture         =   "frmIr.frx":25BA
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2790
         Width           =   1155
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   600
         Index           =   7
         Left            =   -64440
         Picture         =   "frmIr.frx":2E84
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   600
         Index           =   9
         Left            =   -64440
         Picture         =   "frmIr.frx":374E
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1590
         Width           =   1155
      End
      Begin VB.CommandButton cmdOpcion 
         Height          =   600
         Index           =   11
         Left            =   -64440
         Picture         =   "frmIr.frx":4018
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2190
         Width           =   1155
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
         Height          =   555
         Index           =   4
         Left            =   -64440
         Picture         =   "frmIr.frx":48E2
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4080
         Width           =   1155
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Seleccionar"
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
         Left            =   -64440
         Picture         =   "frmIr.frx":49D4
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3480
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   27
         Top             =   4590
         Width           =   11700
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            Left            =   9510
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   2430
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   42
            Top             =   2430
            Width           =   1725
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            Left            =   9510
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   180
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            Left            =   8640
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   180
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   39
            Top             =   180
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            Left            =   9510
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   930
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            Left            =   8640
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   930
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   36
            Top             =   930
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            Left            =   9510
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1680
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            Left            =   8640
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1680
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FFC0C0&
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
            TabIndex        =   33
            Top             =   1680
            Width           =   850
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FF8080&
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
            Left            =   10380
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   930
            Width           =   1250
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FF8080&
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
            Left            =   10380
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1680
            Width           =   1250
         End
         Begin VB.CommandButton cmdKey 
            BackColor       =   &H00FF8080&
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
            Left            =   10380
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   180
            Width           =   1250
         End
         Begin VB.CommandButton cmdKey 
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
            TabIndex        =   29
            Top             =   3195
            Width           =   11535
         End
         Begin VB.CommandButton cmdKey 
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
            TabIndex        =   28
            Top             =   180
            Width           =   750
         End
         Begin VB.TextBox txtTempo 
            Height          =   285
            Left            =   10170
            TabIndex        =   44
            Top             =   3375
            Width           =   765
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
         Height          =   555
         Index           =   0
         Left            =   10530
         Picture         =   "frmIr.frx":4AD6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8040
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Seleccionar"
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
         Left            =   10530
         Picture         =   "frmIr.frx":4BC8
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   7440
         Width           =   1275
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   8325
         Picture         =   "frmIr.frx":4CCA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   8040
         Width           =   855
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   7470
         Picture         =   "frmIr.frx":5594
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   8040
         Width           =   855
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   6615
         Picture         =   "frmIr.frx":5E5E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   8040
         Width           =   855
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   5760
         Picture         =   "frmIr.frx":6728
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   8040
         Width           =   855
      End
      Begin VB.Frame fraLeyenda 
         Caption         =   " Leyenda "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   10110
         TabIndex        =   7
         Top             =   5340
         Width           =   1695
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Reservada :"
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   19
            Top             =   870
            Width           =   870
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Bloqueada :"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   18
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Ocupada :"
            Height          =   195
            Index           =   10
            Left            =   255
            TabIndex        =   17
            Top             =   585
            Width           =   750
         End
         Begin VB.Label Label 
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   1050
            TabIndex        =   16
            Top             =   840
            Width           =   555
         End
         Begin VB.Label Label 
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   4
            Left            =   1050
            TabIndex        =   15
            Top             =   1410
            Width           =   555
         End
         Begin VB.Label Label 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   1050
            TabIndex        =   14
            Top             =   555
            Width           =   555
         End
         Begin VB.Label Label 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   1050
            TabIndex        =   13
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Lista :"
            Height          =   195
            Index           =   11
            Left            =   585
            TabIndex        =   12
            Top             =   300
            Width           =   420
         End
         Begin VB.Label Label 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   1050
            TabIndex        =   11
            Top             =   1125
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Sucia :"
            Height          =   195
            Index           =   8
            Left            =   510
            TabIndex        =   10
            Top             =   1155
            Width           =   495
         End
         Begin VB.Label Label 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   5
            Left            =   1050
            TabIndex        =   9
            Top             =   1695
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fuera Serv. :"
            Height          =   195
            Index           =   6
            Left            =   90
            TabIndex        =   8
            Top             =   1725
            Width           =   915
         End
      End
      Begin VB.Frame fraLocal 
         Caption         =   " Locales "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4950
         Left            =   10110
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton cmdBusca 
            Height          =   555
            Left            =   205
            Picture         =   "frmIr.frx":6FF2
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4140
            Width           =   1275
         End
         Begin VB.CommandButton cmdSalon 
            Height          =   555
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   330
            Width           =   1275
         End
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
         TabIndex        =   2
         Top             =   360
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
            TabIndex        =   3
            Top             =   420
            Width           =   915
         End
         Begin VB.Image imgFondo 
            Height          =   7875
            Left            =   120
            Stretch         =   -1  'True
            Top             =   270
            Width           =   5505
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Buscar"
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
         Left            =   9210
         Picture         =   "frmIr.frx":757C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   8040
         Width           =   1275
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   7590
         Left            =   5760
         TabIndex        =   26
         Top             =   390
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   13388
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
         Caption         =   "Cuentas Abiertas"
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
      Begin TrueOleDBGrid80.TDBGrid grdGrillaPedido 
         Height          =   3765
         Left            =   -74850
         TabIndex        =   45
         Top             =   810
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   6641
         _LayoutType     =   4
         _RowHeight      =   25
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
      Begin VB.Label Label 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   -72990
         TabIndex        =   48
         Top             =   405
         Width           =   150
      End
      Begin VB.Label txtResultado 
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
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   -72810
         TabIndex        =   47
         Top             =   360
         Width           =   7395
      End
      Begin VB.Label txtPredeterm 
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
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmIr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMesa As Recordset
Dim RsSalon As Recordset
Dim RsPedido As Recordset
Dim RsGrilla As Recordset
Dim sMesaActual As String
Dim sTitulo As String
Dim nIndex As Integer
Dim xIndex As Integer
Dim xSalon As String
Dim sMesa() As String
Dim Estado As String

Dim RsGrillaPedido As Recordset
Public nPredeterm As Integer
Dim sCriterio As String

Private Sub cmdBusca_Click()
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    
    sTemp = ""
    Isql = "select * from vSalon where lActivo = 1"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       sTitulo = Trim(Right(cmdMesa(nIndex).Caption, InStr(StrReverse(cmdMesa(nIndex).Caption), " ")))
    
       RsSalon.MoveFirst
       RsSalon.Find "tSalon='" & sCodigo & "'"
       fraMesa.Caption = RsSalon!dSalon
       InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
       Desasigna (RsMesa.RecordCount)
       AsignaMesa
    End If
End Sub

Private Sub cmdSalon_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   'On Error Resume Next
   sTitulo = Trim(Right(cmdMesa(nIndex).Caption, InStr(StrReverse(cmdMesa(nIndex).Caption), " ")))

   RsSalon.MoveFirst
   RsSalon.Move Index - 1

   fraMesa.Caption = RsSalon!dSalon
   InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
   Desasigna (RsMesa.RecordCount)
   AsignaMesa
End Sub

Private Sub cmdMesa_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    wEnter = True
    RsMesa.MoveFirst
    RsMesa.Move (Index - 1)
    sCodigo = RsMesa!tCodigoMesa
         lTabBuscar = False
    Unload Me
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case Index
          Case Is = 0  'Salir
               Unload Me
           Case Is = 4
                Unload Me
                
          Case Is = 1 'Observacion
               frmKeyBoard.Show vbModal
               If wEnter Then
                  If Len(Trim(sDescrip)) = 0 Then Exit Sub
                     Sw = True
               End If
               wEnter = False
               Unload Me
          
           Case Is = 3 'Observacion
              If RsGrillaPedido.RecordCount = 0 Then
                   Exit Sub
                End If
                wEnter = True
                     lTabBuscar = True
                sCodigo = IIf(RsGrillaPedido.EOF = True, "", RsGrillaPedido!Pedido)
                Unload Me
          
          
          Case Is = 2 'Seleccionar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               sDescrip = grdGrilla.Columns(0).value
                    lTabBuscar = False
               Sw = True
               wEnter = False
               Unload Me
             
          Case Is = 7  ' Primero
               MoverPuntero Primero, grdGrillaPedido
               txtTempo.SetFocus

          Case Is = 8 ' PG Up
               MoverPuntero pgup, grdGrillaPedido
               txtTempo.SetFocus

          Case Is = 9 ' Previo
               MoverPuntero previo, grdGrillaPedido
               txtTempo.SetFocus

          Case Is = 11  ' Siguiente
               MoverPuntero siguiente, grdGrillaPedido
               txtTempo.SetFocus

          Case Is = 12  ' PG Dn
               MoverPuntero pgdn, grdGrillaPedido
               txtTempo.SetFocus

          Case Is = 10 ' Ultimo
               MoverPuntero Ultimo, grdGrillaPedido
               txtTempo.SetFocus
   End Select
End Sub

Private Sub Form_Activate()
 txtResultado.Caption = sTemp
 If lTabBuscar = True Then
    Me.SSTab1.Tab = 1
 Else
    Me.SSTab1.Tab = 0
 End If
    If lBuscarPedidoVisualizarGrilla = True Then
     Me.SSTab1.Tab = 1
    Else
     Me.SSTab1.Tab = 0
    End If
End Sub

Private Sub Form_Load()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Dim i As Integer
   On Error Resume Next
   Sw = False
   Set RsMesa = Lib.OpenRecordset("select * from TMESA where tCodigoMesa <> '000' and lActivo= 1 order by tCodigoMesa", Cn)
   Set RsSalon = Lib.OpenRecordset("select tCodigo as tSalon, tDetallado as dSalon, tiCono from TTABLA where tTABLA ='Salon' and lActivo = 1", Cn)
   Set RsPedido = Lib.OpenRecordset("select tCodigoPedido as Codigo, tMesa, fFecha, nPrecuenta from MPEDIDO where tEstadoPedido = '01' order by tMesa", Cn)
   Set RsGrilla = Lib.OpenRecordset("select tObservacion from MPEDIDO where tEstadoPedido = '01' and not len(rtrim(tObservacion)) = 0 and tCaja = '" & sCaja & "' order by tObservacion", Cn)
   
   'Configuración de Grillas
   Call ConfGrilla(1, grdGrilla, "Mesa", 2, "tObservacion", 3650, 0, 0, "")
   
   Set grdGrilla.DataSource = RsGrilla
      
   'Salon
   Call MatrizBotones(IIf(RsSalon.RecordCount >= 6, 6, RsSalon.RecordCount), 1, 60, 60, cmdSalon)
   For i = 1 To RsSalon.RecordCount
       cmdSalon(i).Caption = RsSalon!dSalon
       RsSalon.MoveNext
   Next i
   RsSalon.MoveFirst
   RsSalon.Find ("tSalon='" & sSalon & "'")
   If RsSalon.EOF Then
      RsSalon.MoveFirst
   End If
   xSalon = RsSalon!tSalon
   
   fraMesa.Caption = RsSalon!dSalon
   InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
   
   RsPedido.MoveFirst
   RsPedido.Find ("Codigo ='" & sPedido & "'")

   AsignaMesa
   
   
   'otrapestaña
   MatrizBotones 40, 10, 1, 1, cmdKey
   LlenaKey
   If lBuscarPedidoFiltrarMesa = True Then
        Set RsGrillaPedido = Lib.OpenRecordset("exec usp_Inforest_ObtienePedidosPorMesas", Cn)
   Else
        Set RsGrillaPedido = Lib.OpenRecordset("exec usp_Inforest_ObtienePedidosPorSocios", Cn)
   End If
   '
   Set grdGrillaPedido.DataSource = RsGrillaPedido
   'otrapestaña
   Screen.MousePointer = vbDefault
End Sub

Public Sub Desasigna(Cantidad As Integer)
   Dim i As Integer
   For i = 1 To Cantidad
       Unload cmdMesa(i)
   Next i
End Sub

Private Sub Form_Paint()
  txtPredeterm.Caption = grdGrillaPedido.Columns(nPredeterm).Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsMesa = Nothing
   Set RsSalon = Nothing
   Set RsPedido = Nothing
   Set frmIr = Nothing
End Sub

Public Sub AsignaMesa()
   Dim i As Integer
   Dim nHora As Long
   
   RsMesa.Filter = "tSalon = '" & RsSalon!tSalon & "'"
   
   If RsMesa.RecordCount = 0 Then
      Exit Sub
   End If
   
   RsMesa.MoveFirst
   cmdMesa(0).Visible = False
   ReDim sMesa(RsMesa.RecordCount)
   
   For i = 1 To RsMesa.RecordCount
       'Cantidad de Mesas
       With cmdMesa(i)
            Load cmdMesa(i)
            .Top = RsMesa!tX
            .Left = RsMesa!tY
            .Visible = True
            cmdMesa(i).Caption = Trim(RsMesa!tResumido)
            sMesa(i) = Trim(RsMesa!tResumido)
            cmdMesa(i).Enabled = False
            
            Select Case RsMesa!tEstadoMesa
                   Case Is = "01" ' Lista
                        cmdMesa(i).backColor = vbLista
                        
                   Case Is = "02" ' Ocupada
                        If RsPedido.RecordCount <> 0 Then
                           RsPedido.MoveFirst
                           RsPedido.Find ("tMesa = '" & RsMesa!tCodigoMesa & "'")
                           cmdMesa(i).Enabled = True
                           
                           If Not RsPedido.EOF Then
                              nHora = DateDiff("n", RsPedido!fFecha, FechaServidor())
                              sTitulo = Format(Format(Trim(str(Int(nHora / 60))), "00") & ":" & Format(Trim(str(nHora Mod 60)), "00"), "hh:mm")
                              cmdMesa(i).Caption = "" & cmdMesa(i).Caption & " " & sTitulo
                              
                              cmdMesa(i).backColor = IIf(IsNull(RsPedido!nPrecuenta), vbOcupada, vbPrecuentaImp)
                              'cmdMesa(i).BackColor = vbOcupada
                           Else
                              cmdMesa(i).backColor = IIf(IsNull(RsPedido!nPrecuenta), vbOcupada, vbPrecuentaImp)
                              'cmdMesa(i).BackColor = vbOcupada
                           End If
                        End If
                        
                   Case Is = "03" ' Reservada
                        cmdMesa(i).backColor = vbReservada
                   Case Is = "04" ' Sucia
                        cmdMesa(i).backColor = vbSucia
                   Case Is = "05" ' Bloqueada
                        cmdMesa(i).backColor = vbBloqueada
                   Case Is = "06" ' Fuera de Servicio
                        cmdMesa(i).backColor = vbFServicio
            End Select
            
       End With
       RsMesa.MoveNext
   Next i
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

Private Sub cmdkey_Click(Index As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    
    Select Case Index
           Case Is = 41 'Escape
                wEnter = False
                Unload Me
           
           Case Is = 42 'Backspace
                If Len(Trim(txtResultado.Caption)) > 0 Then
                   txtResultado.Caption = Mid(txtResultado.Caption, 1, Len(txtResultado.Caption) - 1)
                Else
                   txtResultado.Caption = ""
                End If
                txtTempo.SetFocus
                
           Case Is = 43 'Enter
                If RsGrillaPedido.RecordCount = 0 Then
                   Exit Sub
                End If
                wEnter = True
                lCargaDesdePedido = True
                sCodigo = IIf(RsGrillaPedido.EOF = True, "", RsGrillaPedido!Pedido)
                lCargaDesdePedido = True
                     lTabBuscar = True
                Unload Me
                
           Case Is = 44 'Espacio
                txtResultado.Caption = txtResultado.Caption + " "
                
           Case Else
                If cmdKey(Index).Caption = "&&" Then
                   txtResultado.Caption = txtResultado.Caption + "&"
                Else
                   txtResultado.Caption = txtResultado.Caption + cmdKey(Index).Caption
                End If
                txtTempo.SetFocus
    End Select
    
    '       '    numeros   letras    sp  numpad     delet  ñ
    '       Case 48 To 57, 65 To 90, 32, 96 To 105, 8, 46, 192
    '            Filtrar
End Sub

Private Sub txtresultado_Change()
   Filtrar
End Sub
Private Sub grdGrillaPedido_Click()
If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
End Sub

Private Sub grdGrillaPedido_DblClick()
 If RsGrillaPedido.RecordCount = 0 Then
      Exit Sub
   End If
     sCodigo = IIf(RsGrillaPedido.EOF = True, "", RsGrillaPedido!Pedido)
     lCargaDesdePedido = True
     lTabBuscar = True
   wEnter = True
   Unload Me
End Sub

Private Sub grdGrillaPedido_HeadClick(ByVal ColIndex As Integer)
    If sModulo = "ADICION" Then
       frmMozoUsuario.ReseteaTimer
    End If
    Screen.MousePointer = vbHourglass
    If Right(RsGrillaPedido.Sort, 3) = "ASC" Then
       RsGrillaPedido.Sort = grdGrillaPedido.Columns(ColIndex).DataField & " DESC"
    Else
       RsGrillaPedido.Sort = grdGrillaPedido.Columns(ColIndex).DataField & " ASC"
    End If
    RsGrillaPedido.Requery
    Screen.MousePointer = vbDefault
    nPredeterm = ColIndex
    txtPredeterm.Caption = grdGrillaPedido.Columns(nPredeterm).Caption
    txtTempo.SetFocus
End Sub

Private Sub grdGrillaPedido_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   txtTempo.SetFocus
End Sub
Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   If Shift > 0 Then
      Select Case KeyCode
       Case 48
            txtResultado.Caption = txtResultado.Caption + "="
       Case 49
            txtResultado.Caption = txtResultado.Caption + "!"
       Case 50
            txtResultado.Caption = txtResultado.Caption + "_"
       Case 51
            txtResultado.Caption = txtResultado.Caption + "#"
       Case 52
            txtResultado.Caption = txtResultado.Caption + "$"
       Case 53
            txtResultado.Caption = txtResultado.Caption + "%"
       Case 54
            txtResultado.Caption = txtResultado.Caption + "&"
       Case 55
            txtResultado.Caption = txtResultado.Caption + "/"
       Case 56
            txtResultado.Caption = txtResultado.Caption + "("
       Case 57
            txtResultado.Caption = txtResultado.Caption + ")"
       '-----para los valores no repetidos------------------------------------
        Case 32
             txtResultado.Caption = txtResultado.Caption + " "
        Case 65
             txtResultado.Caption = txtResultado.Caption + "A"
        Case 66
             txtResultado.Caption = txtResultado.Caption + "B"
        Case 67
             txtResultado.Caption = txtResultado.Caption + "C"
        Case 68
             txtResultado.Caption = txtResultado.Caption + "D"
        Case 69
             txtResultado.Caption = txtResultado.Caption + "E"
        Case 70
             txtResultado.Caption = txtResultado.Caption + "F"
        Case 71
             txtResultado.Caption = txtResultado.Caption + "G"
        Case 72
             txtResultado.Caption = txtResultado.Caption + "H"
        Case 73
             txtResultado.Caption = txtResultado.Caption + "I"
        Case 74
             txtResultado.Caption = txtResultado.Caption + "J"
        Case 75
             txtResultado.Caption = txtResultado.Caption + "K"
        Case 76
             txtResultado.Caption = txtResultado.Caption + "L"
        Case 77
             txtResultado.Caption = txtResultado.Caption + "M"
        Case 78
             txtResultado.Caption = txtResultado.Caption + "N"
        Case 79
             txtResultado.Caption = txtResultado.Caption + "O"
        Case 80
             txtResultado.Caption = txtResultado.Caption + "P"
        Case 81
             txtResultado.Caption = txtResultado.Caption + "Q"
        Case 82
             txtResultado.Caption = txtResultado.Caption + "R"
        Case 83
             txtResultado.Caption = txtResultado.Caption + "S"
        Case 84
             txtResultado.Caption = txtResultado.Caption + "T"
        Case 85
             txtResultado.Caption = txtResultado.Caption + "U"
        Case 86
             txtResultado.Caption = txtResultado.Caption + "V"
        Case 87
             txtResultado.Caption = txtResultado.Caption + "W"
        Case 88
            txtResultado.Caption = txtResultado.Caption + "X"
        Case 89
             txtResultado.Caption = txtResultado.Caption + "Y"
        Case 90
             txtResultado.Caption = txtResultado.Caption + "Z"
        Case 192
             txtResultado.Caption = txtResultado.Caption + "Ñ"
    '-----------------------------------------------------------------------------------
       'Case Else
       '     MsgBox KeyCode
      End Select
   Else
      Select Case KeyCode
          Case 13
               Call cmdkey_Click(43)
          Case 27
               Call cmdkey_Click(41)
          Case 8
               Call cmdkey_Click(42)
          Case 32
               txtResultado.Caption = txtResultado.Caption + " "
          Case 65
               txtResultado.Caption = txtResultado.Caption + "A"
          Case 66
               txtResultado.Caption = txtResultado.Caption + "B"
          Case 67
               txtResultado.Caption = txtResultado.Caption + "C"
          Case 68
               txtResultado.Caption = txtResultado.Caption + "D"
          Case 69
               txtResultado.Caption = txtResultado.Caption + "E"
          Case 70
               txtResultado.Caption = txtResultado.Caption + "F"
          Case 71
               txtResultado.Caption = txtResultado.Caption + "G"
          Case 72
               txtResultado.Caption = txtResultado.Caption + "H"
          Case 73
               txtResultado.Caption = txtResultado.Caption + "I"
          Case 74
               txtResultado.Caption = txtResultado.Caption + "J"
          Case 75
               txtResultado.Caption = txtResultado.Caption + "K"
          Case 76
               txtResultado.Caption = txtResultado.Caption + "L"
          Case 77
               txtResultado.Caption = txtResultado.Caption + "M"
          Case 78
               txtResultado.Caption = txtResultado.Caption + "N"
          Case 79
               txtResultado.Caption = txtResultado.Caption + "O"
          Case 80
               txtResultado.Caption = txtResultado.Caption + "P"
          Case 81
               txtResultado.Caption = txtResultado.Caption + "Q"
          Case 82
               txtResultado.Caption = txtResultado.Caption + "R"
          Case 83
               txtResultado.Caption = txtResultado.Caption + "S"
          Case 84
               txtResultado.Caption = txtResultado.Caption + "T"
          Case 85
               txtResultado.Caption = txtResultado.Caption + "U"
          Case 86
               txtResultado.Caption = txtResultado.Caption + "V"
          Case 87
               txtResultado.Caption = txtResultado.Caption + "W"
          Case 88
              txtResultado.Caption = txtResultado.Caption + "X"
          Case 89
               txtResultado.Caption = txtResultado.Caption + "Y"
          Case 90
               txtResultado.Caption = txtResultado.Caption + "Z"
          Case 96, 48
               txtResultado.Caption = txtResultado.Caption + "0"
          Case 97, 49
               txtResultado.Caption = txtResultado.Caption + "1"
          Case 98, 50
               txtResultado.Caption = txtResultado.Caption + "2"
          Case 99, 51
               txtResultado.Caption = txtResultado.Caption + "3"
          Case 100, 52
               txtResultado.Caption = txtResultado.Caption + "4"
          Case 101, 53
               txtResultado.Caption = txtResultado.Caption + "5"
          Case 102, 54
               txtResultado.Caption = txtResultado.Caption + "6"
          Case 103, 55
               txtResultado.Caption = txtResultado.Caption + "7"
          Case 104, 56
               txtResultado.Caption = txtResultado.Caption + "8"
          Case 105, 57
               txtResultado.Caption = txtResultado.Caption + "9"
          Case 110, 190
               txtResultado.Caption = txtResultado.Caption + "."
          Case 106, 188
               txtResultado.Caption = txtResultado.Caption + ","
          Case 109, 189
               txtResultado.Caption = txtResultado.Caption + "-"
          Case 192
               txtResultado.Caption = txtResultado.Caption + "Ñ"
'          Case Else
'               MsgBox KeyCode
      End Select
    End If
End Sub

Public Sub Filtrar()
On Error Resume Next
   If Len(Trim(txtResultado.Caption)) > 0 Then
      Select Case VarType(grdGrillaPedido.Columns(nPredeterm).value)
             Case 2 To 6
            ' MsgBox Val(txtResultado)
             sCriterio = Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & "= " & val(txtResultado.Caption)
             Case 7
             sCriterio = Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & " >= #" & txtResultado.Caption & "# and " & Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & " <= #" & txtResultado.Caption & " 23:59#"
            Case 1 ' es nulo
                GoTo nulo
             Case Else
             sCriterio = Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & " Like " & "'*" & txtResultado.Caption & "*'"
      End Select
      Screen.MousePointer = vbHourglass
      With RsGrillaPedido
           .Filter = sCriterio
           If .BOF = False Then
              .MoveFirst
           End If
      End With
      Screen.MousePointer = vbDefault
      Exit Sub
      'cmdTexto.Caption = "Total : " & Format(RsGrilla.RecordCount, "##,##0") & " de " & Format(nTotal, "##,##0") & " Registros"
nulo:
        sCriterio = Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & "= " & val(txtResultado.Caption)
        RsGrillaPedido.Filter = sCriterio
        If RsGrillaPedido.BOF = True Then
             sCriterio = Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & " Like " & "'*" & txtResultado.Caption & "*'"
             RsGrillaPedido.Filter = sCriterio
             If RsGrillaPedido.BOF = True Then
                 sCriterio = Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & " >= #" & txtResultado.Caption & "# and " & Trim(grdGrillaPedido.Columns(nPredeterm).DataField) & " <= #" & txtResultado.Caption & " 23:59#"
                 RsGrillaPedido.Filter = sCriterio
             End If
        End If
      Screen.MousePointer = vbDefault
      Exit Sub
   Else
      Screen.MousePointer = vbHourglass
      RsGrillaPedido.Filter = adFilterNone
      RsGrillaPedido.Requery
      If RsGrillaPedido.RecordCount <> 0 Then
         RsGrillaPedido.MoveFirst
      End If
      'cmdTexto.Caption = "Total : " & Format(nTotal, "##,##0") & " de " & Format(nTotal, "##,##0") & " Registros"
      Screen.MousePointer = vbDefault
      sCriterio = ""
   End If
End Sub


