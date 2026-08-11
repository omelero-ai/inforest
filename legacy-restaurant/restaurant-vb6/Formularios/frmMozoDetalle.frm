VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{675D2E3C-5FBF-4F10-8213-EE25D71049AC}#2.3#0"; "sgfplibx.ocx"
Begin VB.Form frmMozoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4560
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11865
   Icon            =   "frmMozoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   11865
   Begin VB.CheckBox chklimportacionreservas 
      Caption         =   "Autoriza importación de reservas"
      Height          =   195
      Left            =   3000
      TabIndex        =   66
      Top             =   3240
      Width           =   3675
   End
   Begin VB.Frame fraBoton 
      Caption         =   " Botonera "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   8940
      TabIndex        =   12
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdNA 
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2310
         TabIndex        =   32
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   19
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   18
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   17
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   16
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2160
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   15
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   14
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   13
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   12
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1530
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   10
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   9
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   8
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   7
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   900
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   5
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   3
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   270
         Width           =   510
      End
      Begin VB.Label txtBoton 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1830
         TabIndex        =   34
         Top             =   2820
         Width           =   990
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   1095
         TabIndex        =   33
         Top             =   2865
         Width           =   510
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11805
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3810
      Width           =   11865
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6060
         TabIndex        =   5
         Top             =   60
         Width           =   6120
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4620
            Picture         =   "frmMozoDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5100
            Picture         =   "frmMozoDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5580
            Picture         =   "frmMozoDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmMozoDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmMozoDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmMozoDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   6
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
            Left            =   1440
            TabIndex        =   35
            Top             =   150
            Width           =   3105
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
         Left            =   7410
         Picture         =   "frmMozoDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   6240
         Picture         =   "frmMozoDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Left            =   8580
         Picture         =   "frmMozoDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   9750
         Picture         =   "frmMozoDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   1170
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   1440
      TabIndex        =   36
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmMozoDetalle.frx":3026
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkCambia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtBanda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkVerPedido"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPassword"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDetallado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtResumido"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkActivo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkAutorizaEnviosMozo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Huella Dactilar"
      TabPicture(1)   =   "frmMozoDetalle.frx":3042
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSecugen"
      Tab(1).Control(1)=   "fraDigitalPersona"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkAutorizaEnviosMozo 
         Caption         =   "Autoriza envíos de productos a producción"
         Height          =   195
         Left            =   1545
         TabIndex        =   65
         Top             =   2880
         Width           =   3675
      End
      Begin VB.Frame fraDigitalPersona 
         Caption         =   "Digital Persona"
         Height          =   3375
         Left            =   -74960
         TabIndex        =   50
         Top             =   360
         Width           =   7335
         Begin VB.PictureBox HiddenPict 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   495
            Left            =   5160
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   64
            Top             =   1920
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ListBox Status 
            Height          =   2400
            Left            =   2640
            TabIndex        =   53
            Top             =   600
            Width           =   4575
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   2600
            Left            =   0
            ScaleHeight     =   2535
            ScaleWidth      =   2535
            TabIndex        =   51
            Top             =   240
            Width           =   2600
            Begin VB.Label lblHuella 
               Alignment       =   2  'Center
               Caption         =   "Huella Dactilar Registrada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   855
               Left            =   0
               TabIndex        =   52
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Indicaciones :"
            Height          =   255
            Left            =   2640
            TabIndex        =   57
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Prompt 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Toque el lector de Huella Dactilar"
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
            Left            =   2640
            TabIndex        =   56
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Samples 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   55
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Huellas Dactilares Necesarias :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   2880
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   5640
         TabIndex        =   44
         Top             =   2520
         Width           =   840
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
         Left            =   1545
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   480
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
         Left            =   1545
         MaxLength       =   15
         TabIndex        =   42
         Text            =   " "
         Top             =   1305
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
         Left            =   1545
         MaxLength       =   50
         TabIndex        =   41
         Top             =   885
         Width           =   4950
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1545
         MaxLength       =   24
         PasswordChar    =   "*"
         TabIndex        =   40
         Text            =   " "
         Top             =   1710
         Width           =   2595
      End
      Begin VB.CheckBox chkVerPedido 
         Caption         =   "Permite ver todas las cuentas"
         Height          =   195
         Left            =   1545
         TabIndex        =   39
         Top             =   2535
         Width           =   2595
      End
      Begin VB.TextBox txtBanda 
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
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   38
         Text            =   " "
         Top             =   2130
         Width           =   2595
      End
      Begin VB.CheckBox chkCambia 
         Alignment       =   1  'Right Justify
         Caption         =   "Mensaje debes Cambiar Clave"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4275
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   510
         Width           =   2220
      End
      Begin VB.Frame fraSecugen 
         Caption         =   "SecuGen Hamster Plus (Model  HSDU03P)"
         Height          =   3375
         Left            =   -74960
         TabIndex        =   58
         Top             =   360
         Width           =   7335
         Begin SGFPLIBXLib.FpLibXCapture FpLibXCapture1 
            Height          =   2295
            Index           =   1
            Left            =   3840
            TabIndex        =   59
            Top             =   600
            Width           =   2715
            _Version        =   131075
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   64
            Brightness      =   20306
            Contrast        =   18758
            Gain            =   17740
         End
         Begin SGFPLIBXLib.FpLibXCapture FpLibXCapture1 
            Height          =   2295
            Index           =   0
            Left            =   960
            TabIndex        =   60
            Top             =   600
            Width           =   2715
            _Version        =   131075
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   64
            Brightness      =   -30584
            Contrast        =   -30584
            Gain            =   -31096
         End
         Begin SGFPLIBXLib.FpLibXVerify FpLibXVerify1 
            Left            =   6720
            Top             =   2280
            _Version        =   131075
            _ExtentX        =   1058
            _ExtentY        =   1058
            _StockProps     =   0
         End
         Begin VB.Label LabelTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2nd Huella Dactilar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   63
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label LabelTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1st Huella Dactilar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   62
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label LabelMessage 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   2925
            Width           =   7140
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   3
         Left            =   885
         TabIndex        =   49
         Top             =   525
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   48
         Top             =   1350
         Width           =   1260
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Detallada :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   47
         Top             =   930
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   46
         Top             =   1755
         Width           =   780
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Banda Magnética :"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Top             =   2175
         Width           =   1350
      End
   End
   Begin VB.Image Image 
      Height          =   3720
      Left            =   15
      Picture         =   "frmMozoDetalle.frx":305E
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1380
   End
End
Attribute VB_Name = "frmMozoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim RsBoton As Recordset
Dim sBanda As String

'HUELLA
Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim CreateFtrs As DPFPFeatureExtraction
Dim CreateTempl As DPFPEnrollment
Dim ConvertSample As DPFPSampleConversion
Dim Templ As DPFPTemplate
Dim blob() As Byte
Dim rHuella As Object


Dim lGrabaHuella As Boolean
Dim X As String

Dim xNuevo As String

' General error
Dim i As Integer
Const ERROR_NONE = 0
Const ERROR_CREATION_FAILED = 1
Const ERROR_FUNCTION_FAILED = 2
Const ERROR_INVALID_PARAM = 3
Const ERROR_NOT_USED = 4
Const ERROR_DLLLOAD_FAILED = 5
Const ERROR_DLLLOAD_FAILED_DRV = 6
Const ERROR_DLLLOAD_FAILED_ALGO = 7
      'Dim Cn As Connection
                   
' Device error
Const ERROR_SYSLOAD_FAILED = 51     'system file load fail
Const ERROR_INITIALIZE_FAILED = 52  'chip initialize fail
Const ERROR_LINE_DROPPED = 53       'image data drop
Const ERROR_TIME_OUT = 54           'getliveimage timeout error
Const ERROR_DEVICE_NOT_FOUND = 55   'device not found
Const ERROR_DRVLOAD_FAILED = 56     'dll file load fail
Const ERROR_WRONG_IMAGE = 57         'wrong image
Const ERROR_LACK_OF_BANDWIDTH = 58   'USB Bandwith Lack Error
Const ERROR_DEV_ALREADY_OPEN = 59     'Device Exclusive access Error
Const ERROR_GETSN_FAILED = 60         'Fail to get Device Serial Number
Const ERROR_UNSUPPORTED_DEV = 61      ' Unsupported device
   
   
' Extract&verification error
Const ERROR_FEAT_NUMBER = 101               'too small number of minutiae
Const ERROR_INVALID_TEMPLATE_TYPE = 102     'wrong template type
Const ERROR_INVALID_TEMPLATE1 = 103         'error in decoding template 1
Const ERROR_INVALID_TEMPLATE2 = 104         'error in decoding template 2
Const ERROR_EXTRACT_FAIL = 105
Const ERROR_MATCH_FAIL = 106

' Template Format
Const TEMPLATE_FORMAT_ANSI378 = 256
Const TEMPLATE_FORMAT_SG400 = 512
Const TEMPLATE_FORMAT_ISO19794 = 768


Dim g_FirstMinData() As Byte
Dim g_SecondMinData() As Byte
Dim g_MatchingMinData() As Byte
Dim g_firstStep As Boolean
Dim g_secondStep As Boolean
Dim g_vefiryResult As Boolean

Sub Asignar()
    With frmMozo.RsCabecera
         'Cuadro de Texto
         txtCodigo = IIf(IsNull(!codigo), "", !codigo)
         txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
         txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
         txtBoton = IIf(IsNull(!nBoton), "", !nBoton)
         txtPassword = Desencapsula(IIf(IsNull(!tvalor), "", !tvalor))
         txtBanda = Desencapsula(IIf(IsNull(!tBandaMagnetica), "", !tBandaMagnetica))
         'Check Box
         chkVerPedido = IIf(IsNull(!nValor), 0, !nValor)
         chkActivo = IIf(!lActivo = True, 1, 0)
         chkCambia = IIf(IsNull(!nTamano), 0, !nTamano)
         
         chkAutorizaEnviosMozo = IIf(!lAutorizaEnviosMozo = True, 1, 0)
         Me.chklimportacionreservas = IIf(!lImportacionReservas = True, 1, 0)
         
         
         'HUELLA
        If !thuella <> "" Then
            LabelMessage.Caption = "Huella Dactilar Registrada"
            lblHuella.Visible = True
        Else
            LabelMessage.Caption = ""
            lblHuella.Visible = False
        End If
         
         Botonera
    End With
End Sub

Private Sub cmdBoton_Click(Index As Integer)
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
   End If
   cmdBoton(Index).backColor = vbRed
   cmdBoton(Index).Enabled = False
   txtBoton.Caption = Index
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmMozo.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmMozo.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmMozo.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmMozo.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmMozo.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmMozo.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmMozo.RsCabecera.RecordCount = 0, 0, frmMozo.RsCabecera.AbsolutePosition) & " de " & frmMozo.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               Botonera
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               
               'Chequea Datos
               If LTrim(txtDetallado.Text) = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If LTrim(txtResumido.Text) = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                       
                       
               'HUELLA
               If lGrabaHuella = True Then
               Dim X As String
               Dim blob() As Byte
               If lHuellaDigitalPersona = True Then
                        If Templ Is Nothing Then
                        X = ""
                        
                        Else
                             
                                 blob = Templ.Serialize
                                 X = arraytohex(blob)
                        End If
                End If
                If lHuellaSecugen = True Then
                    X = arraytohex(g_FirstMinData)
                End If
               '----------------
                End If
                       
               If Sw Then
                  If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                   If Calcular("select count(tPassword) as codigo from tusuario where tpassword='" & Encapsula(txtPassword.Text) & "'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as codigo from tTabla where tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as codigo from tTabla where tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(TBANDAMAGNETICA) as codigo from TUSUARIO where tbandamagnetica='" & Encapsula(txtBanda.Text) & "'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                             
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='MOZO' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "0001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 4)
                  End If
                  Sw = False
                  sBanda = Extrae(txtBanda.Text)
                  
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nboton, tValor, nValor, tIcono, nTamano, lAutorizaEnviosMozo, lActivo,lImportacionReservas) " & _
                         "values ('MOZO', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       val(txtBoton.Caption) & ", " & _
                                " '" & Encapsula(txtPassword.Text) & "', " & _
                                       chkVerPedido.Value & ", " & _
                                " '" & Encapsula(sBanda) & "', " & _
                                       chkCambia.Value & ", " & _
                                       chkAutorizaEnviosMozo.Value & ", " & _
                                       chkActivo.Value & ", " & _
                                       chklimportacionreservas.Value & ") "
           
                  Cn.Execute Isql
                  
                  If lGrabaHuella = True Then
                    Isql = "update ttabla set tvalor2='" & X & "' where ttabla='MOZO' and tcodigo='" & txtCodigo.Text & "'"
                    Cn.Execute Isql
                  End If
                  RsBoton.Requery
                  frmMozo.RsCabecera.Sort = "CODIGO ASC"
                  frmMozo.RsCabecera.Requery
                  frmMozo.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmMozo.RsCabecera.RecordCount = 0, 0, frmMozo.RsCabecera.AbsolutePosition) & " de " & frmMozo.RsCabecera.RecordCount
               Else
                  If txtPassword.Text <> "" And Calcular("select count(tValor) as codigo from tTabla where tCodigo<>'" & txtCodigo & "' and tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                   If Calcular("select count(tValor) as codigo from tTabla where tValor='" & Encapsula(txtPassword.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  If Calcular("select count(tPassword) as codigo from tusuario where tpassword='" & Encapsula(txtPassword.Text) & "'", Cn) > 0 Then
                     MsgBox "Clave no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtPassword.SetFocus
                     Exit Sub
                  End If
                  
                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as Codigo from tTabla where tCodigo<>'" & txtCodigo & "' and tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MOZO'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If

                  If txtBanda.Text <> "" And Calcular("select count(tIcono) as codigo from tTabla where tIcono='" & Encapsula(txtBanda.Text) & "' and tTabla='MAITRE'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  If txtBanda.Text <> "" And Calcular("select count(TBANDAMAGNETICA) as codigo from TUSUARIO where tbandamagnetica='" & Encapsula(txtBanda.Text) & "'", Cn) > 0 Then
                     MsgBox "Banda Magnética no permitida, intente de nuevo", vbExclamation, sMensaje
                     txtBanda.SetFocus
                     Exit Sub
                  End If
                  
                  sBanda = Extrae(txtBanda.Text)
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nBoton =" & val(txtBoton.Caption) & ", " & _
                         "tValor = '" & Encapsula(txtPassword.Text) & "', " & _
                         "tIcono = '" & Encapsula(sBanda) & "', " & _
                         "nValor =" & chkVerPedido.Value & ", " & _
                         "lImportacionReservas =" & chklimportacionreservas.Value & ", " & _
                         "nTamano =" & chkCambia.Value & ",  " & _
                         "lAutorizaEnviosMozo =" & chkAutorizaEnviosMozo.Value & ",  " & _
                         "lActivo =" & chkActivo.Value & ", lReplica=1 " & _
                         " where tTAbla = 'MOZO' and tCodigo = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                  
                  If lGrabaHuella = True Then
                    Isql = "update ttabla set tvalor2='" & X & "' where ttabla='MOZO' and tcodigo='" & txtCodigo.Text & "'"
                    Cn.Execute Isql
                  End If
                  
                   nPos = frmMozo.RsCabecera.Bookmark
                   frmMozo.RsCabecera.Requery
                   If frmMozo.RsCabecera.RecordCount = 0 Then
                      frmMozo.RsCabecera.Filter = adFilterNone
                   End If
                   frmMozo.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
                If frmMozo.RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                'Cambia el MsgBox
                
                Dim mensj As String
                
'                If pais = "005" Then
'                    mensj = "Seguro de Eliminar al Camarero"
'                Else
'                    mensj = "Seguro de Eliminar al Mozo"
'                End If
                
                mensj = "Seguro de Eliminar al " & Mesero
                
                If MsgBox(mensj & " " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
               If Calcular("select count(tMozo) as Codigo From MPEDIDO where tMozo='" & txtCodigo.Text & "'", Cn) > 0 Then
'                 If pais = "005" Then
'                  MsgBox "Se tiene movimientos con este Camarero, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
'                  Exit Sub
'                 Else
'                  MsgBox "Se tiene movimientos con este Mozo, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
'                  Exit Sub
'                 End If

                  MsgBox "Se tiene movimientos con este " & Camarero & ", " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
                  
               End If
                
                'Cambia el Delete
                Cn.Execute "delete from TTABLA where tTabla = 'MOZO' and tCodigo = '" & txtCodigo & "'"
                frmMozo.RsCabecera.Requery
                If frmMozo.RsCabecera.RecordCount <> 0 Then
                   frmMozo.RsCabecera.MoveLast
                   Asignar
                   cmdTexto.Caption = "Registro " & IIf(frmMozo.RsCabecera.RecordCount = 0, 0, frmMozo.RsCabecera.AbsolutePosition) & " de " & frmMozo.RsCabecera.RecordCount
                Else
                   ActivarBotones False
                   Blanquear Me
                   Sw = True
                End If
                
          Case Is = 3 ' Salir
               Unload Me
               
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    
     If lHuellaDigitalPersona = True Then
        'HUELLA
        Set Capture = New DPFPCapture
        Capture.StartCapture
        Set CreateFtrs = New DPFPFeatureExtraction
        Set CreateTempl = New DPFPEnrollment
        Samples.Caption = CreateTempl.FeaturesNeeded
        Set ConvertSample = New DPFPSampleConversion
        '----------------------------------
            fraDigitalPersona.Visible = True
            fraSecugen.Visible = False
        
    End If
    If lHuellaSecugen = True Then
            fraDigitalPersona.Visible = False
            fraSecugen.Visible = True
            g_firstStep = False
            g_secondStep = False
            g_vefiryResult = False
          FpLibXCapture1(0).CodeName = 2
            FpLibXCapture1(1).CodeName = 2
            ' hardware initialize
            ' FDP02(0), FDU02(1), FDU03(2), FDU04(3)
            Dim DefaultDevice As Integer
            DefaultDevice = 2   ' FDU03
'            For i = 0 To 4
'             If (Option1(i).value = True) Then
'                 DefaultDevice = i
'                 Exit For
'             End If
'            Next i
            FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ANSI378        ' ANSI 378
            FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
            FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
            '  FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
            ReDim g_FirstMinData(FpLibXCapture1(0).MinutiaeSize) As Byte
            ReDim g_SecondMinData(FpLibXCapture1(1).MinutiaeSize) As Byte
            'ReDim g_MatchingMinData(FpLibXCapture1(2).MinutiaeSize) As Byte

           ' FpLibXCapture1(2).CodeName = 2
            
 
            
            FpLibXCapture1(0).Refresh
            FpLibXCapture1(1).Refresh
           ' FpLibXCapture1(2).Refresh
    End If
    'Ingrese el SubTitulo
'    If pais = "005" Then
'        Me.Caption = " Mantenimiento de Camareros "
'    Else
'        Me.Caption = " Mantenimiento de Mozos "
'    End If
    Me.Caption = "Mantenimiento de " & Mesero
    'fraDetalle.Caption = Me.Caption
    If lInfhotel Then
       cmdOpcion(0).Enabled = False
       cmdOpcion(2).Enabled = False
    End If
        
    'Botones
    Isql = "select tCodigo, nBoton from TTABLA where nBoton > 0 and TTABLA = 'MOZO' and lActivo=1 order by tCodigo"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
       Botonera
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmMozo.RsCabecera.RecordCount = 0, 0, frmMozo.RsCabecera.AbsolutePosition) & " de " & frmMozo.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmMozoDetalle = Nothing
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

Private Sub cmdNA_Click()
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
      txtBoton.Caption = "NA"
   End If
End Sub

Private Sub Botonera()
    Dim i As Integer
    txtBoton.Caption = "NA"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 19
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).backColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tCodigo = txtCodigo.Text Then
                  txtBoton.Caption = str(i)
                  cmdBoton(i).backColor = vbRed
               Else
                  cmdBoton(i).backColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 19
           cmdBoton(i).backColor = vbButtonFace
           cmdBoton(i).Enabled = True
       Next i
    End If
End Sub

Private Sub txtDetallado_LostFocus()
    Call ValidaStr(txtDetallado)
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub





'HUELLA
Private Sub ReportStatus(ByVal str As String)
 ' Add string to list box.
 Status.AddItem (str)
 ' Move list box selection down.
 Status.ListIndex = Status.NewIndex
End Sub

Private Sub Capture_OnReaderConnect(ByVal ReaderSerNum As String)
 ReportStatus ("Lector de Huella Dactilar Conectado")
End Sub

Private Sub Capture_OnReaderDisconnect(ByVal ReaderSerNum As String)
 ReportStatus ("Lector de Huella Dactilar Desconectado")
End Sub

Private Sub Capture_OnFingerTouch(ByVal ReaderSerNum As String)
    lblHuella.Visible = False
 ReportStatus ("Lector de Huella Dactilar presionado")
End Sub
Private Sub Capture_OnFingerGone(ByVal ReaderSerNum As String)
 ReportStatus ("El dedo fue retirado del Lector de Huella Dactilar")
End Sub
Private Sub Capture_OnSampleQuality(ByVal ReaderSerNum As String, ByVal Feedback As DPFPCaptureFeedbackEnum)
 If Feedback = CaptureFeedbackGood Then
  ReportStatus ("La calidad de la huella obtenida es buena")
  Else
  ReportStatus ("La calidad de la huella obtenida es no es adecuada")
  End If
End Sub

Private Sub DrawPicture(ByVal Pict As IPictureDisp)
 ' Must use hidden PictureBox to easily resize picture.
 Set HiddenPict.Picture = Pict
 Picture1.PaintPicture HiddenPict.Picture, _
       0, 0, Picture1.ScaleWidth, _
       Picture1.ScaleHeight, _
       0, 0, HiddenPict.ScaleWidth, _
       HiddenPict.ScaleHeight, vbSrcCopy
 Picture1.Picture = Picture1.Image
End Sub
Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
 Dim Feedback As DPFPCaptureFeedbackEnum
 'lGrabaHuella = False
 ReportStatus ("Huella dactilar capturada")
 ' Draw fingerprint image.
 DrawPicture ConvertSample.ConvertToPicture(Sample)
 ' Process sample and create feature set for purpose of enrollment.
 Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeEnrollment)
 ' Quality of sample is not good enough to produce feature set.
 If Feedback = CaptureFeedbackGood Then
  ReportStatus ("Patrón de Huella Dactilar creado.")
  Prompt.Caption = "Presione el lector de Huella Dactilar con el mismo dedo"
  ' Add feature set to template.
  CreateTempl.AddFeatures CreateFtrs.FeatureSet
  ' Show number of samples needed to complete template.
  Samples.Caption = CreateTempl.FeaturesNeeded
  ' Check if template has been created.
  If CreateTempl.TemplateStatus = TemplateStatusTemplateReady Then
  '  MainFrame.SetTemplete CreateTempl.Template
    ' Template has been created, so stop capturing samples.
    Capture.StopCapture
    Prompt.Caption = "Puede guardar el registro."
    lGrabaHuella = True
       Me.SetTemplete CreateTempl.Template
   ' MsgBox "The fingerprint template was created."
  End If
 End If
 End Sub


Private Function arraytohex(arr() As Byte) As String

Dim templatestr As String
Dim tempstr As String
Dim i As Integer

templatestr = ""

For i = LBound(arr) To UBound(arr)

tempstr = Hex$(arr(i))

If Len(tempstr) = 1 Then tempstr = "0" + tempstr 'padhex
templatestr = templatestr + tempstr
Next i

arraytohex = templatestr

End Function
 
Public Sub hextoarray(inphex As String, outarray() As Byte)

ReDim outarray(0 To Len(inphex) / 2)

Dim i As Integer
For i = 1 To Len(inphex) Step 2

outarray(((i + 1) / 2) - 1) = val("&H" + Mid$(inphex, i, 2))
Next i

End Sub
Public Sub SetTemplete(ByVal Template As Object)
 Set Templ = Template
End Sub

Public Sub limpiaVariables()
Picture1.Picture = Nothing
lblHuella.Visible = False
Prompt.Caption = "Toque el lector de Huella Dactilar"
Samples.Caption = ""
Status.Clear

' Create capture operation.
 Set Capture = New DPFPCapture
 Capture.StartCapture
 Set CreateFtrs = New DPFPFeatureExtraction
 Set CreateTempl = New DPFPEnrollment
 Samples.Caption = CreateTempl.FeaturesNeeded
 Set ConvertSample = New DPFPSampleConversion

End Sub

Private Sub TemplateFormat_Click(Index As Integer)

    g_firstStep = False
    g_secondStep = False
    g_vefiryResult = False

    ' Set format of template to use.
    ' two modes of FpLibXVerify1 and FpLibXCapture1 MUST be the same minutiae mode
    ' otherwise comment out to use MinutiaeMode property
    Select Case Index
    Case 0
        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_SG400         ' By default
        FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_SG400
        FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_SG400
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_SG400
        
    Case 1
        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ANSI378        ' ANSI 378
        FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        
    Case 2
        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ISO19794      ' ISO 19794-2
        FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_ISO19794
        FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_ISO19794
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ISO19794
    End Select

    'Redim these variables after setting the MinutiaeMode,
    'because size of data is different for each MinutiaeMode
    ReDim g_FirstMinData(FpLibXCapture1(0).MinutiaeSize) As Byte
    ReDim g_SecondMinData(FpLibXCapture1(1).MinutiaeSize) As Byte
    ReDim g_MatchingMinData(FpLibXCapture1(2).MinutiaeSize) As Byte


End Sub

Private Sub FpLibXCapture1_Paint(Index As Integer, ByVal hwnd As stdole.OLE_HANDLE, ByVal hdc As stdole.OLE_HANDLE)
    Select Case Index
    Case 0
        If g_firstStep = False Then
            TextOut hdc, 30, 70, "Click Aquí..", 12
        End If
        
    Case 1
        If g_firstStep = True And g_secondStep = False Then
            TextOut hdc, 30, 70, "Click Aquí..", 12
        End If
        
    Case 2
        If g_firstStep = True And g_secondStep = True And g_vefiryResult = False Then
            TextOut hdc, 30, 70, "Click Aquí..", 12
        End If
    
    End Select
End Sub

Private Sub FpLibXCapture1_Click(Index As Integer)

Dim register_quality, verify_quality As Integer
Dim Timeout As Integer

register_quality = 50
verify_quality = 30
Timeout = 6000 ' 6seconds

    Select Case (Index)
    Case 0
        Result = FpLibXCapture1(Index).LiveCapture(Timeout, register_quality)
        FpLibXCapture1(Index).GetMinutiaeData (g_FirstMinData)
        
        g_firstStep = True
        g_secondStep = False
        g_vefiryResult = False

'    Dim x As String
'     x = arraytohex(g_FirstMinData)
'    Debug.Print x
'
'    'Dim Cn As Connection
'
'    Set Cn = New Connection
'    Cn.Provider = "SQLOLEDB"
'
'    Cn.CursorLocation = adUseServer
'    Cn.ConnectionString = "User ID=sa" & _
'                          ";password=sistemas" & _
'                          ";Data Source=INFLIMDES02" & _
'                          ";Initial Catalog=Prueba_xxx"
'
'    Cn.CommandTimeout = 300
'    Cn.Open
'    Cn.Execute "UPDATE thuella set huella='" & x & "' where codigo='00001'"
        
        FpLibXCapture1(1).Clear
       ' FpLibXCapture1(2).Clear
        LabelMessage.Caption = ""
    
    Case 1
        If g_firstStep = True Then
            Result = FpLibXCapture1(Index).LiveCapture(Timeout, register_quality)
            FpLibXCapture1(Index).GetMinutiaeData (g_SecondMinData)
            
            'you can get the score
            score = FpLibXVerify1.GetMatchingScore(g_FirstMinData, g_SecondMinData)
            
            Result = FpLibXVerify1.Register(g_FirstMinData, g_SecondMinData)
            If Result = True Then
                g_secondStep = True
                LabelMessage.Caption = "Patrón de Huellas Correcto. Puede Guardar el Registro"
                 lGrabaHuella = True

                g_vefiryResult = False
              '  FpLibXCapture1(2).Clear
            Else
                g_secondStep = False
             '   FpLibXCapture1(2).Clear
                If FpLibXVerify1.ErrorCode = ERROR_VERIFY_FAKE Then
                    LabelMessage.Caption = "Retire el dedo y tocar de nuevo , por favor"
                     lGrabaHuella = False

                Else
                    LabelMessage.Caption = "Patrón de Huellas Incorrecto, Intente de Nuevo"
                     lGrabaHuella = False

                End If
            End If
            
        End If
        
    Case 2
        If g_secondStep = True Then
            bResult = FpLibXCapture1(Index).LiveCapture(Timeout, verify_quality)
            FpLibXCapture1(Index).GetMinutiaeData (g_MatchingMinData)
            
                Dim huella() As Byte
          
                   Set Cn = New Connection
                   Cn.Provider = "SQLOLEDB"
                   
                   Cn.CursorLocation = adUseServer
                   Cn.ConnectionString = "User ID=sa" & _
                                         ";password=sistemas" & _
                                         ";Data Source=INFLIMDES02" & _
                                         ";Initial Catalog=Prueba_xxx"
                                         
                   Cn.CommandTimeout = 300
                   Cn.Open
                   Dim rst1 As ADODB.Recordset
                   Set rst1 = New ADODB.Recordset
                   Dim p As String
                  Set rst1 = Cn.Execute("select huella from thuella where codigo='00001'")
                   p = rst1.Fields(0)
                    hextoarray p, huella()
                    
            
            
            
            
            
            Result = FpLibXVerify1.Verify(huella, g_MatchingMinData)
            
            
            If Result = True Then
                g_vefiryResult = True
                LabelMessage.Caption = "Identical"
            Else
                g_vefiryResult = False
                If FpLibXVerify1.ErrorCode = ERROR_VERIFY_FAKE Then
                    LabelMessage.Caption = "Release your finger and touch on again, please"
                Else
                    LabelMessage.Caption = "NOT Identical"
                End If
                
            End If
        End If
        
    End Select
 
    FpLibXCapture1(0).Refresh
    FpLibXCapture1(1).Refresh
 '   FpLibXCapture1(2).Refresh
End Sub


