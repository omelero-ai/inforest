VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{B52C1CDE-38E9-11D5-98EC-00C0F01C6C81}#1.0#0"; "ifepson.ocx"
Begin VB.Form frmNotaCreditoDetalle 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7485
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11040
   Icon            =   "frmNotaCreditoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNotaCreditoDetalle.frx":030A
   ScaleHeight     =   7485
   ScaleWidth      =   11040
   Begin EPSON_Impresora_Fiscal.PrinterFiscal PrinterFiscalArgentina 
      Left            =   11160
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      PortNumber      =   2
   End
   Begin VB.Frame FrameFeSpring 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   2760
      TabIndex        =   62
      Top             =   840
      Visible         =   0   'False
      Width           =   6315
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":040C
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo hash almacenado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1185
         TabIndex        =   65
         Top             =   1155
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a spring."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1170
         TabIndex        =   64
         Top             =   870
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Facturación Electronica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   15
         TabIndex        =   63
         Top             =   15
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":061F
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":0832
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmNotaCreditoDetalle.frx":0B74
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a Spring......."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1110
         Left            =   210
         TabIndex        =   66
         Top             =   435
         Width           =   5910
      End
   End
   Begin VB.Frame fraDetalle 
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6795
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   14835
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         Caption         =   "Motivo Administrativo"
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
         Height          =   615
         Index           =   16
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1700
         Width           =   1600
      End
      Begin VB.TextBox txtMotivoAdm 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   71
         Top             =   1380
         Width           =   9075
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53870593
         CurrentDate     =   38096
      End
      Begin VB.TextBox txtNC3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtNC2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
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
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   300
         Width           =   870
      End
      Begin VB.CommandButton cmdNotaCredito 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1845
      End
      Begin VB.CommandButton cmdNotaCredito 
         BackColor       =   &H00F2FAED&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txtPrefijo 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1740
         MaxLength       =   17
         TabIndex        =   58
         Top             =   700
         Width           =   255
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         Caption         =   "Correlativo"
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
         Index           =   6
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1700
         Width           =   1600
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         Caption         =   "Tipo de Documento"
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1700
         Width           =   1600
      End
      Begin VB.TextBox txtCorrela 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2940
         MaxLength       =   17
         TabIndex        =   57
         Top             =   700
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         Caption         =   "Motivo"
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
         Index           =   8
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1700
         Width           =   1600
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         Caption         =   "Monto"
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
         Index           =   9
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1700
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00F0F0F0&
         Caption         =   " Resultados "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   11280
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   7875
         Begin VB.TextBox txtNotaLey1__ 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   69
            Text            =   "0.00"
            Top             =   1850
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtDocLey1__ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   68
            Text            =   "0.00"
            Top             =   1850
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   67
            Text            =   "0.00"
            Top             =   1850
            Visible         =   0   'False
            Width           =   1395
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
            Height          =   345
            Index           =   15
            Left            =   3270
            Picture         =   "frmNotaCreditoDetalle.frx":0EB6
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1455
            Visible         =   0   'False
            Width           =   390
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
            Height          =   345
            Index           =   14
            Left            =   3270
            Picture         =   "frmNotaCreditoDetalle.frx":0FB8
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1110
            Visible         =   0   'False
            Width           =   390
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
            Height          =   345
            Index           =   13
            Left            =   3270
            Picture         =   "frmNotaCreditoDetalle.frx":10BA
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   765
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.TextBox txtResTotal 
            Alignment       =   1  'Right Justify
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   2190
            Width           =   1395
         End
         Begin VB.TextBox txtResImp3 
            Alignment       =   1  'Right Justify
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   1485
            Width           =   1395
         End
         Begin VB.TextBox txtResImp2 
            Alignment       =   1  'Right Justify
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtResImp1 
            Alignment       =   1  'Right Justify
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   50
            Text            =   "0.00"
            Top             =   795
            Width           =   1395
         End
         Begin VB.TextBox txtResNeto 
            Alignment       =   1  'Right Justify
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
            Left            =   6060
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   49
            Text            =   "0.00"
            Top             =   450
            Width           =   1395
         End
         Begin VB.TextBox txtDocTotal__ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   48
            Text            =   "0.00"
            Top             =   2190
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp3__ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   47
            Text            =   "0.00"
            Top             =   1485
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp2__ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   46
            Text            =   "0.00"
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp1__ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   795
            Width           =   1395
         End
         Begin VB.TextBox txtDocNeto__ 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3930
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   450
            Width           =   1395
         End
         Begin VB.TextBox txtNCNeto__ 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   38
            Text            =   "0.00"
            Top             =   480
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp1__ 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   795
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp2__ 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   36
            Text            =   "0.00"
            Top             =   1140
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp3__ 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   35
            Text            =   "0.00"
            Top             =   1485
            Width           =   1395
         End
         Begin VB.TextBox txtNCTotal__ 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   34
            Text            =   "0.00"
            Top             =   2190
            Width           =   1395
         End
         Begin VB.Label lblLey1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Ley1 :"
            Height          =   195
            Index           =   12
            Left            =   495
            TabIndex        =   70
            Top             =   1845
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Resultado"
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
            Index           =   11
            Left            =   6315
            TabIndex        =   56
            Top             =   210
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Documento"
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
            Index           =   10
            Left            =   4125
            TabIndex        =   55
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Nota de Crédito"
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
            Index           =   9
            Left            =   1815
            TabIndex        =   54
            Top             =   210
            Width           =   1350
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Total :"
            Height          =   195
            Index           =   8
            Left            =   990
            TabIndex        =   43
            Top             =   2235
            Width           =   450
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 3 :"
            Height          =   195
            Index           =   7
            Left            =   555
            TabIndex        =   42
            Top             =   1530
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Index           =   6
            Left            =   555
            TabIndex        =   41
            Top             =   1185
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Index           =   5
            Left            =   570
            TabIndex        =   40
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Neto :"
            Height          =   195
            Index           =   4
            Left            =   990
            TabIndex        =   39
            Top             =   495
            Width           =   435
         End
      End
      Begin VB.TextBox txtSerie 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2025
         MaxLength       =   17
         TabIndex        =   32
         Top             =   700
         Width           =   870
      End
      Begin VB.CommandButton cmdOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
         Caption         =   "Número de Serie"
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
         Index           =   5
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1700
         Width           =   1600
      End
      Begin VB.TextBox txtObservacion 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1050
         Width           =   9075
      End
      Begin VB.TextBox txtNC1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FAED&
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   255
      End
      Begin VB.Frame frmDetalleNC 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Detalle"
         Height          =   4095
         Left            =   120
         TabIndex        =   75
         Top             =   2400
         Width           =   10815
         Begin VB.Frame frmIngresaItem 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Height          =   3780
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Visible         =   0   'False
            Width           =   10600
            Begin VB.TextBox txtItemPuDoc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   480
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   119
               Text            =   "0.000"
               Top             =   720
               Width           =   1995
            End
            Begin VB.TextBox txtItemCanDoc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   480
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   118
               Text            =   "0.000000"
               Top             =   1320
               Width           =   1995
            End
            Begin VB.TextBox txtItemTotDoc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   480
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   117
               Text            =   "0.000"
               Top             =   1920
               Width           =   1995
            End
            Begin VB.CommandButton cmdParcial 
               BackColor       =   &H00F2FAED&
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
               Index           =   6
               Left            =   5040
               Picture         =   "frmNotaCreditoDetalle.frx":11BC
               Style           =   1  'Graphical
               TabIndex        =   113
               Top             =   1560
               Width           =   1170
            End
            Begin VB.CommandButton cmdParcial 
               BackColor       =   &H00F2FAED&
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
               Index           =   5
               Left            =   5040
               Picture         =   "frmNotaCreditoDetalle.frx":12BE
               Style           =   1  'Graphical
               TabIndex        =   112
               Top             =   840
               Width           =   1170
            End
            Begin VB.TextBox txtItemPuNC 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   285
               Left            =   2520
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   111
               Text            =   "0.000"
               Top             =   720
               Width           =   1875
            End
            Begin VB.CommandButton cmdParcial 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   8
               Left            =   4440
               Picture         =   "frmNotaCreditoDetalle.frx":17F0
               Style           =   1  'Graphical
               TabIndex        =   110
               Top             =   1920
               Width           =   390
            End
            Begin VB.CommandButton cmdParcial 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   7
               Left            =   4440
               Picture         =   "frmNotaCreditoDetalle.frx":18F2
               Style           =   1  'Graphical
               TabIndex        =   109
               Top             =   1320
               Width           =   390
            End
            Begin VB.TextBox txtItemTotNC 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   285
               Left            =   2520
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   108
               Text            =   "0.000"
               Top             =   1920
               Width           =   1875
            End
            Begin VB.TextBox txtItemCanNC 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   285
               Left            =   2520
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   107
               Text            =   "0.000000"
               Top             =   1320
               Width           =   1875
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F0F0&
               Caption         =   "Nota de credito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   27
               Left            =   3000
               TabIndex        =   121
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F0F0&
               Caption         =   "Documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   26
               Left            =   600
               TabIndex        =   120
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F0F0&
               Caption         =   "Total :"
               Height          =   195
               Index           =   25
               Left            =   480
               TabIndex        =   116
               Top             =   1680
               Width           =   450
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F0F0&
               Caption         =   "Cantidad :"
               Height          =   195
               Index           =   24
               Left            =   480
               TabIndex        =   115
               Top             =   1080
               Width           =   720
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F0F0&
               Caption         =   "Precio Unitario :"
               Height          =   195
               Index           =   23
               Left            =   480
               TabIndex        =   114
               Top             =   480
               Width           =   1125
            End
         End
         Begin VB.TextBox txtNotaLey1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   105
            Text            =   "0.00"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtNCTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   103
            Text            =   "0.00"
            Top             =   3120
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   98
            Text            =   "0.00"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   96
            Text            =   "0.00"
            Top             =   3120
            Width           =   1395
         End
         Begin VB.TextBox txtNCImp1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   95
            Text            =   "0.00"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtNCNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   94
            Text            =   "0.00"
            Top             =   3120
            Width           =   1395
         End
         Begin VB.TextBox txtDocLey1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   93
            Text            =   "0.00"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtDocTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   91
            Text            =   "0.00"
            Top             =   3120
            Width           =   1395
         End
         Begin VB.CommandButton cmdParcial 
            BackColor       =   &H00F2FAED&
            Enabled         =   0   'False
            Height          =   550
            Index           =   4
            Left            =   5130
            Picture         =   "frmNotaCreditoDetalle.frx":19F4
            Style           =   1  'Graphical
            TabIndex        =   86
            TabStop         =   0   'False
            ToolTipText     =   "Quitar todos"
            Top             =   2280
            Width           =   600
         End
         Begin VB.CommandButton cmdParcial 
            BackColor       =   &H00F2FAED&
            Enabled         =   0   'False
            Height          =   550
            Index           =   3
            Left            =   5130
            Picture         =   "frmNotaCreditoDetalle.frx":1F36
            Style           =   1  'Graphical
            TabIndex        =   85
            TabStop         =   0   'False
            ToolTipText     =   "Quitar uno"
            Top             =   1740
            Width           =   600
         End
         Begin VB.CommandButton cmdParcial 
            BackColor       =   &H00F2FAED&
            Enabled         =   0   'False
            Height          =   550
            Index           =   2
            Left            =   5130
            Picture         =   "frmNotaCreditoDetalle.frx":2478
            Style           =   1  'Graphical
            TabIndex        =   84
            TabStop         =   0   'False
            ToolTipText     =   "Pasar todos"
            Top             =   1035
            Width           =   600
         End
         Begin VB.CommandButton cmdParcial 
            BackColor       =   &H00F2FAED&
            Enabled         =   0   'False
            Height          =   550
            Index           =   1
            Left            =   5130
            Picture         =   "frmNotaCreditoDetalle.frx":29BA
            Style           =   1  'Graphical
            TabIndex        =   83
            TabStop         =   0   'False
            ToolTipText     =   "Pasar uno"
            Top             =   480
            Width           =   600
         End
         Begin VB.TextBox txtDocImp3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   82
            Text            =   "0.00"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   80
            Text            =   "0.00"
            Top             =   3120
            Width           =   1395
         End
         Begin VB.TextBox txtDocImp1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   79
            Text            =   "0.00"
            Top             =   3720
            Width           =   1395
         End
         Begin VB.TextBox txtDocNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   78
            Text            =   "0.00"
            Top             =   3120
            Width           =   1395
         End
         Begin TrueOleDBGrid80.TDBGrid grdDocumento 
            Height          =   2655
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   4683
            _LayoutType     =   4
            _RowHeight      =   26
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerColor=   16777215
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   7
            Caption         =   "Detalle Documento"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
            RowDividerColor =   15263202
            RowSubDividerColor=   16777215
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.locked=0,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
            _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Arial"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
            _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=2,.fontname=Arial"
            _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&,.bold=0"
            _StyleDefs(19)  =   ":id=6,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(20)  =   ":id=6,.fontname=Arial"
            _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(27)  =   "Splits(0).Style:id=87,.parent=1"
            _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
            _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=89,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=91,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8,.namedParent=38"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
            _StyleDefs(43)  =   "Named:id=33:Normal"
            _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
            _StyleDefs(51)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(52)  =   ":id=34,.fontname=Arial"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
            _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=900,.italic=0,.underline=0"
            _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=36,.fontname=Arial"
            _StyleDefs(60)  =   "Named:id=37:Caption"
            _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(62)  =   "Named:id=38:HighlightRow"
            _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.appearance=0"
            _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=-1,.fontsize=675,.italic=0,.underline=0"
            _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
            _StyleDefs(66)  =   ":id=38,.fontname=Small Fonts"
            _StyleDefs(67)  =   "Named:id=39:EvenRow"
            _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(69)  =   "Named:id=40:OddRow"
            _StyleDefs(70)  =   ":id=40,.parent=33"
            _StyleDefs(71)  =   "Named:id=41:RecordSelector"
            _StyleDefs(72)  =   ":id=41,.parent=34"
            _StyleDefs(73)  =   "Named:id=42:FilterBar"
            _StyleDefs(74)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid80.TDBGrid grdNotaCredito 
            Height          =   2655
            Left            =   5760
            TabIndex        =   77
            Top             =   240
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   4683
            _LayoutType     =   4
            _RowHeight      =   26
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0).ScrollBars=   2
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerColor=   16777215
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=20"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   7
            Caption         =   "Detalle nota de credito"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
            RowDividerColor =   15263202
            RowSubDividerColor=   16777215
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.locked=0,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
            _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Arial"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H70655D&,.fgcolor=&H80000005&"
            _StyleDefs(13)  =   ":id=2,.borderSize=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=2,.fontname=Arial"
            _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&HF2FAED&,.fgcolor=&H232323&,.bold=0"
            _StyleDefs(19)  =   ":id=6,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(20)  =   ":id=6,.fontname=Arial"
            _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.bgcolor=&H808000&"
            _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1"
            _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1"
            _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(27)  =   "Splits(0).Style:id=87,.parent=1"
            _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
            _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=89,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=91,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8,.namedParent=38"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=106,.parent=87"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
            _StyleDefs(43)  =   "Named:id=33:Normal"
            _StyleDefs(44)  =   ":id=33,.parent=0,.valignment=2,.bgcolor=&H80000018&,.locked=-1,.appearance=0"
            _StyleDefs(45)  =   ":id=33,.borderSize=1,.borderColor=&H80000005&,.borderType=0,.bold=0"
            _StyleDefs(46)  =   ":id=33,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(47)  =   ":id=33,.fontname=Small Fonts"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
            _StyleDefs(50)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
            _StyleDefs(51)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(52)  =   ":id=34,.fontname=Arial"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H808080&,.locked=-1,.appearance=0"
            _StyleDefs(57)  =   ":id=36,.borderColor=&H808000&,.bold=0,.fontsize=900,.italic=0,.underline=0"
            _StyleDefs(58)  =   ":id=36,.strikethrough=0,.charset=0"
            _StyleDefs(59)  =   ":id=36,.fontname=Arial"
            _StyleDefs(60)  =   "Named:id=37:Caption"
            _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(62)  =   "Named:id=38:HighlightRow"
            _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H80000012&,.appearance=0"
            _StyleDefs(64)  =   ":id=38,.borderSize=1,.bold=-1,.fontsize=675,.italic=0,.underline=0"
            _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
            _StyleDefs(66)  =   ":id=38,.fontname=Small Fonts"
            _StyleDefs(67)  =   "Named:id=39:EvenRow"
            _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(69)  =   "Named:id=40:OddRow"
            _StyleDefs(70)  =   ":id=40,.parent=33"
            _StyleDefs(71)  =   "Named:id=41:RecordSelector"
            _StyleDefs(72)  =   ":id=41,.parent=34"
            _StyleDefs(73)  =   "Named:id=42:FilterBar"
            _StyleDefs(74)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lblLey1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Ley1 :"
            Height          =   195
            Index           =   1
            Left            =   9120
            TabIndex        =   104
            Top             =   3480
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Total :"
            Height          =   195
            Index           =   22
            Left            =   9120
            TabIndex        =   102
            Top             =   2880
            Width           =   450
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 3 :"
            Height          =   195
            Index           =   21
            Left            =   7440
            TabIndex        =   101
            Top             =   3480
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Index           =   20
            Left            =   7440
            TabIndex        =   100
            Top             =   2880
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Index           =   19
            Left            =   5760
            TabIndex        =   99
            Top             =   3480
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Neto :"
            Height          =   195
            Index           =   18
            Left            =   5760
            TabIndex        =   97
            Top             =   2880
            Width           =   435
         End
         Begin VB.Label lblLey1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Ley1 :"
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   92
            Top             =   3480
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Total :"
            Height          =   195
            Index           =   17
            Left            =   3480
            TabIndex        =   90
            Top             =   2880
            Width           =   450
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 3 :"
            Height          =   195
            Index           =   16
            Left            =   1800
            TabIndex        =   89
            Top             =   3480
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Index           =   15
            Left            =   1800
            TabIndex        =   88
            Top             =   2880
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   87
            Top             =   3480
            Width           =   870
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00F0F0F0&
            Caption         =   "Neto :"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   81
            Top             =   2880
            Width           =   435
         End
      End
      Begin VB.Label lblEstadoFacturacion 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   6480
         Width           =   10215
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Motivo Administrativo:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   72
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Emitido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   10080
         TabIndex        =   61
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Motivo :"
         Height          =   195
         Index           =   2
         Left            =   1050
         TabIndex        =   31
         Top             =   1095
         Width           =   570
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Documento :"
         Height          =   195
         Index           =   3
         Left            =   705
         TabIndex        =   29
         Top             =   700
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   28
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         Caption         =   "Nota de Crédito:"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   27
         Top             =   345
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F0F0F0&
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10980
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6735
      Width           =   11040
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Procesar"
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
         Index           =   12
         Left            =   7440
         Picture         =   "frmNotaCreditoDetalle.frx":2EFC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
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
         Left            =   9780
         Picture         =   "frmNotaCreditoDetalle.frx":323E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
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
         Left            =   8610
         Picture         =   "frmNotaCreditoDetalle.frx":3330
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
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
         Left            =   5100
         Picture         =   "frmNotaCreditoDetalle.frx":3432
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
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
         Left            =   6270
         Picture         =   "frmNotaCreditoDetalle.frx":3964
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H00F0F0F0&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4950
         TabIndex        =   18
         Top             =   30
         Width           =   5010
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00F2FAED&
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmNotaCreditoDetalle.frx":3E96
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00F2FAED&
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmNotaCreditoDetalle.frx":43D8
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00F2FAED&
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmNotaCreditoDetalle.frx":491A
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00F2FAED&
            Height          =   555
            Index           =   5
            Left            =   4470
            Picture         =   "frmNotaCreditoDetalle.frx":4E5C
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00F2FAED&
            Height          =   555
            Index           =   4
            Left            =   3990
            Picture         =   "frmNotaCreditoDetalle.frx":539E
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00F2FAED&
            Height          =   555
            Index           =   3
            Left            =   3510
            Picture         =   "frmNotaCreditoDetalle.frx":58E0
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            BackColor       =   &H00F0F0F0&
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
            Left            =   1470
            TabIndex        =   25
            Top             =   150
            Width           =   1935
         End
      End
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageCab 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imageHash 
      Height          =   615
      Left            =   11040
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmNotaCreditoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Option Explicit

Dim RsDetalle As Recordset
Dim RsDocumento As Recordset
Dim RsTipoDocumento As Recordset

Dim sDetalle As String
Dim sDocumento As String
Dim sNC As String

Dim nNeto As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nTotal As Double

Dim nNCNeto As Double
Dim nNCImp1 As Double
Dim nNCImp2 As Double
Dim nNCImp3 As Double
Dim nNCTotal As Double
Dim nCantidad As Double

'FACTURACION E
Dim RsNotaCredito As Recordset
Dim nEmision As Integer
Dim xTipoNotaCredito As String
Dim xCodigoMotivo As String
Dim cadenaCodigoHash As String
'-----
Dim sImp As String
Dim impTipo As String
Dim iImagenCab As Boolean
Dim RsImpDocumentoE As New Recordset
Dim xMontoTexto As String

Dim lNotaCreditoEmitirFE As Boolean
Dim tPrefijoEnlace As String

Dim TimpresionDolaresDelivery As Boolean
Dim RsDoc As New Recordset
Dim RsDocNC As New Recordset
Sub Asignar()
    With frmNotaCredito.RsCabecera
        'Cuadro de Texto
        sNC = IIf(IsNull(!tNotaCredito), "", !tNotaCredito)
        txtNC1.Text = Mid(sNC, 1, 1)

        dtpFecha = IIf(IsNull(!fFecha), "", !fFecha)
        sDocumento = IIf(IsNull(!tDocumento), "", !tDocumento)
        txtPrefijo = Mid(sDocumento, 1, 1)
        
        If pais = "002" Then
            txtNC2.Text = Mid(sNC, 2, 6)
            txtNC3.Text = Mid(sNC, 8, 9)
            txtSerie = Mid(sDocumento, 2, 6)
            txtCorrela = Mid(sDocumento, 8, 9)
        Else
            txtNC2.Text = Mid(sNC, 2, 5)
            txtNC3.Text = Mid(sNC, 7, 9)
            txtSerie = Mid(sDocumento, 2, 5)
            txtCorrela = Mid(sDocumento, 7, 9)
        End If

        txtMotivoAdm.Text = IIf(IsNull(!MotivoAdm), "", !MotivoAdm)
        
        txtObservacion = IIf(IsNull(!tObservacion), "", !tObservacion)
        lblEstado.Caption = IIf(IsNull(!EstadoDocumento), "", !EstadoDocumento)
        
        nNCNeto = IIf(IsNull(!nNeto), 0, !nNeto)
        nNCImp1 = IIf(IsNull(!nImpuesto1), 0, !nImpuesto1)
        nNCImp2 = IIf(IsNull(!nImpuesto2), 0, !nImpuesto2)
        nNCImp3 = IIf(IsNull(!nImpuesto3), 0, !nImpuesto3)
        nNCTotal = IIf(IsNull(!nVenta), 0, !nVenta)
        
        nNeto = IIf(IsNull(!nDocNeto), 0, !nDocNeto)
        nImpuesto1 = IIf(IsNull(!nDocImpuesto1), 0, !nDocImpuesto1)
        nImpuesto2 = IIf(IsNull(!nDocImpuesto2), 0, !nDocImpuesto2)
        nImpuesto3 = IIf(IsNull(!nDocImpuesto3), 0, !nDocImpuesto3)
        nTotal = IIf(IsNull(!nDocVenta), 0, !nDocVenta)
                        
        txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
        txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
        txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
        txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
        txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")

        txtDocNeto.Text = Format(nNeto, "#,###,##0.00")
        txtDocImp1.Text = Format(nImpuesto1, "#,###,##0.00")
        txtDocImp2.Text = Format(nImpuesto2, "#,###,##0.00")
        txtDocImp3.Text = Format(nImpuesto3, "#,###,##0.00")
        txtDocTotal.Text = Format(nTotal, "#,###,##0.00")
        
        txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
        txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
        txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
        txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
        txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
        Me.txtDocLey1.Text = Format(Calcular("select isnull(nley1,0) as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), "#,###,##0.00")
        
        If Round(nNCTotal, 2) = Round(nTotal, 2) Then
            Me.txtNotaLey1.Text = Me.txtDocLey1.Text
        End If
        
         xCodigoMotivo = !tMotivo
        
        cmdOpcion(4).Enabled = False
        cmdOpcion(5).Enabled = False
        cmdOpcion(6).Enabled = False
        cmdOpcion(8).Enabled = True
        cmdOpcion(9).Enabled = True
        cmdOpcion(16).Enabled = True
        cmdNotaCredito(1).Enabled = False
        cmdNotaCredito(2).Enabled = False
        
        If lFacturacionE Then
            Me.lblEstadoFacturacion.Caption = IIf(IsNull(!RespuestaFacturacion), "", !RespuestaFacturacion)
        Else
            Me.lblEstadoFacturacion.Caption = ""
        End If
        
    End With
    
    If lblEstado.Caption = "ANULADO" Then
       cmdOpcion(1).Enabled = False
       cmdOpcion(2).Enabled = False
       cmdOpcion(12).Enabled = False
       cmdOpcion(8).Enabled = False
       cmdOpcion(9).Enabled = False
       cmdOpcion(16).Enabled = False
       dtpFecha.Enabled = False
       
        cmdParcial(1).Enabled = False
        cmdParcial(2).Enabled = False
        cmdParcial(3).Enabled = False
        cmdParcial(4).Enabled = False
       
    ElseIf lblEstado.Caption = "PROCESADO" Then
       cmdOpcion(1).Enabled = False
       cmdOpcion(12).Enabled = False
       cmdOpcion(8).Enabled = False
       cmdOpcion(9).Enabled = False
       cmdOpcion(16).Enabled = False
       dtpFecha.Enabled = False
       
        cmdParcial(1).Enabled = False
        cmdParcial(2).Enabled = False
        cmdParcial(3).Enabled = False
        cmdParcial(4).Enabled = False
       
    ElseIf lblEstado.Caption = "PAGADO" Then
       cmdOpcion(1).Enabled = False
       cmdOpcion(12).Enabled = False
       cmdOpcion(8).Enabled = False
       cmdOpcion(9).Enabled = False
       cmdOpcion(16).Enabled = False
       dtpFecha.Enabled = False
       
        cmdParcial(1).Enabled = False
        cmdParcial(2).Enabled = False
        cmdParcial(3).Enabled = False
        cmdParcial(4).Enabled = False
       
    Else
       cmdOpcion(1).Enabled = True
       cmdOpcion(2).Enabled = True
       cmdOpcion(12).Enabled = True
       cmdOpcion(16).Enabled = True
        If lNCElimina Then
         cmdOpcion(2).Enabled = False
        End If
        If lParcialNC Then
             Me.frmDetalleNC.Enabled = False
             cmdOpcion(9).Enabled = False
        End If
        If lactivaFechaNC Then
          dtpFecha.Enabled = False
        Else
          dtpFecha.Enabled = True
        End If
        
        cmdParcial(1).Enabled = True
        cmdParcial(2).Enabled = True
        cmdParcial(3).Enabled = True
        cmdParcial(4).Enabled = True
        
    End If
    cmdTexto.Caption = "Registro " & frmNotaCredito.RsCabecera.AbsolutePosition & " de " & frmNotaCredito.RsCabecera.RecordCount
    
    Set RsDoc = Lib.OpenRecordset(" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','',1", Cn)
    Set grdDocumento.DataSource = RsDoc
    
    Set RsDocNC = Lib.OpenRecordset(" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','',2", Cn)
    Set grdNotaCredito.DataSource = RsDocNC
    
    Call CargaTotalesNC
    
    On Error Resume Next
    
    
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 2 'Previo
                MoverPuntero previo, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmNotaCredito.grdGrilla
                Asignar
                cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
    End Select
End Sub

Private Sub cmdNotaCredito_Click(Index As Integer)
   Dim xDescripcion As String
   Dim xRsNotaCredito As Recordset
   
        If pais = "002" Then 'Ecuador
           Set xRsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 1 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
        Else
           Set xRsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lActivo = 1 order by tTipoEmision", Cn)
        End If
   
   Select Case Index
   
       Case Is = 1
           xDescripcion = Trim(cmdNotaCredito(Index).Caption)
           xRsNotaCredito.Find "Descripcion='" & xDescripcion & "'"
           txtNC1.Text = xRsNotaCredito!prefijo
           xTipoNotaCredito = xRsNotaCredito!TTipoEmision
           lNotaCreditoEmitirFE = xRsNotaCredito!lFacturacionElectronica
           cmdNotaCredito(Index).Enabled = False
           cmdNotaCredito(2).Enabled = True
           tPrefijoEnlace = xRsNotaCredito!tPrefijoEnlace
       Case Is = 2
           xDescripcion = Trim(cmdNotaCredito(Index).Caption)
           xRsNotaCredito.Find "Descripcion='" & xDescripcion & "'"
           txtNC1.Text = xRsNotaCredito!prefijo
           xTipoNotaCredito = xRsNotaCredito!TTipoEmision
           lNotaCreditoEmitirFE = xRsNotaCredito!lFacturacionElectronica
           cmdNotaCredito(Index).Enabled = False
           cmdNotaCredito(1).Enabled = True
           tPrefijoEnlace = xRsNotaCredito!tPrefijoEnlace
   End Select
   
   If lNotaCreditoEmitirFE And tPrefijoEnlace <> "" Then
        cmdOpcion(4).Enabled = False
        txtPrefijo.Text = tPrefijoEnlace
   End If
   
   If pais = "003" And tPrefijoEnlace <> "" Then
        cmdOpcion(4).Enabled = False
        txtPrefijo.Text = tPrefijoEnlace
   End If

End Sub

Private Sub cmdOpcion_Click(Index As Integer)

    Dim oComando As clsComando
    Dim oComandoDetalle As clsComando
    Dim sImporteLetra As String
    Dim RsCodigoHash As New ADODB.Recordset
    Dim fDocumento As String
    Dim lcodigoHash As Boolean
    
    Dim oComandoCabeceraOfisis As clsComando
    Dim oComandoDetalleOfisis As clsComando
    Dim oComandoFirmaDocumentoOfisis As clsComando
    Dim fso1 As Object
                                            
    Dim RsT3 As Recordset
    Dim RsT_P As Recordset
                                            
   Select Case Index
          Case Is = 0 'Agregar
          
               If Supervisor("27") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
               End If
          
               ActivarBotones (False)
               Blanquear Me
               Sw = True
               nNeto = 0
               nImpuesto1 = 0
               nImpuesto2 = 0
               nImpuesto3 = 0
               nTotal = 0
               
               nNCNeto = 0
               nNCImp1 = 0
               nNCImp2 = 0
               nNCImp3 = 0
               nNCTotal = 0
               
               txtNCNeto.Text = "0.00"
               txtNCImp1.Text = "0.00"
               txtNCImp2.Text = "0.00"
               txtNCImp3.Text = "0.00"
               txtNCTotal.Text = "0.00"
               
               txtDocNeto.Text = "0.00"
               txtDocImp1.Text = "0.00"
               txtDocImp2.Text = "0.00"
               txtDocImp3.Text = "0.00"
               txtDocTotal.Text = "0.00"
               
               txtResNeto.Text = "0.00"
               txtResImp1.Text = "0.00"
               txtResImp2.Text = "0.00"
               txtResImp3.Text = "0.00"
               txtResTotal.Text = "0.00"
               
               cmdOpcion(1).Enabled = True
               cmdOpcion(3).Enabled = True
               cmdOpcion(4).Enabled = True
               cmdOpcion(5).Enabled = True
               cmdOpcion(6).Enabled = True
               cmdOpcion(8).Enabled = False
               cmdOpcion(9).Enabled = False
               dtpFecha.Enabled = True
               
               RsNotaCredito.Requery
               AsignaComando 2, RsNotaCredito, cmdNotaCredito()
               If RsNotaCredito.RecordCount = 1 Then
                  cmdNotaCredito(1).Enabled = True
               Else
                  cmdNotaCredito(1).Enabled = True
                  cmdNotaCredito(2).Enabled = True
               End If

               
               'cmdOpcion(10).Enabled = True
               'cmdOpcion(11).Enabled = True
               'txtPrefijo.Text = "F"
               'txtNC1.Text = "N"
               
          Case Is = 1 'Grabar
               'Chequea Datos
               Dim nCorrela As String
               Dim nPos As Integer

                'FACTURACION E
               If txtNC1.Text = "" Then MsgBox "Seleccione el Tipo de Documento", vbExclamation, sMensaje: Exit Sub
               If txtPrefijo.Text = "" Then MsgBox "Ingrese el Documento a Afectar", vbExclamation, sMensaje: Exit Sub
               If txtSerie.Text = "" Then MsgBox "Ingrese la Serie del Documento a Afectar", vbExclamation, sMensaje: Exit Sub
               If txtCorrela.Text = "" Then MsgBox "Ingrese el Correlativo del Documento a Afectar", vbExclamation, sMensaje: Exit Sub
               'If nNCTotal <= 0 Then MsgBox "El valor de la Nota de Crédito debe ser mayor a cero", vbExclamation, sMensaje: Exit Sub
               
               If nNCTotal <= 0 Then MsgBox "El valor de la Nota de crédito debe ser mayor a cero", vbExclamation, sMensaje: Exit Sub
               
               If nNCImp1 < 0 Then MsgBox "El valor Impuesto de la Nota de crédito no debe ser negativo!!!", vbExclamation, sMensaje: Exit Sub
               If nNCImp2 < 0 Then MsgBox "El valor Impuesto de la Nota de crédito no debe ser negativo!!!", vbExclamation, sMensaje: Exit Sub
               If nNCImp3 < 0 Then MsgBox "El valor Impuesto de la Nota de crédito no debe ser negativo!!!", vbExclamation, sMensaje: Exit Sub
               
               If txtObservacion.Text = "" Then MsgBox "Ingrese el Motivo de la Nota de Crédito", vbExclamation, sMensaje: Exit Sub
               
               If nNCTotal > CDbl(Me.txtDocTotal) Then MsgBox "El valor de la Nota de crédito no puede ser mayor al documento", vbExclamation, sMensaje: Exit Sub
               
               
                Dim CantNotaCredito3, CantDocumento3, SumLey1_3  As Double
                Dim numnotacredito3 As Integer
                
                SumLey1_3 = Calcular("SELECT isnull(nley1,0) as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                CantDocumento3 = Calcular("SELECT nVenta as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                CantNotaCredito3 = Calcular("SELECT sum(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04' and tnotacredito<>'" & sNC & "'", Cn)
                
                numnotacredito3 = Calcular("SELECT count(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)
                nNCTotal = nNCTotal - CDbl(SumLey1_3)
                'CantNotaCredito2 = Val(CantNotaCredito2) + Val(nNCTotal)

                If CDbl(Round((CDbl(CantNotaCredito3) + CDbl(nNCTotal)), 2)) > Round(CantDocumento3, 2) Then
                 MsgBox ("!La cantidad Asiganada + las cantidades de las Notas de creditos Generados al documento es Mayor al monto Del documento ¡")
                   nNCTotal = CDbl(CantDocumento3) - CDbl(CantNotaCredito3) - CDbl(SumLey1_3)
                 Exit Sub
                End If
               
               
               Dim ObligaMotivoAdm As Boolean
               ObligaMotivoAdm = Calcular("select ISNULL(lMotivoAdm,0) as Codigo FROM TCAJA where tCaja = '" & sCaja & "'", Cn)

               'ObligaMotivoAdm = Calcular("select lMotivoAdm from TCAJA where tCaja = '" & sCaja & "'", Cn)

               If ObligaMotivoAdm Then
                    If txtMotivoAdm.Text = "" Then MsgBox "Ingrese el Motivo Administrativo de la Nota de Crédito", vbExclamation, sMensaje: Exit Sub
               End If

                                          
               If (lFEBiz) Then
                    Dim EstadoDoc As String
                    EstadoDoc = Calcular("select top 1 CodigoEstadoBizlink  as codigo from TESTADOBIZLINK where tDocumento= '" & sDocumento & "'  order by fregistro_rpt desc", Cn)
                    
                    If EstadoDoc <> "SIGNED/AC_03" Then
                        If (lBloquearNCsinEstado) Then
                            MsgBox "El Documento: documento Nro: " & Format(sDocumento, "@-@@@@@-@@@@@@@@@") & ", aún no se encuentra aceptado en SUNAT", vbExclamation, sMensaje: Exit Sub
                            Exit Sub
                         Else
                            If MsgBox("El Documento: documento Nro: " & Format(sDocumento, "@-@@@@@-@@@@@@@@@") & ", aún no se encuentra aceptado en SUNAT, Desea continuar con la creación de la NC ?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                                Exit Sub
                            End If
                        End If
                    End If
               End If
              
                                          
                                          
               Dim sNSerie As String
               Dim sNPrefijo As String
               Dim sNCorrela As String
               Dim sNTipoEmision As String
               Dim sNDocumento As String
               '---------------
               If Sw Then
                  Sw = False
                   
                'FACTURACION E
                RsNotaCredito.Requery
                RsNotaCredito.MoveFirst
                RsNotaCredito.Find ("tTipoEmision='" & xTipoNotaCredito & "'")
    
                sNSerie = RsNotaCredito!tSerie
                sNPrefijo = RsNotaCredito!prefijo
                sNCorrela = Lib.Correlativo(RsNotaCredito!tUltimoNumero, 9)
                sNTipoEmision = RsNotaCredito!TTipoEmision
                sNC = sNPrefijo & sNSerie & sNCorrela
                Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sNCorrela & "' where tTipoEmision ='" & sNTipoEmision & "' and tCaja ='" & sCaja & "'"
                '-------------
                                                      
                'Cambiar el SQL
                Isql = "insert into MNOTACREDITO( " & _
                         "tNotaCredito, fFecha, tDocumento, nNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
                         "tEstadoDocumento, tTurno, tCaja, tTipoDocumento, tMotivo, tUsuario, tObservacion, fDiaContable, fRegistro, MotivoAdm) " & _
                         "values ('" & sNC & "', '" & _
                                Format(dtpFecha.Value, "yyyy/mm/dd") + " " + Format(Time, "hh:mm:ss") & "' , " & _
                                "'" & sDocumento & "', " & _
                                nNCNeto & ", " & _
                                nNCImp1 & ", " & _
                                nNCImp2 & ", " & _
                                nNCImp3 & ", " & _
                                nNCTotal & ", " & _
                                "'01', " & _
                                " '" & sTurno & "', " & _
                                " '" & sCaja & "', " & _
                                " '" & xTipoNotaCredito & "', " & _
                                " '" & xCodigoMotivo & "', " & _
                                "'" & sUsuario & "'," & _
                                " '" & txtObservacion.Text & "', '" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & _
                                "getdate(), '" & txtMotivoAdm.Text & "')"
                  Cn.Execute Isql
                  
                  
                  'cambio anulacion por notas de credito
                  If modProcedimiento.pasa = True Then
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.Find "tNotaCredito ='" & sNC & "'"
                  End If
                  
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  
                    txtNC1.Text = Mid(sNC, 1, 1)
                    If pais = "002" Then
                        txtNC2.Text = Mid(sNC, 2, 6)
                        txtNC3.Text = Mid(sNC, 8, 9)
                    Else
                        txtNC2.Text = Mid(sNC, 2, 5)
                        txtNC3.Text = Mid(sNC, 7, 9)
                    End If
                  
                  ' cambio anulacion por nota de credito
                  ActivarBotones (True)
                  If modProcedimiento.pasa = True Then
                  cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.MoveLast
                  End If

               Else
               
               
                  'Cambiar el SQL
                  Isql = "update MNOTACREDITO set " & _
                         "nNeto =" & nNCNeto & ", " & _
                         "nImpuesto2 =" & nNCImp2 & ", " & _
                         "nImpuesto3 =" & nNCImp3 & ", " & _
                         "nVenta =" & nNCTotal & ", " & _
                         "fFecha ='" & Format(dtpFecha.Value, "yyyy/mm/dd") + " " + Format(Time, "hh:mm:ss") & "', " & _
                         "nImpuesto1 =" & nNCImp1 & ", " & _
                         "tMotivo ='" & xCodigoMotivo & "', " & _
                         "fDiaContable ='" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & _
                         "tObservacion ='" & txtObservacion.Text & "',lreplica=1,MotivoAdm = '" & txtMotivoAdm.Text & "'" & _
                         "where tNotaCredito= '" & sNC & "'"

                  Cn.Execute Isql
                  
                  
                  
                  'Cambiar el Nombre del Formulario
                  'cambios nota de credito anulacion
                  If modProcedimiento.pasa = True Then
                  nPos = frmNotaCredito.RsCabecera.AbsolutePosition
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.AbsolutePosition = nPos
                  End If
                  MsgBox "Registro Modificado", vbInformation, sMensaje
                  
                  cmdOpcion(12).Enabled = True
               End If
               
               Cn.Execute " update DNOTACREDITO set tNotaCredito = '" & sNC & "' where tNotaCredito = '" & sDocumento & "'"
               
                'GCAA 01032021
               If (lNCElimina = True) Then
                    cmdOpcion(2).Enabled = False
               End If
          
          Case Is = 2 'Eliminar
          
               If frmNotaCredito.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If

                frmNotaCredito.RsCabecera.Requery
                frmNotaCredito.RsCabecera.MoveLast
               
               lblPaso1.Visible = True
                lblPaso2.Visible = True
                imgProceso(0).Visible = False
                imgProceso(1).Visible = False
                imgProceso(2).Visible = False
                imgProceso(3).Visible = False
                FrameFeSpring.Visible = False
               
               If frmNotaCredito.RsCabecera!tTurno = sTurno Then
                   'Password
                   If Supervisor("05") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                Else
                   'Password
                   If MsgBox("El Documento es de un turno Anterior, deseas continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                      Exit Sub
                   End If
                   
                   If Supervisor("06") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If

               If MsgBox("Seguro de Eliminar la Nota de Crédito Nro." & sNC & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                   
               If pais = "000" Then
                    If lFacturacionE Then
                          Dim lDocElecInfofact, lDocElecInfofactOfisis As Boolean
                          Dim xDocumentoVenta As String
                          Dim RsDocNotaCreditoFE As Recordset
                          lDocElecInfofact = Calcular("select isnull(tdi.lFacturacionElectronica,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MNOTACREDITO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tNotaCredito= '" & sNC & "'", Cn)
                          lDocElecInfofactOfisis = Calcular("select isnull(tdi.lDocumentoElectronicoOfisis,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MNOTACREDITO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tNotaCredito= '" & sNC & "'", Cn)
                          xDocumentoVenta = Calcular("select tdocumento as codigo from MNOTACREDITO where tNotaCredito='" & sNC & "'", Cn)
                          
                          If lDocElecInfofact And lDocElecInfofactOfisis = False Then
                            If lFEpape Then
                                If (Calcular("select isnull(tEstadoDocumento,'01')  as codigo from mnotacredito where tnotacredito='" & sNC & "'", Cn) = "01") Then
                                    Dim tTipoDocNotacredito, xUltimoCorrelativo As String
                                    tTipoDocNotacredito = Calcular("select isnull(ttipodocumento,'')  as codigo from mnotacredito where tnotacredito='" & sNC & "'", Cn)
                                    Cn.Execute "Delete mnotacredito Where tnotacredito= '" & sNC & "'"
                                    xUltimoCorrelativo = Calcular("select MAX(tnotacredito) as codigo from mnotacredito where tcaja='" & sCaja & "' and tTipoDocumento='" & tTipoDocNotacredito & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & tTipoDocNotacredito & "' and tCaja ='" & sCaja & "'"
                                    frmNotaCredito.RsCabecera.Requery
                                    Unload Me
                                    Exit Sub
                                Else
                                    If (Calcular("select isnull(tEstadoDocumento,'01')  as codigo from mnotacredito where tnotacredito='" & sNC & "'", Cn) = "01") Then
                                        Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                                        frmNotaCredito.RsCabecera.Requery
                                    End If
                                End If
                            
                            ElseIf lFESpring Then
                            ElseIf lFEBiz Then
                            ElseIf lFEGesa Then
                            ElseIf lFECarbajal Then
                                    Label2.Caption = "   Proceso de anulación de documento en InfoFact......."
                                    lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                    lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                    sImporteLetra = NumeroCadena(str(Round(nNCTotal, 2))) + " " + sMonedaN
                                    FrameFeSpring.Visible = True
                                    Sleep 1000
                                    If Not INSERTAFE_CARVAJAL(sNC, sImporteLetra, 1, 1) Then '----CABECERA
                                            imgProceso(2).Visible = True
                                            imgProceso(3).Visible = True
                                            Sleep 1000
                                            FrameFeSpring.Visible = False
                                            Exit Sub
                                     End If
                                     imgProceso(0).Visible = True
                                     imgProceso(1).Visible = True
                                     Sleep 1500
                                     FrameFeSpring.Visible = False
                                     Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                                     frmNotaCredito.RsCabecera.Requery
                            Else ' INFOFACT
                                Isql = "Select * From dbo.MNOTACREDITO where tNotaCredito='" & sNC & "' and isnull(lEstadoFacturacion,0)=1 and testadodocumento<>'01';"
                                Set RsDocNotaCreditoFE = Lib.OpenRecordset(Isql, Cn)
                                If RsDocNotaCreditoFE.RecordCount > 0 Then
                                    Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04', fRegistroAnulado = GETDATE(), tUsuarioAnulado ='" & sUsuario & "' ,lreplica=1 where tNotaCredito = '" & sNC & "'"
                                    frmNotaCredito.RsCabecera.Requery
                                Else
                                    If MsgBox("La Nota de Credito: " & sNC & vbNewLine & " no se encuentra procesada o no se ha enviado a facturacion electronica, " & vbNewLine & " Favor de revisar el documento!,¿Desea anular el documento?", vbYesNo, sMensaje) = vbYes Then
                                        Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04', fRegistroAnulado = GETDATE(), tUsuarioAnulado ='" & sUsuario & "' ,lreplica=1 where tNotaCredito = '" & sNC & "'"
                                        frmNotaCredito.RsCabecera.Requery
                                    End If
                                End If
                            End If
                          Else
                            If lDocElecInfofactOfisis Then
                                MsgBox "La interfaz con Ofisis no soporta Anulación, La anulación se realizara solo para el sistema!!!", vbInformation
                            End If
                            Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04', fRegistroAnulado = GETDATE(), tUsuarioAnulado ='" & sUsuario & "' ,lreplica=1 where tNotaCredito = '" & sNC & "'"
                            frmNotaCredito.RsCabecera.Requery
                          End If
                    Else
                          Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04', fRegistroAnulado = GETDATE(), tUsuarioAnulado ='" & sUsuario & "' ,lreplica=1 where tNotaCredito = '" & sNC & "'"
                          frmNotaCredito.RsCabecera.Requery
                    End If
               Else
                    Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04', fRegistroAnulado = GETDATE(), tUsuarioAnulado ='" & sUsuario & "' ,lreplica=1 where tNotaCredito = '" & sNC & "'"
                    frmNotaCredito.RsCabecera.Requery
               End If
               

               If frmNotaCredito.RsCabecera.RecordCount <> 0 Then
                  frmNotaCredito.RsCabecera.Requery
                  frmNotaCredito.RsCabecera.Find "tNotaCredito ='" & sNC & "'"
                  Asignar
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If

          Case Is = 3 'Salir
                If sNC = "" Then
                    Cn.Execute " DELETE FROM DNOTACREDITO where tNotaCredito =' " & sDocumento & " '"
                End If
          
               Unload Me
          
          Case Is = 4 'Tipo Documento
               RsTipoDocumento.MoveNext
               If RsTipoDocumento.EOF Then
                  RsTipoDocumento.MoveFirst
               End If
               cmdOpcion(4).Caption = RsTipoDocumento!descripcion
               txtPrefijo.Text = RsTipoDocumento!prefijo
                                       
          Case Is = 5 'KB Numero de Serie
               sDescrip = ""
               frmKeyBoard.txtResultado.Text = sDescrip
               frmKeyBoard.Show vbModal
               If wEnter Then
                    If pais = "002" Then 'ECUADOR
                       sDescrip = Mid(Trim(sDescrip), 1, 6)
                       txtSerie.Text = Mid("000000", 1, 6 - Len(Trim(sDescrip))) & Trim(sDescrip)
                    Else
                       sDescrip = Mid(Trim(sDescrip), 1, 5)
                       txtSerie.Text = Mid("00000", 1, 5 - Len(Trim(sDescrip))) & Trim(sDescrip)
                    End If
               End If
               
               sDescrip = ""

          Case Is = 6 'KB correlativo
               sTipo = "Numero"
               frmNumPad.Show vbModal
               If wEnter Then
                  txtCorrela.Text = Mid("000000000", 1, 9 - Len(Trim(sDescrip))) & Trim(sDescrip)
               End If
               
             
               'Consistencia
               If txtPrefijo.Text = "" Then
                  txtCorrela.Text = ""
                  Exit Sub
               End If
               
               If txtSerie.Text = "" Then
                  txtCorrela.Text = ""
                  Exit Sub
               End If
               
               
               'Busqueda
               sDocumento = txtPrefijo.Text & txtSerie.Text & txtCorrela.Text
            
            
               If lFacturacionE Then
                    If lNotaCreditoEmitirFE Then
                        If lCajaContingencia Then 'GCAA 22032021
                        
                        Else
                            If tPrefijoEnlace <> txtPrefijo.Text Then
                                MsgBox "Error : El tipo de documento " & sDocumento & " no se puede asociar a un tipo " & tPrefijoEnlace & " de Nota de Credito ", vbExclamation, sMensaje
                                txtPrefijo.Text = ""
                                txtSerie.Text = ""
                                txtCorrela.Text = ""
                                Exit Sub
                            End If
                        End If
                    End If
               End If
               
               
               ' validar  si esta cancelado el documento
               Set RsDocumento = Lib.OpenRecordset("select * from MDOCUMENTO where tDocumento ='" & sDocumento & "' and tEstadoDocumento = '01'", Cn)
               If RsDocumento.RecordCount = 1 Then
                  MsgBox "Error : Documento no Cancelado", vbExclamation, sMensaje
                  txtCorrela.Text = ""
                  Set RsDocumento = Nothing
                  Exit Sub
               End If

               Set RsDocumento = Lib.OpenRecordset("select * from MDOCUMENTO where tDocumento ='" & sDocumento & "' and tEstadoDocumento <>'04'", Cn)
               If RsDocumento.RecordCount = 0 Then
                  MsgBox "Error : Documento no Existe", vbExclamation, sMensaje
                  txtCorrela.Text = ""
                  Set RsDocumento = Nothing
                  Exit Sub
               End If

               Dim CantNotaCredito, CantDocumento, SumLey1 As Double
               Dim numnotacredito As Integer
               
                SumLey1 = Calcular("SELECT isnull(nley1,0) as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                CantDocumento = Calcular("SELECT nVenta as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                CantNotaCredito = Calcular("SELECT sum(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)
                numnotacredito = Calcular("SELECT count(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)
                
               If CDbl(CantNotaCredito) >= CDbl(CantDocumento) And numnotacredito > 0 Then
                MsgBox ("!Ya se ha generado (" & numnotacredito & ") Notas de Credito por el total del Documento¡")
                 txtCorrela.Text = ""
                 Set RsDocumento = Nothing
                Exit Sub
               End If
               
                If Periodo(RsDocumento!fdiacontable, vbNewLine + "Documento de referencia con fecha de un periodo cerrado!!!") Then
                    txtCorrela.Text = ""
                    Exit Sub
                End If
                
                If (Calcular(" select (select isnull(lOpcion29,0) from tgrupousuario where tgrupousuario= tusuario.tgrupousuario) as codigo from tusuario  where tCodigoUsuario='" & tcodigoUsuarioA & "'", Cn)) Then
                    If (Format(RsDocumento!fdiacontable, "YYYYMM") <> Format(FechaServidor, "YYYYMM")) Then
'                        If Supervisor("29") = False Then
                          MsgBox "El documento de referencia fue emitida fuera del mes, Usuario sin permiso!!!", vbExclamation, sMensaje
                          txtCorrela.Text = ""
                          Exit Sub
'                          Exit Sub
'                        End If
                    End If
               End If
               'Mensaje de documento pertenece a otra caja solo FE activo
               If lFacturacionE Then
                    If RsDocumento!tCaja <> sCaja Then
                        MsgBox "Tener en consideración que el Documento a canjear no pertenece a esta Caja", vbExclamation, sMensaje
                    End If
               End If
               
               
               Set RsDoc = Lib.OpenRecordset(" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','','',1", Cn)
               Set grdDocumento.DataSource = RsDoc
               
               Set RsDocNC = Lib.OpenRecordset(" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','','',2", Cn)
               Set grdNotaCredito.DataSource = RsDocNC
               
               cmdOpcion(4).Enabled = False
               cmdOpcion(5).Enabled = False
               cmdOpcion(6).Enabled = False
               'cmdOpcion(10).Enabled = False
               'cmdOpcion(11).Enabled = False
               cmdOpcion(8).Enabled = True
               cmdOpcion(9).Enabled = True
               cmdOpcion(16).Enabled = True
               
               cmdNotaCredito(1).Enabled = False
               cmdNotaCredito(2).Enabled = False
               
               cmdParcial(1).Enabled = True
               cmdParcial(2).Enabled = True
               cmdParcial(3).Enabled = True
               cmdParcial(4).Enabled = True
                
               nNeto = RsDocumento!nNeto
               nImpuesto1 = RsDocumento!nprecioImpuesto1
               nImpuesto2 = RsDocumento!nprecioImpuesto2
               nImpuesto3 = RsDocumento!nprecioImpuesto3
               nTotal = RsDocumento!nVenta
               
               txtDocNeto.Text = Format(nNeto, "#,###,##0.00")
               txtDocImp1.Text = Format(nImpuesto1, "#,###,##0.00")
               txtDocImp2.Text = Format(nImpuesto2, "#,###,##0.00")
               txtDocImp3.Text = Format(nImpuesto3, "#,###,##0.00")
               txtDocTotal.Text = Format(nTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")

            If lParcialNC Then
                
                Call cmdParcial_Click(2)
            
'                Dim Acum As Double
'                cmdOpcion(9).Enabled = False
'                  nNCTotal = Round(nTotal, 2)
'
'                  If CDbl(nNCTotal) > CDbl(txtDocTotal.Text) Then
'                    MsgBox ("!La cantidad Asignada no puede ser Mayor al monto Del documento¡"), vbInformation
'                    nNCTotal = CDbl(txtDocTotal.Text)
'                  End If
'                  Acum = 0
'                  Acum = IIf(nPorcentaje1 > 0, Acum + nPorcentaje1, Acum)
'                  Acum = IIf(nPorcentaje2 > 0, Acum + nPorcentaje2, Acum)
'                  Acum = IIf(nPorcentaje3 > 0, Acum + nPorcentaje3, Acum)
'                  Acum = 1 + (Acum / 100)
'                    Select Case pais ' ok
'                        Case "001" 'Bolivia
'                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal * nPorcentaje1 / 100, 0)
'                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal * nPorcentaje2 / 100, 0)
'                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal * nPorcentaje3 / 100, 0)
'                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
'
'                        Case Else 'Peru, Ecuador
'                                nNCImp1 = IIf(nPorcentaje1 > 0, nNCTotal / Acum * nPorcentaje1 / 100, 0)
'                                nNCImp2 = IIf(nPorcentaje2 > 0, nNCTotal / Acum * nPorcentaje2 / 100, 0)
'                                nNCImp3 = IIf(nPorcentaje3 > 0, nNCTotal / Acum * nPorcentaje3 / 100, 0)
'                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
'
'                        End Select
'
'                nNCNeto = Calcular("select isnull(nNeto,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
'                nNCImp1 = Calcular("select isnull(nPrecioImpuesto1,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
'                nNCImp2 = Calcular("select isnull(nPrecioImpuesto2,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
'                nNCImp3 = Calcular("select isnull(nPrecioImpuesto3,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
'                nNCTotal = Calcular("select isnull(nVenta,0) as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
'
'                txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
'                txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
'                txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
'                txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
'                txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
'               'End If
'
'               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
'               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
'               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
'               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
'               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
'
'               Me.txtNotaLey1.Text = Format(SumLey1, "#,###,##0.00")
            Else
                Call CargaTotalesNC
            End If
            
            'Me.txtDocLey1.Text = Format(SumLey1, "#,###,##0.00")
            
          Case Is = 8 'KB Observacion
          
                    Select Case pais
                    
                      Case "000"
                      'GCAA 02032021
                          If lFacturacionE Then
'                                If lNotaCreditoEmitirFE Then
                                        sTemp = ""
                                        Isql = "Select * From vMotivoNotaCredito order by Codigo"
                                        
                                        frmBusquedaRapida.cmdOpcion(1).Enabled = False
                                        frmBusquedaRapida.cmdOpcion(2).Enabled = False
                                        frmBusquedaRapida.cmdOpcion(3).Enabled = False
                                        frmBusquedaRapida.nPredeterm = 1
                                        
                                        Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                          "Motivo", 2, "Descripcion", 6600, 0, 0, "")
                                                          
                                        sTemp = IIf(sTemp = "0", "", sTemp)
                                        xCodigoMotivo = ""
                                        frmBusquedaRapida.Show vbModal
                                        
                                        If wEnter = True And sCodigo <> "" Then
                                           xCodigoMotivo = sCodigo
                                           txtObservacion.Text = Calcular("SELECT tDetallado As codigo FROM ttabla WHERE tTabla='MOTIVONOTACREDITO' And tCodigo='" & xCodigoMotivo & "'", Cn)
                                        Else
                                           Exit Sub
                                        End If
'                                Else
'                                        frmKeyBoard.txtResultado.Text = txtObservacion.Text
'                                        frmKeyBoard.Show vbModal
'                                        txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
'                                End If
                          Else
                              frmKeyBoard.txtResultado.Text = txtObservacion.Text
                              frmKeyBoard.Show vbModal
                              txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
                          End If
                      
                      Case Else
                              frmKeyBoard.txtResultado.Text = txtObservacion.Text
                              frmKeyBoard.Show vbModal
                              txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
                              
                    End Select
                               
          Case Is = 9 'NumPad Cantidad
               Dim Acumulado As Double
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter Then
                  nNCTotal = Round(CDbl(sDescrip), 2)
                  
                  If CDbl(nNCTotal) > CDbl(txtDocTotal.Text) Then
                  MsgBox ("!La cantidad Asignada no puede ser Mayor al monto Del documento¡"), vbInformation
                  nNCTotal = CDbl(txtDocTotal.Text)
                  End If
                  
                    Dim CantNotaCredito2, CantDocumento2, SumLey1_2  As Double
                    Dim numnotacredito2 As Integer
                    
                    SumLey1_2 = Calcular("SELECT isnull(nley1,0) as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                    CantDocumento2 = Calcular("SELECT nVenta as codigo FROM MDOCUMENTO WHERE TDOCUMENTO = '" & sDocumento & "'", Cn)
                    CantNotaCredito2 = Calcular("SELECT sum(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04' and tnotacredito<>'" & sNC & "'", Cn)
                    
                    numnotacredito2 = Calcular("SELECT count(nVenta) as codigo FROM MNOTACREDITO WHERE tDocumento= '" & sDocumento & "' AND tEstadoDocumento <>'04'", Cn)
                    nNCTotal = nNCTotal - CDbl(SumLey1_2)
                    'CantNotaCredito2 = Val(CantNotaCredito2) + Val(nNCTotal)

                    If (CDbl(CantNotaCredito2) + CDbl(nNCTotal) > CantDocumento2) Then
                     MsgBox ("!La cantidad Asiganada + las cantidades de las Notas de creditos Generados al documento es Mayor al monto Del documento ¡")
                       nNCTotal = CDbl(CantDocumento2) - CDbl(CantNotaCredito2) - CDbl(SumLey1_2)
                     'Exit Sub
                    End If

                  Acumulado = 0
                  Acumulado = IIf(nPorcentaje1 > 0 And CDbl(Me.txtDocImp1.Text) > 0, Acumulado + nPorcentaje1, Acumulado)
                  Acumulado = IIf(nPorcentaje2 > 0 And CDbl(Me.txtDocImp2.Text) > 0, Acumulado + nPorcentaje2, Acumulado)
                  Acumulado = IIf(nPorcentaje3 > 0 And CDbl(Me.txtDocImp3.Text) > 0, Acumulado + nPorcentaje3, Acumulado)
                  Acumulado = 1 + (Acumulado / 100)
                    Select Case pais ' ok
                        Case "001" 'Bolivia
                                nNCImp1 = IIf(nPorcentaje1 > 0 And CDbl(Me.txtDocImp1.Text) > 0, nNCTotal * nPorcentaje1 / 100, 0)
                                nNCImp2 = IIf(nPorcentaje2 > 0 And CDbl(Me.txtDocImp2.Text) > 0, nNCTotal * nPorcentaje2 / 100, 0)
                                nNCImp3 = IIf(nPorcentaje3 > 0 And CDbl(Me.txtDocImp3.Text) > 0, nNCTotal * nPorcentaje3 / 100, 0)
                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
                                
                        Case Else 'Peru, Ecuador
                                nNCImp1 = IIf(nPorcentaje1 > 0 And CDbl(Me.txtDocImp1.Text) > 0, nNCTotal / Acumulado * nPorcentaje1 / 100, 0)
                                nNCImp2 = IIf(nPorcentaje2 > 0 And CDbl(Me.txtDocImp2.Text) > 0, nNCTotal / Acumulado * nPorcentaje2 / 100, 0)
                                nNCImp3 = IIf(nPorcentaje3 > 0 And CDbl(Me.txtDocImp3.Text) > 0, nNCTotal / Acumulado * nPorcentaje3 / 100, 0)
                                nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
                                
                    End Select
                  
               End If
               nNCTotal = nNCTotal + SumLey1_2
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
               
               'Me.txtDocLey1.Text = Format(SumLey1_2, "#,###,##0.00")
               Me.txtNotaLey1.Text = Format(SumLey1_2, "#,###,##0.00")
                        
          Case Is = 12 'Procesar
               Dim xPedido As String
               Dim xEstadoNC As String
               Dim DocGuia As Boolean
               DocGuia = False
                lblPaso1.Visible = True
                lblPaso2.Visible = True
                imgProceso(0).Visible = False
                imgProceso(1).Visible = False
                imgProceso(2).Visible = False
                imgProceso(3).Visible = False
                FrameFeSpring.Visible = False
               
               If nNCTotal <= 0 Then MsgBox "El valor de la Nota de crédito debe ser mayor a cero", vbExclamation, sMensaje: Exit Sub
               
               If nNCImp1 < 0 Then MsgBox "El valor Impuesto de la Nota de crédito no debe ser negativo!!!", vbExclamation, sMensaje: Exit Sub
               If nNCImp2 < 0 Then MsgBox "El valor Impuesto de la Nota de crédito no debe ser negativo!!!", vbExclamation, sMensaje: Exit Sub
               If nNCImp3 < 0 Then MsgBox "El valor Impuesto de la Nota de crédito no debe ser negativo!!!", vbExclamation, sMensaje: Exit Sub
               If nNCTotal > CDbl(Me.txtDocTotal) Then MsgBox "El valor de la Nota de crédito no puede ser mayor al documento", vbExclamation, sMensaje: Exit Sub
               
               If Calcular("select isnull(tEstadoDocumento,'') as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn) = "03" Then
                MsgBox ("El documento de referencia esta en estado por cobrar, " & vbNewLine & "es necesario una forma de pago para la emision de nota de credito, revisar!!"), vbInformation, sMensaje
                Exit Sub
               End If
               If Calcular("select count(*) as codigo from mguiatransporte where tdocumento='" & sDocumento & "'", Cn) > 0 Then
                DocGuia = True
               End If
               
               xEstadoNC = Calcular("Select ISNULL(tEstadoDocumento,'') As Codigo From MNOTACREDITO Where tNotaCredito = '" & sNC & "' ", Cn)
               impTipo = "0"
               
               'CAMBIO CSHM 22/08/2025 QA INFOREST
               IsqlFact = "exec usp_inforest_Impresion '" & sNC & "',10 "
               'IsqlFact = "exec usp_inforest_Impresion '" & sDocumento & "',10 "
               Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
               
               If xEstadoNC = "01" Then
                      RsNotaCredito.Requery
                      RsNotaCredito.MoveFirst
                      Dim DOCAnticipo As String
                      DOCAnticipo = Calcular("select isnull(treserva,'') as Codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                      If pais = "002" Then
                          RsNotaCredito.Find ("tSerie='" & Mid(sNC, 2, 6) & "'")
                      Else
                          RsNotaCredito.Find ("tSerie='" & Mid(sNC, 2, 5) & "'")
                      End If
                   
                      If MsgBox("Deseas Procesar la Nota de Crédito Nro: " & sNC & " ? ", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                        
                      
                             TimpresionDolaresDelivery = False
                             '------ impresion en dolares para check de cliente delivery
                             If Calcular("select isnull(lImpresionMonedaExtranjera,0) as codigo from MDOCUMENTO where tDocumento='" & sDocumento & "'", Cn) Then
                                     TimpresionDolaresDelivery = True
                                     Cn.Execute "update mnotacredito set lImpresionMonedaExtranjera=1 where tnotacredito='" & sNC & "'"
                             Else
                                 TimpresionDolaresDelivery = False
                                 Cn.Execute "update mnotacredito set lImpresionMonedaExtranjera=0 where tnotacredito='" & sNC & "'"
                             End If
                            
                            '-----------------------
                            If pais = "000" And lFEpape And IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                 If Not FacturarTCPIP(2, sNC, 1) Then
                                    Cn.Execute "update mnotacredito set lImpresionMonedaExtranjera=0 where tnotacredito='" & sNC & "'"
                                    Exit Sub
                                 End If
                            End If
                            '------------------------
                            If Round(nNCTotal, 2) = Round(nTotal, 2) Then
                                    xPedido = Calcular("select min(tCodigoPedido) as codigo from DDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
                                    If IsNull(xPedido) Then
                                      MsgBox "Error : Documento sin pedido", vbExclamation, sMensaje
                                      Exit Sub
                                    End If
                                    If Calcular("select count(*) as codigo from mguiatransporte where tdocumento='" & sDocumento & "'", Cn) = 0 Then
                                        If Not lFECarbajal And DOCAnticipo = "" And pais <> "003" And DocGuia = False Then
                                            Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                            Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                        End If
                                    Else
                                        Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                    End If
                                    'FACTURACION ELECTRONICA ECUADOR
                                    If pais = "002" Then
                                        If lFacturacionE And lFEEcuador = False Then
                                           If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                'CABECERA
                                                Set oComando = New clsComando
                                                If Not oComando.CreateCmdSp("USP_InsertaFactNotaCredito", Cn) Then
                                                     Set oComando = Nothing
                                                     Exit Sub
                                                End If
                                                oComando.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
                        
                                                If Not oComando.GetParamOK Then
                                                     Set oComando = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComando.ExecSP Then
                                                     Set oComando = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'DETALLE
                                                Set oComandoDetalle = New clsComando
                                                If Not oComandoDetalle.CreateCmdSp("USP_InsertaFactNotaCreditoDetalle", Cn) Then
                                                     Set oComandoDetalle = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoDetalle.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
                            
                                                If Not oComandoDetalle.GetParamOK Then
                                                     Set oComandoDetalle = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoDetalle.ExecSP Then
                                                     Set oComandoDetalle = Nothing
                                                     Exit Sub
                                                End If
                                           End If
                                           If lFacturacionE And lFEEcuador Then
                                                 If INSERTA_FE_INFOREST(sDocumento, 2, DateTime.Now) = False Then
                                                     MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                                                 End If
                                           End If
                                        End If
                                    End If
    
                            Else
                                    If Not lFECarbajal Then
                                        Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                    End If
                            End If
                        
                            If DOCAnticipo <> "" Then
                                Cn.Execute "update mdocumento set treserva='' where tdocumento='" & sDocumento & "'"
                            End If
                            
                            'FACTURACION_E_PERU
                            If pais = "000" Then
                                If lFacturacionE Then
                                    If lFEOfisis Then 'OFISIS
                                            '----CABECERA
                                            Set oComandoCabeceraOfisis = New clsComando
                                            If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactNotaCreditoOfisis", Cn) Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sNC
                
                                            If Not oComandoCabeceraOfisis.GetParamOK Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComandoCabeceraOfisis.ExecSP Then
                                                 Set oComandoCabeceraOfisis = Nothing
                                                 Exit Sub
                                            End If
                                            
                                            '----FIRMA DOCUMENTO OFISIS
                                            If RsNotaCredito!lDocumentoElectronicoOfisis Then
                                                Set oComandoFirmaDocumentoOfisis = New clsComando
                                                If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sNC
                    
                                                If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                     Set oComandoFirmaDocumentoOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                'VALIDAR RESPUESTA CODIGO DE BARRA
                                                fDocumento = Mid(sDocumento, 1, 1) + Mid(sNC, 4, 3) + "-" + CStr(CLng(Mid(sNC, 8, 8)))
                                            End If
                                            
                                    ElseIf lFESpring Then
                                        If Round(nNCTotal, 2) = Round(nTotal, 2) And DOCAnticipo = "" And DocGuia = False Then
                                            Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                            Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                        Else
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                        End If
                                        
                                    ElseIf lFECarbajal Then
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            
                                            Label2.Caption = "   Proceso de envio de documento a InfoFact......."
                                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                            sImporteLetra = NumeroCadena(str(Round(nNCTotal, 2))) + " " + sMonedaN
                                            FrameFeSpring.Visible = True
                                            Sleep 1000
                                            If Not INSERTAFE_CARVAJAL(sNC, sImporteLetra, 1, 0) Then '----CABECERA
                                                    imgProceso(2).Visible = True
                                                    imgProceso(3).Visible = True
                                                    Sleep 1000
                                                    FrameFeSpring.Visible = False
                                                    Exit Sub
                                             End If
                                             imgProceso(0).Visible = True
                                             imgProceso(1).Visible = True
                                             Sleep 1500
                                             FrameFeSpring.Visible = False
                                             impTipo = "1"
                                        End If
                                        If Round(nNCTotal, 2) = Round(nTotal, 2) And DOCAnticipo = "" And DocGuia = False Then
                                            Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                            Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                        Else
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                        End If
                                        
                                        
                                    ElseIf lFEpape Then
                                        impTipo = "1" 'IMPRESION FORMATO ELECTRONICO
                                    ElseIf lFEBiz Then
                                        impTipo = "1"
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            If Not INSERTA_FE_INFOREST(sNC, 2, DateTime.Date) Then '----CABECERA
                                                Cn.Execute "update MNOTACREDITO set tEstadodocumento='01' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                                MsgBox "Eror Bizlink: No se pudo procesar la Nota de credito!! Favor de verificar la informacion", vbInformation, sMensaje
                                                Exit Sub
                                             End If
                                             Sleep 1500
                                        End If
                                        If Round(nNCTotal, 2) = Round(nTotal, 2) And DOCAnticipo = "" And DocGuia = False Then
                                            Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                            Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                        Else
                                            Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                        End If
                                        
                                    ElseIf lFEGesa Then
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            'CABECERA
                                            Set oComando = New clsComando
                                            If Not oComando.CreateCmdSp("USP_FactNotaCredito", Cn) Then
                                                 Set oComando = Nothing
                                                 Exit Sub
                                            End If
                                            oComando.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
                                            oComando.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 250, ""
                                            If Not oComando.GetParamOK Then
                                                 Set oComando = Nothing
                                                 'Exit Sub
                                            End If
                                            If Not oComando.ExecSP Then
                                                 Set oComando = Nothing
                                                 'Exit Sub
                                            End If
                                            impTipo = "1" 'IMPRESION FORMATO ELECTRONICO
                                        End If
                                    
                                    
                                    Else 'INFOFACT
                                        If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                            ' DESARROLLADOR: ELDCQ - 20/04/2023
                                            If Not INSERTAFE(sNC, "", 2, "") Then
                                                MsgBox "Ocurrio un problema al enviar el documento, Revisar la configuraciones electronicas en el Sistema!!!"
                                                Screen.MousePointer = vbDefault
                                                'Exit Sub
                                            End If
                                            ' FIN DE CAMBIOS 20/04/2023
                                            
'                                            sImporteLetra = NumeroCadena(str(Round(nNCTotal, 2))) + " " + sMonedaN
'                                            'CABECERA
'                                            Set oComando = New clsComando
'                                            If Not oComando.CreateCmdSp("USP_FactNotaCredito", Cn) Then
'                                                 Set oComando = Nothing
'                                                 Exit Sub
'                                            End If
'                                            oComando.CreateParameter "@NotaCredito", adVarChar, adParamInput, 20, sNC
'                                            oComando.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 250, sImporteLetra
'
'                                            If Not oComando.GetParamOK Then
'                                                 Set oComando = Nothing
'                                                 Exit Sub
'                                            End If
'                                            If Not oComando.ExecSP Then
'                                                 Set oComando = Nothing
'                                                 Exit Sub
'                                            End If
                                            
                                            impTipo = "1" 'IMPRESION FORMATO ELECTRONICO
                                            fDocumento = Mid(sDocumento, 1, 1) + Mid(sNC, 4, 3) + Mid(sNC, 8, 8)
                                            
                                        End If
                                    
                                    End If

                                End If
                            End If
                            
                            
                            
                            'IMPRESION NC
                            Imprimir (sImp)
                            Printer.FontName = sFont
                            Printer.FontBold = False
                            
                            Dim RsImpresion As Recordset
                            'CAMBIO CSHM 22/08/2025 QA INFOREST MODIFICAR SP
                            Isql = "exec usp_inforest_Impresion '" & sNC & "',11 "
                            'Isql = "Select * From vNotaCreditoImpresora Where tNotaCredito='" & sNC & "'"
                            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                            
                            Dim rstFuente As Recordset
                            Set rstFuente = New ADODB.Recordset
                            imageCab.Picture = Nothing
                            imagepIE.Picture = Nothing
                            Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                            imageCab.DataField = "foto"
                            Set imageCab.DataSource = rstFuente
                            imagepIE.DataField = "fotoPie"
                            Set imagepIE.DataSource = rstFuente
                            sNTipoEmision = RsNotaCredito!TTipoEmision
                            
                            If pais = "000" Then 'PERU
                                If lFacturacionE Then
                                        If lFEOfisis Then
                                                If RsNotaCredito!lDocumentoElectronicoOfisis Then
                                                  impTipo = "1"
                                                  Sleep 2000
                                                  
                                                  If lImpresionCodigoBarras Then
                                                        imageHash.DataField = "foto"
                                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','D','' ", Cn)
                                                        Set imageHash.DataSource = RsCodigoHash
                                                        
                                                    ElseIf lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, "D"))
                                                  Else
                                                        Dim RscadenaCodigoHash As Recordset

                                                        Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','1' ", Cn)
                                                        If RscadenaCodigoHash.RecordCount > 0 Then
                                                            cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                        End If
                                                        'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and TI_DOCU='D' ", CnFE)
                                                  End If
                                                  
                                                  ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                  impTipo = "0"
                                                  ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                End If
                                        
                                        ElseIf lFESpring Then
                                        
                                            ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                            
                                        ElseIf lFECarbajal Then
                                            If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sNC, 1))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sNC, 1))
                                                         Else
                                                             cadenaCodigoHash = ImagenFeCarvajal(2, sNC, 1)
                                                         End If
                                                     End If
                                                 End If
                                                If RsNotaCredito!tFormulario = "01" Then
                                                    ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    ImprimeFormatoA
                                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                    If fso1.FileExists(App.path & "\fact.bmp") Then
                                                        Kill App.path & "\fact.bmp"
                                                    End If
                                                End If
                                            End If
                                        ElseIf lFEpape Then
                                            If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                               If RsNotaCredito!tFormulario = "01" Then
                                                    If tCodigoFE = "000" Then
                                                         If lQRFE Then
                                                             Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                                         Else
                                                             If lImpresionCodigoBarras Then
                                                                 'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                             Else
                                                                 cadenaCodigoHash = PapeMatricial
                                                             End If
                                                         End If
                                                     End If
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    CrearImagenQR (PapeTermico)
                                                    ImprimeFormatoA
                                                    Kill App.path & "\BaseTempQr.bmp"
                                                End If
                                            End If
                                        ElseIf lFEBiz Then
                                            If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sNC, 1))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sNC, 1))
                                                         Else
                                                             cadenaCodigoHash = QRHASH_FE_INFOREST(2, sNC, 1)
                                                         End If
                                                     End If
                                                 End If
                                                If RsNotaCredito!tFormulario = "01" Then
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    ImprimeFormatoA
                                                    
                                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                    If fso1.FileExists(App.path & "\fact.bmp") Then
                                                        Kill App.path & "\fact.bmp"
                                                    End If
                                                End If
                                            End If
                                        ElseIf lFEGesa Then
                                        
                                             If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                             
                                                If lQRFE Then
                                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sNC, 1))
                                                Else
                                                    If lImpresionCodigoBarras Then
                                                        Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sNC, 1))
                                                    Else
                                                        cadenaCodigoHash = QRHASH_FE_INFOREST(2, sNC, 1)
                                                    End If
                                                End If
                                                     
                                                If RsNotaCredito!tFormulario = "01" Then
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                Else
                                                    ImprimeFormatoA
                                                    
                                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                                    If fso1.FileExists(App.path & "\fact.bmp") Then
                                                        Kill App.path & "\fact.bmp"
                                                    End If
                                                End If
                                            End If
                                        
                                        Else 'INFOFACT
                                                If IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                                     If RsNotaCredito!tFormulario = "01" Then
                                                         'VALIDAR RESPUESTA DE CODIGO HASH
                                                         If tCodigoFE = "000" Then
                                                         
                                                            If lQRFE Then
                                                                Set imageHash.Picture = LoadPicture(ImagenQR(sNC))
                                                            Else
                                                            
                                                                If lImpresionCodigoBarras Then
                                                                    Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sNC))
                                                                Else
                                                                    cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sNC)
                                                                End If
                                                            End If

                                                         End If
                                                                                          
                                                         ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                         
                                                     Else
                                                         'FORMATO A4
                                                         If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "' and tipodocu = '07'", "imagen", "\fact.bmp") = True Then
                                                            ImprimeFormatoA
                                                            Kill App.path & "\fact.bmp"
                                                         Else
                                                            ImprimeFormatoA
                                                         End If
                                                     End If
                                                 
                                                Else 'NO ELECTRONICO
                                                     ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                                End If
                                        End If
                                    
                                 Else 'NO ELECTRONICO
                                      ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                 End If

                            ElseIf pais = "003" Then
                                If lFETusFacturasAPP = True Then
                                    If Not INSERTA_FE_INFOREST_ARGE(sNC, 2, DateTime.Date) Then '----CABECERA
                                       Cn.Execute "update MNOTACREDITO set tEstadodocumento='01' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                       MsgBox "Eror TUSFACTURASAPP: No se pudo procesar la Nota de credito!! Favor de verificar la informacion", vbInformation, sMensaje
                                       Exit Sub
                                    End If
                                    Sleep 1000
                                    
                                    If Round(nNCTotal, 2) = Round(nTotal, 2) And DOCAnticipo = "" And DocGuia = False Then
                                        Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                        Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                        Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                    Else
                                        Cn.Execute "update MNOTACREDITO set tEstadodocumento='05' ,lreplica=1 where tNotaCredito='" & sNC & "'"
                                    End If
                                    
                                    Dim Termico As String
                                    Termico = Calcular("select isnull(tImprTermica,'') as codigo from MNOTACREDITO where tnotacredito='" & sNC & "'", Cn)
                                    If Termico <> "" Then
                                        Set imageHash.Picture = LoadPicture(CrearImagenQR(Termico))
                                        ImprimeNotaCreditoARGE RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                                    Else
                                       ' MsgBox "Por el momento los servicios de AFIP estan caidos o hubo un percance en la facturacion el documento intentara enviarse en uno minutos a mas tardar 1 dia, favor indicar al cliente que revise su correo electronico para la recepcion de la Nota de Credito: " & sNC
                                        Dim MENSAJERROR As String '''
                                        MENSAJERROR = Calcular("select isnull(error,'') as codigo from DOCUMENTO_ARGENTINA where documento='" & sNC & "'", Cn)
                                        MsgBox MENSAJERROR, vbInformation
                                     
                                    End If

                                Else
                                    Dim RespDev As Integer
                                    RespDev = ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 3, RsImpresion)
                                    If RespDev = 0 Then
                                        MsgBox "Ocurrio un inconveniente en la facturacion con la impresora fiscal", vbInformation, "Facturacion impresora fiscal"
                                        Exit Sub
                                        'GoTo fin
                                    ElseIf RespDev = 1 Then
                                        Cn.Execute "update MNOTACREDITO set tEstadoDocumento = '04',lreplica=1 where tNotaCredito = '" & sNC & "'"
                                        frmNotaCredito.RsCabecera.Requery
                                        MsgBox "Ocurrio un incoveninete en la impresion, favor de revisar funete de papel, fuente de energia!!!", vbInformation, "Facturacion impresora Fiscal"
                                    Else
                                    
                                        Cn.Execute "update MPEDIDO set tEstadoPedido ='01'  where tCodigoPedido in (select tcodigopedido from dpedido where tdocumento='" & sDocumento & "')"
                                        Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                        Cn.Execute "update MNOTACREDITO set tEstadodocumento='05',lreplica=1 where tNotaCredito='" & sNC & "'"
                                End If
                                End If
                            Else 'ECUADOR
                                ImprimeNotaCredito RsImpresion, imageHash, impTipo, cadenaCodigoHash, TimpresionDolaresDelivery, imageCab, imagepIE, sNTipoEmision
                            End If
                            '------------------------------
                            '-----------------------
                            Set RsT3 = Lib.OpenRecordset("select distinct tCodigoPedido from DPEDIDO where tDocumento = '" & sDocumento & "'", Cn)
                            If Calcular("select count(*) as codigo from mguiatransporte where tdocumento='" & sDocumento & "'", Cn) > 0 Then
                            
                                '19/05/2025
                                Set RsT_P = Lib.OpenRecordset("select isnull(sum(nVenta),0) as total from MNOTACREDITO where tDocumento = '" & sDocumento & "' and tEstadoDocumento in ('02','05')", Cn)
                                
                                If Round(RsT_P!TOTAL, 2) = Round(nTotal, 2) Then
                                
                                    'Cn.Execute "Update mguiatransporte set tEstadoGuia = '03',tdocumento='' where tdocumento ='" & sDocumento & "'"
                                    Cn.Execute "update MGUIATRANSPORTE set tEstadoGuia = '03', tUsuarioAnulado='" & IIf(IsNull(sVar1), "NOTA DE CREDITO", sVar1) & "', fRegistroAnulado=getdate() where tdocumento ='" & sDocumento & "'"
                                    Cn.Execute "update mpedido set testadopedido='04', lReplica=1 where tcodigopedido in (select tCodigoPedido from dpedido where tGuiaTransporte in (select tGuiaTransporte from MGUIATRANSPORTE where tDocumento = '" & sDocumento & "') group by tCodigoPedido)"
                                    Cn.Execute "update DPEDIDO set tGuiaTransporte='',tFacturado ='' where tGuiaTransporte in (select tGuiaTransporte from MGUIATRANSPORTE where tDocumento = '" & sDocumento & "')"
                                    
                                    'If RsT3.RecordCount <> 0 Then
                                      'RsT3.MoveFirst
                                      'Do While Not RsT3.EOF
                                        'Cn.Execute "Update MPEDIDO set tEstadoPedido = '04', lReplica=1  where tCodigoPedido ='" & RsT3!tCodigoPedido & "'"
                                        'RsT3.MoveNext
                                      'Loop
                                      'Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                    'End If
                                End If
                                
                                
                                
                                'Cn.Execute "Update mguiatransporte set tEstadoGuia = '01',tdocumento='' where tdocumento ='" & sDocumento & "'"
                                
                                'If RsT3.RecordCount <> 0 Then
                                '  RsT3.MoveFirst
                                '  Do While Not RsT3.EOF
                                '    Cn.Execute "Update MPEDIDO set tEstadoPedido = '04', lReplica=1  where tCodigoPedido ='" & RsT3!tCodigoPedido & "'"
                                '    RsT3.MoveNext
                                '  Loop
                                '  Cn.Execute "update DPEDIDO set tDocumento ='', tFacturado ='' where tdocumento='" & sDocumento & "'"
                                'End If
                                
                            End If
                            
                            If lNcOfisis Then
                                Dim MNTurno As String
                                Dim MNDocRef As String
                                'Dim nMontoDoc As Float
                                'MNTurno = Calcular("select isnull(tturno,'') as codigo from mnotacredito where tNotaCredito='" & sNC & "'", Cn)
                                'GCAA 23032021
                                MNTurno = Calcular("select isnull(tturno,'') as codigo from MDOCUMENTO where tDocumento='" & sDocumento & "'", Cn)
                                MNDocRef = Calcular("select isnull(tdocumento,'') as codigo from mnotacredito where tNotaCredito='" & sNC & "'", Cn)
                                If MNTurno = sTurno And Round(nNCTotal, 2) = Round(nTotal, 2) Then
                                    Cn.Execute "delete from dpagodocumento where tdocumento='" & MNDocRef & "'"
                                    Cn.Execute "UPDATE VC " _
                                                & "SET VC.lESTADO_USO = 0 " _
                                                & "FROM DPAGODOCUMENTO_VC PVC " _
                                                & "INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                                & "WHERE PVC.tDocumento = '" & MNDocRef & "' AND " _
                                                & "CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "
                
                                    Cn.Execute "DELETE PVC " _
                                                & " FROM DPAGODOCUMENTO_VC PVC " _
                                                & " INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                                & " WHERE PVC.tDocumento = '" & MNDocRef & "' AND " _
                                                & " CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "
                                    
                                     Isql = "insert into DPAGODOCUMENTO " & _
                                     "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, fRegistro, tUsuario,fDiaContable ) " & _
                                     "Values(   '" & MNDocRef & "'," _
                                                & "1," _
                                                & "'" & sTurno & "'," _
                                                & "'04'," _
                                                & "'002'," _
                                                & "'01'," _
                                                & nTC & ", " _
                                                 & Round(nNCTotal, 2) & ", " _
                                                & "'" & sNC & "', " _
                                                & "'', " _
                                                & "getdate()," _
                                                & "'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "')"
                                    Cn.Execute Isql
'                                    If sOtroTipoCancelacion = "001" Then
'                                        Cn.Execute "update MINGRESO set tEstadoDocumento ='02' where tRecibo ='" & sTipoDocumento & "'"
'                                    ElseIf sOtroTipoCancelacion = "002" Then
                                    Cn.Execute "update MNOTACREDITO set tEstadoDocumento ='02',lreplica=1 where tNotaCredito ='" & sNC & "'"
'                                    End If
                                    
                                    
                                End If
                            End If
                            
                            If pais = "000" And lFEpape And IIf(RsNotaCredito!lFacturacionElectronica = True, 1, 0) Then
                                 If Not FacturarTCPIP(3, sNC, 1) Then
                                    MsgBox ("La confirmacion ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
                                 End If
                            End If
                            '------------------------
                            If modProcedimiento.pasa = True Then
                                frmNotaCredito.RsCabecera.Requery
                                frmNotaCredito.RsCabecera.Find "tNotaCredito ='" & sNC & "'"
                                Asignar
                            Else
                                Unload Me
                            End If
                            
                        
                      End If
                 
              End If
               
               
          Case Is = 13 'Correccion Impuesto1
               sTipo = ""
               frmNumPad.Show vbModal
               
               If wEnter Then
                  nNCImp1 = Val(sDescrip)
                  nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
               End If
               
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
                    
                    
          Case Is = 14 'Correccion Impuesto1
               sTipo = ""
               frmNumPad.Show vbModal
               
               If wEnter Then
                  nNCImp2 = Val(sDescrip)
                  nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
               End If
               
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
                    
          Case Is = 15 'Correccion Impuesto1
               sTipo = ""
               frmNumPad.Show vbModal
               
               If wEnter Then
                  nNCImp3 = Val(sDescrip)
                  nNCNeto = nNCTotal - nNCImp1 - nNCImp2 - nNCImp3
               End If
               
               txtNCNeto.Text = Format(nNCNeto, "#,###,##0.00")
               txtNCImp1.Text = Format(nNCImp1, "#,###,##0.00")
               txtNCImp2.Text = Format(nNCImp2, "#,###,##0.00")
               txtNCImp3.Text = Format(nNCImp3, "#,###,##0.00")
               txtNCTotal.Text = Format(nNCTotal, "#,###,##0.00")
               
               txtResNeto.Text = Format(nNeto - nNCNeto, "#,###,##0.00")
               txtResImp1.Text = Format(nImpuesto1 - nNCImp1, "#,###,##0.00")
               txtResImp2.Text = Format(nImpuesto2 - nNCImp2, "#,###,##0.00")
               txtResImp3.Text = Format(nImpuesto3 - nNCImp3, "#,###,##0.00")
               txtResTotal.Text = Format(nTotal - nNCTotal, "#,###,##0.00")
                    
          Case Is = 16 'KB Observacion
          
                    Select Case pais
                    
                      Case "000"
                            
                              frmKeyBoard.txtResultado.Text = txtMotivoAdm.Text
                              frmKeyBoard.Show vbModal
                              txtMotivoAdm.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
                      
                      Case Else
                              frmKeyBoard.txtResultado.Text = txtMotivoAdm.Text
                              frmKeyBoard.Show vbModal
                              txtMotivoAdm.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
                              
                    End Select
   End Select
End Sub

Private Sub ImprimeFormatoA()
                    
                    Dim xImpresionFE As String
                    
                    'NEW
                    Dim RsImpresionNC As Recordset
                    Isql = "Select * From MNOTACREDITO Where tNotaCredito='" & sNC & "'"
                    Set RsImpresionNC = Lib.OpenRecordset(Isql, Cn)
                    
                    Dim xMotivoNT As String
                    xMotivoNT = Calcular("Select ISNULL(tMotivo,'06') As Codigo from MNOTACREDITO Where tNotacredito = '" & sNC & "'", Cn)
                    
                    Dim xNeto As String
                    Dim xVenta As String
                    Dim xImp1 As String
                    Dim xImp2 As String
                    
                    xNeto = Format(RsImpresionNC!nNeto, "##,###,##0.00")
                    xVenta = Format(RsImpresionNC!nVenta, "##,###,##0.00")
                    xImp1 = Format(RsImpresionNC!nImpuesto1, "##,###,##0.00")
                    xImp2 = Format(RsImpresionNC!nImpuesto2, "##,###,##0.00")
                    '----------------------
                    
                    
                    xImpresionFE = Calcular("SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MNOTACREDITO WHERE tNotaCredito='" & sNC & "')", Cn)

                    If RsNotaCredito!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                    
                    If xMotivoNT = "06" Then
                        Dim Reporte As New dsrNotaCredito
                        
                        Reporte.DiscardSavedData
                        Reporte.Database.SetDataSource RsImpDocumentoE
                        
                        Reporte.Text13.SetText "NOTA DE CREDITO ELECTRONICA"
                        
                        Reporte.Text8.SetText sRazonSocial
                        Reporte.ReportTitle = sDireccion
                        Reporte.Text15.SetText sTelefono
                        Reporte.Text14.SetText sFax
                        Reporte.Text16.SetText sRUC
                        Reporte.Text50.SetText sWeb
                        
                        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MNOTACREDITO WHERE tNotaCredito='" & sNC & "')", Cn) = 1 Then
                        Reporte.ReportComments = tTextoAgenteRetencion
                        End If
                        
                        xMontoTexto = "SON: " & NumeroCadena(str(RsImpDocumentoE!nVenta)) & " " & sMonedaN
                        Reporte.Text4.SetText xMontoTexto
                        Reporte.Text31.SetText xImpresionFE
    
'                    frmEmite.CRViewer.DisplayGroupTree = False
'                    frmEmite.CRViewer.ReportSource = Reporte
'                    frmEmite.CRViewer.ViewReport
'                    frmEmite.Show vbModal
    
                        Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Reporte.PaperOrientation = crPortrait
                        Reporte.PrintOut False, 1, False, 1, 1
                        '----------------
                        
                        
                    Else
                    
                        Dim Reporte1 As New dsrNotaCreditoObservacion
                        
                        Reporte1.DiscardSavedData
                        Reporte1.Database.SetDataSource RsImpDocumentoE
                        
                                         
                        Reporte1.Text13.SetText "NOTA DE CREDITO ELECTRONICA"
                        
                        Reporte1.Text8.SetText sRazonSocial
                        Reporte1.ReportTitle = sDireccion
                        Reporte1.Text15.SetText sTelefono
                        Reporte1.Text14.SetText sFax
                        Reporte1.Text16.SetText sRUC
                        
                        Reporte1.Text29.SetText xVenta
                        Reporte1.Text36.SetText xNeto
                        Reporte1.Text38.SetText xImp1
                        Reporte1.Text45.SetText xImp2
                        Reporte1.Text49.SetText xVenta
                        Reporte1.Text50.SetText sWeb
                        
                        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MNOTACREDITO WHERE tNotaCredito='" & sNC & "')", Cn) = 1 Then
                        Reporte1.ReportComments = tTextoAgenteRetencion
                        End If
                        
                        xMontoTexto = "SON: " & NumeroCadena(str(xVenta)) & " " & sMonedaN
                        Reporte1.Text4.SetText xMontoTexto
                        Reporte1.Text31.SetText xImpresionFE
    
'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = Reporte1
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal

                        Reporte1.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Reporte1.PaperOrientation = crPortrait
                        Reporte1.PrintOut False, 1, False, 1, 1
                                       
                    End If
                                    
                    If iImagenCab Then
                       Kill App.path & "\cliente.jpg"
                    End If
End Sub

Private Sub cmdParcial_Click(Index As Integer)
On Error GoTo fin
    cmdOpcion(12).Enabled = False
    Select Case Index
        
          Case Is = 1 'Agregar Item
                
                If RsDoc.RecordCount > 0 Then
                     frmIngresaItem.Visible = True
                     Me.txtItemPuDoc.Text = Format(RsDoc!nprecioVenta, "#,###,##0.000") '
                     Me.txtItemCanDoc.Text = Format(RsDoc!nCantidad, "#,###,##0.000000") '
                     Me.txtItemTotDoc.Text = Format(RsDoc!nVenta, "#,###,##0.000") '
                     
                     Me.txtItemPuNC.Text = Format(RsDoc!nprecioVenta, "#,###,##0.000") '
                     Me.txtItemCanNC.Text = Format(RsDoc!nCantidad, "#,###,##0.000000") '
                     Me.txtItemTotNC.Text = Format(RsDoc!nVenta, "#,###,##0.000") '
                End If
            Exit Sub
          Case Is = 2 'Agregar Todos
                Cn.Execute (" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','',3")
          Case Is = 3 'Quitar un Item
                If RsDocNC.RecordCount > 0 Then
                    Cn.Execute (" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','" & RsDocNC!tItem & "',4")
                End If
          Case Is = 4 'Quitar Todos los items
                Cn.Execute (" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','',5")
          Case Is = 5 ' Grabar datos
                frmIngresaItem.Visible = False
                If RsDoc.RecordCount > 0 Then
                    Cn.Execute (" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','" & RsDoc!tItem & "',7," & CDbl(Me.txtItemPuNC.Text) & "," & CDbl(Me.txtItemCanNC.Text) & " ")
                End If
          Case Is = 6 ' Cancelar datos
          
                frmIngresaItem.Visible = False
                Me.txtItemPuDoc.Text = Format(0, "#,###,##0.000")
                Me.txtItemCanDoc.Text = Format(0, "#,###,##0.000000")
                Me.txtItemTotDoc.Text = Format(0, "#,###,##0.000")
                
                Me.txtItemPuNC.Text = Format(0, "#,###,##0.000")
                Me.txtItemCanNC.Text = Format(0, "#,###,##0.000000")
                Me.txtItemTotNC.Text = Format(0, "#,###,##0.000")
                Exit Sub
          Case Is = 7 ' Modifica Cantidad
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter Then
                  If CDbl(sDescrip) = 0 Then
                    MsgBox "Cantidad no puede ser 0!!!.", vbInformation, sMensaje
                    Exit Sub
                  End If
                  If CDbl(sDescrip) > CDbl(Me.txtItemCanDoc.Text) Then
                    MsgBox "Cantidad no puede ser mayor al documento!!!.", vbInformation, sMensaje
                    Exit Sub
                  End If
               
                  Me.txtItemCanNC.Text = Format(CDbl(sDescrip), "#,###,##0.000000")
                  Me.txtItemTotNC.Text = Format(CDbl(sDescrip) * CDbl(Me.txtItemPuNC.Text), "#,###,##0.000")
               End If
               Exit Sub
          Case Is = 8 ' Modifica Monto
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter Then
                  If CDbl(sDescrip) = 0 Then
                    MsgBox "Monto no puede ser 0!!!.", vbInformation, sMensaje
                    Exit Sub
                  End If
                  If CDbl(sDescrip) > CDbl(Me.txtItemTotDoc.Text) Then
                    MsgBox "Monto no puede ser mayor al documento!!!.", vbInformation, sMensaje
                    Exit Sub
                  End If
               
                  Me.txtItemTotNC.Text = Format(CDbl(sDescrip), "#,###,##0.000")
                  Me.txtItemCanNC.Text = Format(CDbl(sDescrip) / CDbl(Me.txtItemPuNC.Text), "#,###,##0.000000")
                  'Me.txtItemPuNC.Text = Format(CDbl(sDescrip) / CDbl(Me.txtItemCanNC.Text), "#,###,##0.00")
               End If
               Exit Sub
          Case Else

   End Select
   
   If Index <> 1 And Index <> 7 And Index <> 8 Then
        Call CargaTotalesNC
   End If
   
    Set RsDoc = Lib.OpenRecordset(" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','',1", Cn)
    Set grdDocumento.DataSource = RsDoc
    
    Set RsDocNC = Lib.OpenRecordset(" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','',2", Cn)
    Set grdNotaCredito.DataSource = RsDocNC
    
   RsDoc.Requery
   RsDocNC.Requery
    
   Exit Sub
fin:
    MsgBox "Error: " & Error, vbInformation, sMensaje
End Sub

Private Sub CargaTotalesNC()
On Error GoTo fin
        Dim RsCargaTotalNC As Recordset
        Set RsCargaTotalNC = Lib.OpenRecordset(" EXEC [dbo].[SP_NOTACREDITO_PARCIAL] '" & sDocumento & "','" & sNC & "','',6", Cn)
        If RsCargaTotalNC.RecordCount > 0 Then
            RsCargaTotalNC.MoveFirst
            
            'nNCTotal = nNCTotal + SumLey1_2
            txtNCNeto.Text = Format(RsCargaTotalNC!Neto, "#,###,##0.00")
            txtNCImp1.Text = Format(RsCargaTotalNC!IMP1, "#,###,##0.00")
            txtNCImp2.Text = Format(RsCargaTotalNC!Imp2, "#,###,##0.00")
            txtNCImp3.Text = Format(RsCargaTotalNC!Imp3, "#,###,##0.00")
            txtNCTotal.Text = Format(RsCargaTotalNC!Venta, "#,###,##0.00")
            txtNotaLey1.Text = Format(RsCargaTotalNC!Ley1, "#,###,##0.00")

            nNCNeto = RsCargaTotalNC!Neto
            nNCImp1 = RsCargaTotalNC!IMP1
            nNCImp2 = RsCargaTotalNC!Imp2
            nNCImp3 = RsCargaTotalNC!Imp3
            nNCTotal = RsCargaTotalNC!Venta

        Else
            txtNCNeto.Text = Format(0, "#,###,##0.00")
            txtNCImp1.Text = Format(0, "#,###,##0.00")
            txtNCImp2.Text = Format(0, "#,###,##0.00")
            txtNCImp3.Text = Format(0, "#,###,##0.00")
            txtNCTotal.Text = Format(0, "#,###,##0.00")
            txtNotaLey1.Text = Format(0, "#,###,##0.00")
            
            nNCNeto = 0
            nNCImp1 = 0
            nNCImp2 = 0
            nNCImp3 = 0
            nNCTotal = 0
        End If

Exit Sub
fin:
    MsgBox "Error: " & Error, vbInformation, sMensaje
End Sub

Private Sub Form_Activate()
    Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumento where Codigo <> '00'  and Canjear= 1 and lActivo = 1", Cn)
    If RsTipoDocumento.RecordCount = 0 Then
        MsgBox "No hay Tipos de Documentos con opción de Canje por Nota de Crédito", vbCritical
        Unload Me
    End If
        
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Centrar Me
       
   'Ingrese el SubTitulo
   Me.Caption = " Mantenimiento de Notas de Crédito "
   fraDetalle.Caption = Me.Caption
      
      If lactivaFechaNC Then
        dtpFecha.Enabled = False
      Else
        dtpFecha.Enabled = True
      End If
      

    
    'FACTURACION E
    
    
    If pais = "002" Then 'Ecuador
      Set RsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 1 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
    Else
      Set RsNotaCredito = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 1 And lActivo = 1 order by tTipoEmision", Cn)
    End If
    
    sImp = RsNotaCredito!timpresora
    nEmision = RsNotaCredito.RecordCount
    
    
'    If nEmision = 0 Then
'      MsgBox "No se ha ingresado los Documentos de Nota de Credito por Caja", vbCritical
'      Unload Me
'    End If

    AsignaComando 2, RsNotaCredito, cmdNotaCredito()
    '------
    
   
   If Sw = True Then
      ActivarBotones (False)
      Blanquear Me
      dtpFecha.Value = FechaServidor()
      nNeto = 0
      nImpuesto1 = 0
      nImpuesto2 = 0
      nImpuesto3 = 0
      nTotal = 0
      
      nNCNeto = 0
      nNCImp1 = 0
      nNCImp2 = 0
      nNCImp3 = 0
      nNCTotal = 0
      
      txtNCNeto.Text = "0.00"
      txtNCImp1.Text = "0.00"
      txtNCImp2.Text = "0.00"
      txtNCImp3.Text = "0.00"
      txtNCTotal.Text = "0.00"
      
      txtDocNeto.Text = "0.00"
      txtDocImp1.Text = "0.00"
      txtDocImp2.Text = "0.00"
      txtDocImp3.Text = "0.00"
      txtDocTotal.Text = "0.00"
      
      txtResNeto.Text = "0.00"
      txtResImp1.Text = "0.00"
      txtResImp2.Text = "0.00"
      txtResImp3.Text = "0.00"
      txtResTotal.Text = "0.00"
      txtDocLey1.Text = "0.00"
      txtNotaLey1.Text = "0.00"
      
      cmdOpcion(8).Enabled = False
      cmdOpcion(9).Enabled = False
      'txtPrefijo.Text = "F"
      
   Else
      ActivarBotones (True)
      Asignar
   End If
   
   If nPorcentaje1 = 0 Then
      Label(5).Visible = False
      txtNCImp1.Visible = False
      txtDocImp1.Visible = False
      txtResImp1.Visible = False
      cmdOpcion(13).Visible = False
   Else
      Label(5).Caption = sImpuesto1 & " : "
   End If
   
   If nPorcentaje2 = 0 Then
      Label(6).Visible = False
      txtNCImp2.Visible = False
      txtDocImp2.Visible = False
      txtResImp2.Visible = False
      cmdOpcion(14).Visible = False
   Else
      Label(6).Caption = sImpuesto2 & " : "
   End If
   
   If nPorcentaje3 = 0 Then
      Label(7).Visible = False
      txtNCImp3.Visible = False
      txtDocImp3.Visible = False
      txtResImp3.Visible = False
      cmdOpcion(15).Visible = False
   Else
      Label(7).Caption = sImpuesto3 & " : "
   End If

    If lNCElimina Then
     cmdOpcion(2).Enabled = False
    End If
    If lParcialNC Then
         Me.frmDetalleNC.Enabled = False
         cmdOpcion(9).Enabled = False
    End If
    If lactivaFechaNC Then
      dtpFecha.Enabled = False
    Else
      dtpFecha.Enabled = True
    End If
    
    If sLey1 <> "" Then
        Me.lblLey1(12).Caption = sLey1 + ":"
        Me.lblLey1(12).Visible = True
        Me.txtDocLey1.Visible = True
        Me.txtNotaLey1.Visible = True

    End If
    
  If modProcedimiento.pasa = True Then
   cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
  End If
   
    Call ConfGrilla(7, grdDocumento, "Documento", 2, "tDocumento", 0, 1, 0, "#0", _
                 "Item", 2, "tItem", 0, 0, 0, "", _
                 "Cod", 2, "tCodigoProducto", 0, 1, 0, "###,###,##0.00", _
                 "Producto", 2, "Descripcion", 1900, 3, 0, "###,###,##0.00", _
                 "P.Und", 2, "nPrecioVenta", 800, 1, 0, "###,###,##0.00", _
                 "Cant.", 2, "nCantidad", 750, 1, 0, "#,##0.000", _
                 "Total", 2, "nVenta", 1000, 1, 0, "###,###,##0.00")
               
    grdDocumento.Columns(0).Visible = False
    grdDocumento.Columns(1).Visible = False
    grdDocumento.Columns(2).Visible = False
   
    Call ConfGrilla(7, grdNotaCredito, "Documento", 2, "tDocumento", 0, 1, 0, "#0", _
                 "Item", 2, "tItem", 0, 0, 0, "", _
                 "Cod", 2, "tCodigoProducto", 0, 1, 0, "###,###,##0.00", _
                 "Producto", 2, "Descripcion", 1900, 3, 0, "###,###,##0.00", _
                 "P.Und", 2, "nPrecioVenta", 800, 1, 0, "###,###,##0.00", _
                 "Cant.", 2, "nCantidad", 750, 1, 0, "#,##0.000", _
                 "Total", 2, "nVenta", 1000, 1, 0, "###,###,##0.00")
               
    grdNotaCredito.Columns(0).Visible = False
    grdNotaCredito.Columns(1).Visible = False
    grdNotaCredito.Columns(2).Visible = False
   
    Call CargaTotalesNC
   
   'cmdTexto.Caption = "Registro " & IIf(frmNotaCredito.RsCabecera.RecordCount = 0, 0, frmNotaCredito.RsCabecera.AbsolutePosition) & " de " & frmNotaCredito.RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    If sNC = "" Then
        Cn.Execute " DELETE FROM DNOTACREDITO where tNotaCredito =' " & sDocumento & " '"
    End If
    
    Set frmNotaCreditoDetalle = Nothing
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
    cmdOpcion(12).Enabled = Activa
End Sub


'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
                fImpresionDiaContable = obtieneDiaContable
End Function
