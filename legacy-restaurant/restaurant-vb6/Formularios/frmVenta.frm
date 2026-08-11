VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{B52C1CDE-38E9-11D5-98EC-00C0F01C6C81}#1.0#0"; "ifepson.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmVenta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   75
   ClientWidth     =   11970
   ForeColor       =   &H00808080&
   Icon            =   "frmVenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameComboItem 
      Caption         =   "Vista Previa"
      Height          =   5775
      Left            =   3240
      TabIndex        =   390
      Top             =   1440
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton BtnSalir 
         Caption         =   "Ocultar"
         Height          =   375
         Left            =   4320
         TabIndex        =   392
         Top             =   5280
         Width           =   975
      End
      Begin MSComctlLib.ListView lstCombo 
         Height          =   4935
         Left            =   240
         TabIndex        =   391
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   8705
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin MCI.MMControl MMControl1 
      Height          =   615
      Left            =   4920
      TabIndex        =   389
      Top             =   8760
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1085
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame frmPedidoPagado 
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
      Height          =   735
      Left            =   2160
      TabIndex        =   385
      Top             =   0
      Visible         =   0   'False
      Width           =   7755
      Begin VB.CommandButton cmdPedidoPagoEnviar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enviar Pedido"
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   387
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblPedidoPagado 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   3720
         TabIndex        =   388
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Pedido pagado pendiente de envio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   386
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Frame FrmOrdenCompra 
      Height          =   1215
      Left            =   5880
      TabIndex        =   382
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      Begin VB.TextBox txtOrdenCompra 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   384
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdOrdenCompra 
         Caption         =   "Orden de Compra"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   383
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox ImageBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
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
      Height          =   735
      Left            =   240
      ScaleHeight     =   705
      ScaleWidth      =   1305
      TabIndex        =   379
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal PrinterFiscalArgentina 
      Left            =   12000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
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
      Height          =   1710
      Left            =   2760
      TabIndex        =   353
      Top             =   3240
      Visible         =   0   'False
      Width           =   6315
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
         TabIndex        =   356
         Top             =   15
         Width           =   2490
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmVenta.frx":038A
         Top             =   855
         Width           =   240
      End
      Begin VB.Label lblPaso1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enviando información de documento a XXXX."
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
         TabIndex        =   355
         Top             =   870
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.Label lblPaso2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obteniendo codigo XXXX almacenado."
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
         TabIndex        =   354
         Top             =   1155
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmVenta.frx":059D
         Top             =   1155
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmVenta.frx":07B0
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmVenta.frx":0AF2
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "   Proceso de envio de documento a XXXXX......."
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
         TabIndex        =   357
         Top             =   435
         Width           =   5910
      End
   End
   Begin VB.TextBox txtEnvioAntes 
      Height          =   285
      Left            =   1680
      TabIndex        =   346
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraProductoCombo 
      Caption         =   " Productos de Combos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   8610
      Left            =   7800
      TabIndex        =   195
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   2529
         Picture         =   "frmVenta.frx":0E34
         Style           =   1  'Graphical
         TabIndex        =   244
         Top             =   7740
         Width           =   1530
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   48
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   243
         Top             =   7740
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   47
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   242
         Top             =   7740
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   46
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   241
         Top             =   7740
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   45
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   44
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   43
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   238
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   42
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   41
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   6915
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   40
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   39
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   234
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   38
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   233
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   37
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   232
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   36
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   231
         Top             =   6090
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   35
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   230
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   34
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   229
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   33
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   228
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   32
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   227
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   31
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   226
         Top             =   5265
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   30
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   29
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   28
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   27
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   26
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   4440
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   25
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   220
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   24
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   23
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   22
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   21
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   3615
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   2790
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   1965
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   204
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   1140
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   2529
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1731
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   933
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   315
         Width           =   720
      End
      Begin VB.CommandButton cmdProductoCombo 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdOpcion 
      Caption         =   "Ofertas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   17
      Left            =   3330
      Picture         =   "frmVenta.frx":1276
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   7170
      Width           =   1080
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   7
      Left            =   6510
      Picture         =   "frmVenta.frx":13C0
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   6555
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   8
      Left            =   6510
      Picture         =   "frmVenta.frx":14C2
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   7140
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   6
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   7140
      Width           =   615
   End
   Begin VB.Frame fraCombo 
      Caption         =   " Combo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3600
      Left            =   0
      TabIndex        =   125
      Top             =   0
      Visible         =   0   'False
      Width           =   7725
      Begin VB.PictureBox pbxBarra 
         Height          =   375
         Left            =   6360
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   347
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Timer TimerColorWAP 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   7140
         Top             =   765
      End
      Begin MSWinsockLib.Winsock WinSock 
         Left            =   7140
         Top             =   270
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer TimerColor 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   6705
         Top             =   765
      End
      Begin VB.Timer TimerDelivery 
         Interval        =   10000
         Left            =   6705
         Top             =   270
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   8
         Left            =   6345
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   2955
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Orden"
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
         Left            =   6345
         Picture         =   "frmVenta.frx":15C4
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   2400
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Orden"
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
         Left            =   6345
         Picture         =   "frmVenta.frx":16C6
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   1845
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Cantidad"
         Height          =   555
         Index           =   4
         Left            =   5107
         TabIndex        =   186
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "&Disminuir"
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
         Left            =   5107
         Picture         =   "frmVenta.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   1290
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Elimina"
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
         Left            =   5107
         Picture         =   "frmVenta.frx":18CA
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   1845
         Width           =   1230
      End
      Begin VB.CommandButton cmdCombo 
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
         Left            =   5107
         Picture         =   "frmVenta.frx":19CC
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   2955
         Width           =   1230
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   14
         Left            =   4485
         Picture         =   "frmVenta.frx":1ABE
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   1845
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   13
         Left            =   4485
         Picture         =   "frmVenta.frx":2388
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   1290
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   15
         Left            =   4485
         Picture         =   "frmVenta.frx":2C52
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   2955
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   12
         Left            =   4485
         Picture         =   "frmVenta.frx":351C
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "&Aumentar"
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
         Left            =   5107
         Picture         =   "frmVenta.frx":3DE6
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   735
         Width           =   1230
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   16
         Left            =   4485
         Picture         =   "frmVenta.frx":3EE8
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   17
         Left            =   4485
         Picture         =   "frmVenta.frx":47B2
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   735
         Width           =   615
      End
      Begin VB.CommandButton cmdCombo 
         Caption         =   "Propiedad y Observación"
         Height          =   555
         Index           =   5
         Left            =   6345
         TabIndex        =   126
         Top             =   1290
         Width           =   1230
      End
      Begin TrueOleDBGrid80.TDBGrid grdCombo 
         Height          =   3240
         Left            =   90
         TabIndex        =   135
         Top             =   270
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   5715
         _LayoutType     =   4
         _RowHeight      =   26
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
         Splits(0).ScrollBars=   0
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
   End
   Begin VB.CommandButton CmdOpcion 
      Caption         =   "Enviar Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   2220
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   7170
      Width           =   1080
   End
   Begin VB.CommandButton CmdOpcion 
      Caption         =   "Combos"
      Height          =   585
      Index           =   11
      Left            =   1120
      TabIndex        =   111
      Top             =   7170
      Width           =   1080
   End
   Begin VB.CommandButton CmdOpcion 
      Caption         =   "Visualizar Pedido"
      Height          =   585
      Index           =   12
      Left            =   30
      TabIndex        =   110
      Top             =   7170
      Width           =   1080
   End
   Begin VB.Frame fraEmisiones 
      Caption         =   " Emisiones Rápidas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   900
      Left            =   0
      TabIndex        =   106
      Top             =   7800
      Width           =   7800
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Cargos"
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
         Index           =   14
         Left            =   4470
         Picture         =   "frmVenta.frx":507C
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   240
         Width           =   1050
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Cuenta Corritentes"
         Height          =   555
         Index           =   13
         Left            =   5540
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   240
         Width           =   1115
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "&Pre Cuenta"
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
         Left            =   6650
         Picture         =   "frmVenta.frx":517E
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   240
         Width           =   1115
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   3
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   240
         Width           =   1115
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   2
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   240
         Width           =   1115
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   240
         Width           =   1115
      End
      Begin VB.CommandButton cmdTipoDocumento 
         Height          =   555
         Index           =   4
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   240
         Width           =   1115
      End
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Propiedad y Observación"
      Height          =   585
      Index           =   4
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5970
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Disminuir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   6510
      Picture         =   "frmVenta.frx":56B0
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5400
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Aumentar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   2
      Left            =   6510
      Picture         =   "frmVenta.frx":57B2
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Cantidad"
      Height          =   585
      Index           =   1
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4215
      Width           =   1185
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Elimina Prod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   6510
      Picture         =   "frmVenta.frx":58B4
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3630
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "&Observación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   5
      Left            =   6510
      Picture         =   "frmVenta.frx":59B6
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2985
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Transferencia"
      Height          =   585
      Index           =   4
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Camb Mesa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   6510
      Picture         =   "frmVenta.frx":5AF8
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1815
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Elimina Ped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   2
      Left            =   6510
      Picture         =   "frmVenta.frx":5BF2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1230
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "Modifica"
      Height          =   585
      Index           =   1
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   645
      Width           =   1185
   End
   Begin VB.CommandButton cmdCabecera 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   6510
      Picture         =   "frmVenta.frx":5CF4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   60
      Width           =   1185
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   11
      Left            =   5880
      Picture         =   "frmVenta.frx":6226
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6555
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   10
      Left            =   5880
      Picture         =   "frmVenta.frx":6AF0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5970
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   9
      Left            =   5880
      Picture         =   "frmVenta.frx":73BA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5385
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   8
      Left            =   5880
      Picture         =   "frmVenta.frx":7C84
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   7
      Left            =   5880
      Picture         =   "frmVenta.frx":854E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4215
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   6
      Left            =   5880
      Picture         =   "frmVenta.frx":8E18
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3630
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   5
      Left            =   5880
      Picture         =   "frmVenta.frx":96E2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2985
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   4
      Left            =   5880
      Picture         =   "frmVenta.frx":9FAC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   3
      Left            =   5880
      Picture         =   "frmVenta.frx":A876
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1815
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   2
      Left            =   5880
      Picture         =   "frmVenta.frx":B140
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1230
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   1
      Left            =   5880
      Picture         =   "frmVenta.frx":BA0A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   645
      Width           =   615
   End
   Begin VB.CommandButton cmdNavegar 
      Height          =   585
      Index           =   0
      Left            =   5880
      Picture         =   "frmVenta.frx":C2D4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   615
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   315
      Top             =   7245
   End
   Begin VB.TextBox txtNino 
      Height          =   345
      Left            =   1470
      TabIndex        =   5
      Text            =   "txtNiño"
      Top             =   8070
      Width           =   1095
   End
   Begin VB.TextBox txtAdulto 
      Height          =   345
      Left            =   210
      TabIndex        =   4
      Text            =   "txtAdulto"
      Top             =   8070
      Width           =   1095
   End
   Begin TrueOleDBGrid80.TDBGrid grdDetalle 
      Height          =   3525
      Left            =   0
      TabIndex        =   120
      Top             =   3600
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   6218
      _LayoutType     =   4
      _RowHeight      =   26
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "True/False"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).ScrollBars=   0
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
   Begin TrueOleDBGrid80.TDBGrid grdCabecera 
      Height          =   3465
      Left            =   2880
      TabIndex        =   121
      Top             =   90
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   6112
      _LayoutType     =   4
      _RowHeight      =   26
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
      Splits(0).ScrollBars=   0
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
      Caption         =   "Cabecera de Pedidos"
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
   Begin VB.TextBox txtBarra 
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   8160
      Width           =   1785
   End
   Begin VB.TextBox txtJuntar 
      Height          =   345
      Left            =   2790
      TabIndex        =   176
      Text            =   "txtJuntar"
      Top             =   8055
      Width           =   1095
   End
   Begin VB.Frame fraPropiedad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   8610
      Left            =   7800
      TabIndex        =   71
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   3345
         Picture         =   "frmVenta.frx":CB9E
         Style           =   1  'Graphical
         TabIndex        =   250
         Top             =   4950
         Width           =   720
      End
      Begin VB.TextBox lblObservacion 
         Height          =   555
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   181
         Top             =   7155
         Width           =   3975
      End
      Begin VB.TextBox lblResumen 
         Height          =   870
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   180
         Top             =   5985
         Width           =   3975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   5415
         Left            =   1672
         TabIndex        =   179
         Top             =   225
         Width           =   60
      End
      Begin VB.CommandButton cmdBuscar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   870
         Picture         =   "frmVenta.frx":CFE0
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   1440
         Picture         =   "frmVenta.frx":D422
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   7785
         Width           =   1470
      End
      Begin VB.CommandButton CmdOpcion 
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
         Height          =   720
         Index           =   8
         Left            =   2970
         Picture         =   "frmVenta.frx":D564
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   7785
         Width           =   1110
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   2580
         TabIndex        =   104
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   1815
         TabIndex        =   103
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   3345
         TabIndex        =   100
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   2580
         TabIndex        =   99
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   1815
         TabIndex        =   98
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   3345
         TabIndex        =   95
         Top             =   1100
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   2580
         TabIndex        =   94
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   1815
         TabIndex        =   93
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   3345
         TabIndex        =   90
         Top             =   1870
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   2580
         TabIndex        =   89
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   1815
         TabIndex        =   88
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   3345
         TabIndex        =   85
         Top             =   2640
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2580
         TabIndex        =   84
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1815
         TabIndex        =   83
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3345
         TabIndex        =   80
         Top             =   3410
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   2580
         TabIndex        =   79
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   1815
         TabIndex        =   78
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   3345
         TabIndex        =   75
         Top             =   4180
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   2580
         TabIndex        =   74
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdPropiedad 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   1815
         TabIndex        =   73
         Top             =   4950
         Width           =   720
      End
      Begin VB.CommandButton cmdOperador 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   4950
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Observacion :"
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
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   192
         Top             =   6930
         Width           =   1200
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propiedad :"
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
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   191
         Top             =   5760
         Width           =   990
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         Caption         =   "  Propiedad "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   3
         Left            =   1890
         TabIndex        =   183
         Top             =   45
         Width           =   1050
      End
      Begin VB.Label lblPropiedad 
         AutoSize        =   -1  'True
         Caption         =   "  Operador   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   182
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.Frame fraPlato 
      Caption         =   " Platos de Venta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6585
      Left            =   7800
      TabIndex        =   32
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Código Directo"
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
         Index           =   6
         Left            =   120
         TabIndex        =   248
         Top             =   6180
         Width           =   3945
      End
      Begin VB.Frame fraSubGrupo 
         Caption         =   " SubGrupos "
         ForeColor       =   &H000000C0&
         Height          =   3315
         Left            =   60
         TabIndex        =   35
         Top             =   2820
         Width           =   1635
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1020
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1020
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1740
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1740
            Width           =   720
         End
         Begin VB.CommandButton cmdSubGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   2460
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   840
            Picture         =   "frmVenta.frx":D656
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2460
            Width           =   720
         End
      End
      Begin VB.Frame fraGrupo 
         Caption         =   " Grupos "
         ForeColor       =   &H000000C0&
         Height          =   2595
         Left            =   60
         TabIndex        =   33
         Top             =   210
         Width           =   1635
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   990
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   990
            Width           =   720
         End
         Begin VB.CommandButton cmdGrupo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1710
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   0
            Left            =   840
            Picture         =   "frmVenta.frx":DA98
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1710
            Width           =   720
         End
      End
      Begin VB.Frame fraProducto 
         Caption         =   " Productos "
         ForeColor       =   &H000000C0&
         Height          =   5925
         Left            =   1800
         TabIndex        =   34
         Top             =   240
         Width           =   2325
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   90
            TabIndex        =   70
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   810
            TabIndex        =   69
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1530
            TabIndex        =   68
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   90
            TabIndex        =   67
            Top             =   1070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   810
            TabIndex        =   66
            Top             =   1070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   1530
            TabIndex        =   65
            Top             =   1070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   90
            TabIndex        =   64
            Top             =   1870
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   810
            TabIndex        =   63
            Top             =   1870
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   1530
            TabIndex        =   62
            Top             =   1870
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   90
            TabIndex        =   61
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   810
            TabIndex        =   60
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   1530
            TabIndex        =   59
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   90
            TabIndex        =   58
            Top             =   3470
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   810
            TabIndex        =   57
            Top             =   3470
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   1530
            TabIndex        =   56
            Top             =   3470
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   90
            TabIndex        =   55
            Top             =   4270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   810
            TabIndex        =   54
            Top             =   4270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1530
            TabIndex        =   53
            Top             =   4270
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   90
            TabIndex        =   52
            Top             =   5070
            Width           =   720
         End
         Begin VB.CommandButton cmdProducto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   20
            Left            =   810
            TabIndex        =   51
            Top             =   5070
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   1530
            Picture         =   "frmVenta.frx":DEDA
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   5070
            Width           =   720
         End
      End
   End
   Begin VB.Frame fraEliminacion 
      Caption         =   " Motivo de Eliminación "
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
      Height          =   6585
      Left            =   7800
      TabIndex        =   136
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   38
         Left            =   1710
         TabIndex        =   175
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   37
         Left            =   930
         TabIndex        =   174
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   36
         Left            =   150
         TabIndex        =   173
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   35
         Left            =   3300
         TabIndex        =   172
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   31
         Left            =   150
         TabIndex        =   171
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   32
         Left            =   930
         TabIndex        =   170
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   33
         Left            =   1710
         TabIndex        =   169
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   34
         Left            =   2505
         TabIndex        =   168
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   30
         Left            =   3300
         TabIndex        =   167
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   26
         Left            =   150
         TabIndex        =   166
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   27
         Left            =   930
         TabIndex        =   165
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   28
         Left            =   1725
         TabIndex        =   164
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   29
         Left            =   2505
         TabIndex        =   163
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   25
         Left            =   3300
         TabIndex        =   162
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   21
         Left            =   150
         TabIndex        =   161
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   22
         Left            =   930
         TabIndex        =   160
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   23
         Left            =   1725
         TabIndex        =   159
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   24
         Left            =   2505
         TabIndex        =   158
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   3300
         TabIndex        =   157
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   150
         TabIndex        =   156
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   930
         TabIndex        =   155
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   1725
         TabIndex        =   154
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   2505
         TabIndex        =   153
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3300
         TabIndex        =   152
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   150
         TabIndex        =   151
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   930
         TabIndex        =   150
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1725
         TabIndex        =   149
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2505
         TabIndex        =   148
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   3300
         TabIndex        =   147
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   150
         TabIndex        =   146
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   930
         TabIndex        =   145
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1725
         TabIndex        =   144
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   2505
         TabIndex        =   143
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   3300
         TabIndex        =   142
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   150
         TabIndex        =   141
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   930
         TabIndex        =   140
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   1725
         TabIndex        =   139
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   2505
         TabIndex        =   138
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton CmdOpcion 
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
         Height          =   720
         Index           =   16
         Left            =   2505
         Picture         =   "frmVenta.frx":E31C
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   5625
         Width           =   1515
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   " Opcion "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   7800
      TabIndex        =   112
      Top             =   6600
      Width           =   4155
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Refrescar"
         Height          =   585
         Index           =   10
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   247
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelivery 
         BackColor       =   &H000000FF&
         Caption         =   "Pedidos de la Central Delivery"
         Height          =   585
         Left            =   1470
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdMesa247 
         BackColor       =   &H0000FF00&
         Caption         =   "Pedidos de Mesa247"
         Height          =   585
         Left            =   1470
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   349
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Caja Rápida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   15
         Left            =   2790
         Picture         =   "frmVenta.frx":E8A6
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   330
         Width           =   1230
      End
      Begin VB.CommandButton cmdWebapp 
         BackColor       =   &H000000FF&
         Caption         =   "Pedidos Web Mobile"
         Height          =   585
         Left            =   2760
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   345
         Top             =   330
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Pagos y Div.Docum"
         Height          =   585
         Index           =   5
         Left            =   1470
         TabIndex        =   246
         Top             =   330
         Width           =   1230
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Transferencia Pedidos"
         Height          =   585
         Index           =   7
         Left            =   120
         TabIndex        =   113
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton CmdOpcion 
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
         Height          =   585
         Index           =   0
         Left            =   2790
         Picture         =   "frmVenta.frx":E9F0
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Busca Pedido"
         Height          =   585
         Index           =   9
         Left            =   135
         TabIndex        =   114
         Top             =   330
         Width           =   1230
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Cancelar"
         Height          =   585
         Index           =   2
         Left            =   1470
         TabIndex        =   116
         Top             =   1020
         Width           =   1230
      End
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "&Aceptar"
         Height          =   585
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   1020
         Width           =   1230
      End
   End
   Begin VB.Frame fraPedido 
      Caption         =   " Pedido "
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
      Height          =   3555
      Left            =   0
      TabIndex        =   251
      Top             =   0
      Width           =   2805
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto Max:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   10
         Left            =   120
         TabIndex        =   352
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label txtMontoMax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   351
         Top             =   3180
         Width           =   1800
      End
      Begin VB.Label txtMozo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   275
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mesero :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   90
         TabIndex        =   274
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label txtDireccion 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   915
         TabIndex        =   273
         Top             =   2310
         Width           =   1800
      End
      Begin VB.Label txtCliente 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   915
         TabIndex        =   272
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label txtMotorizado 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   271
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label txtTelefono 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   270
         Top             =   2070
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   255
         TabIndex        =   269
         Top             =   2080
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   225
         TabIndex        =   268
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observac. :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   6
         Left            =   195
         TabIndex        =   267
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   390
         TabIndex        =   266
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Motorizado :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   7
         Left            =   90
         TabIndex        =   265
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   435
         TabIndex        =   264
         Top             =   263
         Width           =   450
      End
      Begin VB.Label txtObservacion 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   263
         Top             =   2700
         Width           =   1800
      End
      Begin VB.Label txtDescuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1305
         TabIndex        =   262
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label txtFecha 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   261
         Top             =   225
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   135
         TabIndex        =   260
         Top             =   960
         Width           =   750
      End
      Begin VB.Label txtPuntoVenta 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   259
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   25
         Left            =   195
         TabIndex        =   258
         Top             =   720
         Width           =   690
      End
      Begin VB.Label txtFechaProg 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   257
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Pedido :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   60
         TabIndex        =   256
         Top             =   480
         Width           =   825
      End
      Begin VB.Label txtTipoPedido 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   255
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label lblMoneda 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   990
         TabIndex        =   254
         Top             =   960
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entregar A :"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   9
         Left            =   165
         TabIndex        =   253
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label txtEntregar 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   915
         TabIndex        =   252
         Top             =   2940
         Width           =   1800
      End
   End
   Begin VB.Frame fraCabecera 
      Caption         =   " Opciones de Pedido "
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
      Height          =   6630
      Left            =   7800
      TabIndex        =   276
      Top             =   0
      Width           =   4155
      Begin VB.Frame fraOrigenVentas 
         Caption         =   "Origen de Ventas"
         ForeColor       =   &H00800000&
         Height          =   3405
         Left            =   90
         TabIndex        =   358
         Top             =   1230
         Width           =   3975
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   3150
            Picture         =   "frmVenta.frx":EAE2
            Style           =   1  'Graphical
            TabIndex        =   378
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   377
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   870
            TabIndex        =   376
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1635
            TabIndex        =   375
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2385
            TabIndex        =   374
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3150
            TabIndex        =   373
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   120
            TabIndex        =   372
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   870
            TabIndex        =   371
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1635
            TabIndex        =   370
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   2385
            TabIndex        =   369
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   3150
            TabIndex        =   368
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   120
            TabIndex        =   367
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   870
            TabIndex        =   366
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   1635
            TabIndex        =   365
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   2385
            TabIndex        =   364
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   3150
            TabIndex        =   363
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   120
            TabIndex        =   362
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   870
            TabIndex        =   361
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1635
            TabIndex        =   360
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdOrigen 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   2385
            TabIndex        =   359
            Top             =   2550
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Monto Max/P : 0"
         Height          =   555
         Index           =   21
         Left            =   2820
         TabIndex        =   350
         Top             =   6045
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Aplica Anticipo"
         Height          =   545
         Index           =   20
         Left            =   210
         TabIndex        =   348
         Top             =   6045
         Width           =   1155
      End
      Begin VB.Frame fraPuntoVenta 
         Caption         =   " Punto de Venta "
         ForeColor       =   &H00800000&
         Height          =   1830
         Left            =   90
         TabIndex        =   281
         Top             =   1230
         Width           =   3975
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   291
            Top             =   315
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   870
            TabIndex        =   290
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1635
            TabIndex        =   289
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2385
            TabIndex        =   288
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3150
            TabIndex        =   287
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   120
            TabIndex        =   286
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   870
            TabIndex        =   285
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1635
            TabIndex        =   284
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdPunto 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   2385
            TabIndex        =   283
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3150
            Picture         =   "frmVenta.frx":EF24
            Style           =   1  'Graphical
            TabIndex        =   282
            Top             =   1050
            Width           =   720
         End
      End
      Begin VB.Frame fraMozo 
         Caption         =   " Mesero "
         ForeColor       =   &H00800000&
         Height          =   3405
         Left            =   90
         TabIndex        =   292
         Top             =   1230
         Width           =   3975
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   312
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   870
            TabIndex        =   311
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1635
            TabIndex        =   310
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2385
            TabIndex        =   309
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3150
            TabIndex        =   308
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   120
            TabIndex        =   307
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   870
            TabIndex        =   306
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1635
            TabIndex        =   305
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   2385
            TabIndex        =   304
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   3150
            TabIndex        =   303
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   120
            TabIndex        =   302
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   870
            TabIndex        =   301
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   1635
            TabIndex        =   300
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   2385
            TabIndex        =   299
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   3150
            TabIndex        =   298
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   120
            TabIndex        =   297
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   870
            TabIndex        =   296
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1635
            TabIndex        =   295
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMozo 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   2385
            TabIndex        =   294
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   3150
            Picture         =   "frmVenta.frx":F366
            Style           =   1  'Graphical
            TabIndex        =   293
            Top             =   2550
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Mesa"
         Height          =   555
         Index           =   13
         Left            =   1515
         TabIndex        =   317
         Top             =   5480
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Cliente"
         Height          =   555
         Index           =   12
         Left            =   210
         TabIndex        =   316
         Top             =   5480
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Observación"
         Height          =   555
         Index           =   11
         Left            =   2820
         TabIndex        =   315
         Top             =   4900
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Asignación"
         Height          =   555
         Index           =   14
         Left            =   1515
         TabIndex        =   280
         Top             =   4900
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Descuento"
         Height          =   555
         Index           =   15
         Left            =   2820
         TabIndex        =   279
         Top             =   5480
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Contacto"
         Height          =   555
         Index           =   16
         Left            =   210
         TabIndex        =   278
         Top             =   4900
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "&Entregar A"
         Height          =   545
         Index           =   19
         Left            =   1515
         TabIndex        =   277
         Top             =   6045
         Width           =   1155
      End
      Begin VB.Frame fraMotorizado 
         Caption         =   " Motorizado "
         ForeColor       =   &H00800000&
         Height          =   3405
         Left            =   90
         TabIndex        =   318
         Top             =   1230
         Width           =   3975
         Begin VB.CommandButton cmdBuscar 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   3120
            Picture         =   "frmVenta.frx":F7A8
            Style           =   1  'Graphical
            TabIndex        =   338
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   19
            Left            =   2370
            TabIndex        =   337
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   18
            Left            =   1620
            TabIndex        =   336
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   17
            Left            =   870
            TabIndex        =   335
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   16
            Left            =   120
            TabIndex        =   334
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   15
            Left            =   3120
            TabIndex        =   333
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   14
            Left            =   2370
            TabIndex        =   332
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   13
            Left            =   1620
            TabIndex        =   331
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   12
            Left            =   870
            TabIndex        =   330
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   11
            Left            =   120
            TabIndex        =   329
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   10
            Left            =   3120
            TabIndex        =   328
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   9
            Left            =   2370
            TabIndex        =   327
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   8
            Left            =   1620
            TabIndex        =   326
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   7
            Left            =   870
            TabIndex        =   325
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   6
            Left            =   120
            TabIndex        =   324
            Top             =   1050
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   3120
            TabIndex        =   323
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   4
            Left            =   2370
            TabIndex        =   322
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   3
            Left            =   1620
            TabIndex        =   321
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   2
            Left            =   870
            TabIndex        =   320
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdMotorizado 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   1
            Left            =   120
            TabIndex        =   319
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame fraTipoPedido 
         Caption         =   " Tipo de Pedido "
         ForeColor       =   &H00800000&
         Height          =   1065
         Left            =   90
         TabIndex        =   339
         Top             =   225
         Width           =   3975
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Delivery"
            Height          =   720
            Index           =   8
            Left            =   1627
            TabIndex        =   344
            Top             =   225
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "En el &Local"
            Height          =   720
            Index           =   7
            Left            =   866
            TabIndex        =   343
            Top             =   225
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Para Llevar"
            Height          =   720
            Index           =   6
            Left            =   105
            TabIndex        =   342
            Top             =   225
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Delivery"
            Height          =   720
            Index           =   17
            Left            =   2388
            TabIndex        =   341
            Top             =   225
            Width           =   720
         End
         Begin VB.CommandButton cmdCabecera 
            Caption         =   "&Delivery"
            Height          =   720
            Index           =   18
            Left            =   3150
            TabIndex        =   340
            Top             =   225
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Alta Prioridad"
         Height          =   555
         Index           =   9
         Left            =   210
         TabIndex        =   313
         Top             =   4900
         Width           =   1155
      End
      Begin VB.CommandButton cmdCabecera 
         Caption         =   "Fecha Entrega"
         Height          =   555
         Index           =   10
         Left            =   1515
         Style           =   1  'Graphical
         TabIndex        =   314
         Top             =   4900
         Width           =   1155
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer bocina 
      Height          =   375
      Left            =   10320
      TabIndex        =   393
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   661
   End
   Begin VB.Label lblestadomotor 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Motor Inactivo"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   7800
      TabIndex        =   381
      Top             =   8490
      Width           =   4065
   End
   Begin VB.Image ImageMercadoPago 
      Height          =   615
      Left            =   5040
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image ImageIntegracion 
      Height          =   615
      Left            =   120
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLey1 
      Alignment       =   1  'Right Justify
      Caption         =   "ICBPER 0.10"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Left            =   4920
      TabIndex        =   380
      Top             =   7665
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imageHash 
      Height          =   615
      Left            =   1680
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imageCab 
      Height          =   135
      Left            =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imagepIE 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "U.Fact :"
      Height          =   195
      Index           =   23
      Left            =   7740
      TabIndex        =   31
      Top             =   8240
      Width           =   570
   End
   Begin VB.Label lblFactura 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   8340
      TabIndex        =   30
      Top             =   8240
      Width           =   1095
   End
   Begin VB.Label lblFecha 
      Alignment       =   1  'Right Justify
      Caption         =   "10/01/1970"
      Height          =   255
      Left            =   9480
      TabIndex        =   18
      Top             =   8240
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Pedido"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   24
      Left            =   4845
      TabIndex        =   3
      Top             =   7170
      Width           =   1005
   End
   Begin VB.Label txtSuma 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4485
      TabIndex        =   2
      Top             =   7335
      Width           =   1350
   End
End
Attribute VB_Name = "frmVenta"
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
'validacionMontoMAXIMO
Dim nMontoPedidoFacturar As Double
'VALIDACIONMONTOMAXIMO
'validacionMontoMInimo
Dim nMontoPedidoFacturarMInimo As Double
'validacionMontoMInimo
Dim xOperador As String
Dim wAgregarPropiedad As Boolean
Dim numeroSerieImpresora As String 'SUNAT
Dim codigoImpresora As String 'SUNAT
Dim rstFuente As New ADODB.Recordset
Public RsCabecera As ADODB.Recordset
Public RsDetalle As ADODB.Recordset
Dim RsCombo As Recordset
Dim RsImpresion As Recordset
Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset
Dim RsProducto As Recordset
Dim RsMozo As Recordset
Dim RsMotorizado As Recordset
Dim RsListaAsignada As Recordset
Dim RsverCD As Recordset

'Validar Descuento
Dim ValidaDesc As Boolean

Dim xSinPrepago As String

Dim rsItemCombo   As New Recordset
     
     
'origen de ventas
Dim RsOrigenVentas As Recordset
Dim RscanalOrigenVentas As Recordset
Dim vOrigenVentas As String
'-----------------------------
Dim RsArea As Recordset
Dim RsPropiedad As Recordset
Dim RsComboPropiedad As Recordset
Dim RsProductoPropiedad As Recordset
Dim RsTipoDocumento As Recordset
Dim RsTimer As Recordset
Dim RsT3 As Recordset
Dim RsPC As Recordset
Dim rsPuntoVenta As Recordset
Dim RsOferta As Recordset
Dim RsPedido As Recordset
Dim RsMotivoEliminacion As Recordset
Dim RsOperador As Recordset
Dim RsProductoCombo As Recordset
Dim RsClienteSelecc As Recordset
Dim RsRecord As Recordset
Dim RsCanalesVenta As Recordset
Dim RsTemporal As Recordset
Dim RsCodigoPCD As ADODB.Recordset
Dim ValProductoAsociado As ADODB.Recordset
Dim RsCodigoTar As Recordset
Dim RsExisteOferta As Recordset
Dim RsPedidosCarta As Recordset

'Variables para oferta
Dim ValorOferta As Recordset
Dim respuestaOferta As String

'Variables Generales
Dim wDetalle As Boolean
Dim wAgrega As Boolean
Dim tCodigoPedidoCD As String

'Variables Combo
Dim wCombo As Boolean
Dim wAgregaCombo As Boolean
Dim nCombo As Integer

'Variables Cabecera
Dim sTipoAtencion As String
Dim sMozoD As String
Dim sMotorizado As String
Dim sCortesia As String
'Dim ntTiempo As Integer
Dim sDetalleConsumo As String
Public lIncluido As Boolean



Dim sPuntoVenta As String
Dim sCompania As String
Dim sContacto As String
Dim sComandaInfhotel As String
Dim wCabecera As Boolean
Dim lTipoDocumento1 As Boolean
Dim lTipoDocumento2 As Boolean
Dim lTipoDocumento3 As Boolean
Dim lTipoDocumento4 As Boolean
'entregarA
Dim lObligaEntregarA As Boolean
Dim lObligaOrigenVenta As Boolean
'Variables Detalles
Dim nPVenta As Double
Dim nPBase As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nRecargo As Double
Dim nDescuento As Double
Dim nOficial As Double
Dim nCantidad As Double
Dim sItem As String
Dim xItem As String
Dim sProducto As String
Dim sProductoCombo As String
Dim sCombo As String
Dim sGrupo As String
Dim xGrupo As String
Dim sSubGrupo As String
Dim xSubGrupo As String
Dim sTD As String
Dim xSuma As Double
Dim xLinea As Double
Dim xConsumo As Double
Dim xDescuento As Double
Dim xMontoMaximo As Double '---- monto maximo erick
Dim lPrecuenta As Boolean
Dim lPropiedad As Boolean
Dim lComboActivo As Boolean
Dim nCCombo As Double
Dim sMonedaBase As String
Dim sDescuento As Double
Dim Acumulado As Double
Dim sUsuarioAutoriza As String
Dim sCodigoDescuento As String
Dim sDescripcionDescuento As String
Dim tAutorizaDescuento As String
Dim wCalcula As Boolean
Dim UltimaComanda As String
Dim PrecioTotal As Double
Dim EstadoC As String
Dim nOrden As Integer
Dim nOperadorPropiedad As Integer
Dim lAplicablePedido As Boolean
Dim sAutoriza As String

'===================== tope
Dim lRatio As Boolean
Dim ltope As Boolean
Dim nTope As Double
Dim codigoanteriordescuento As String
Dim montoanteriorDescuento As Double
Dim lImprimeAlternativa As Boolean

'============================================= extranjero bolivia
Dim tAutorizacion As String
Dim tCodigoControl As String
Dim tDosificacion As String
Dim tIdentidadNIT As String

'insumo critico23
Dim muestra As String

'Canales de Venta
Dim lActivaMozo As Boolean
Dim lActivaMotorizado As Boolean
Dim lObligaMesa As Boolean
'Origen de ventas
Dim lOrigenVentas As Boolean
'Public lObligaPax As Boolean
Dim lObligaMotorizado As Boolean
Dim lObligaMozo As Boolean
Dim lObligaFechaEntrega As Boolean
Dim lObligaClienteFrecuente As Boolean
Dim lCanalDelivery As Boolean
Dim lCanalCentralPedidos As Boolean
Dim Tienda As String
Dim sXML As String

'------VALIDA CORREO----------
Dim sTipoDocum As String
Dim lValidaEmail As Boolean
Dim sEmail As String

'FACTURACION_E_PERU
'Dim RsImpDocumentoE As New Recordset
Dim RsCodigoHash As New ADODB.Recordset
Dim fDocumento As String
Dim lcodigoHash As Boolean
Dim xMontoTexto As String
Dim iImagenCab As Boolean
Dim xImpresionFE As String
Dim xImpresioDE As String

Dim xTiporecepcion As Integer
Dim xttipopedido As String
Dim cadenaCodigoHash As String

'INTEGRACION_E_OFISIS
Dim oComandoCabeceraOfisis As clsComando
Dim oComandoDetalleOfisis As clsComando
Dim oComandoCabeceraOfisis1 As clsComando
Dim oComandoDetalleOfisis1 As clsComando
Dim oComandoFirmaDocumentoOfisis As clsComando
Dim oComandoFirmaDocumentoOfisis1 As clsComando

Dim rdi As Integer
'Public sTipoDocumVENTA As String
Dim TimpresionDolaresDelivery As Boolean

'Caja Contigencia
Dim Rscajacontigencia As Recordset
Dim scajacontingencia As Boolean
Dim sfemision As Date

'Verificacion de Combo
Dim iProductoVerificaComboMaximo As Integer
Dim iProductoVerificaComboMinimo As Integer
Dim iProductoVerificaComboUsado As Integer
Dim sProductoVerificaComboUsado As String

Dim idPedidoMesa As String
Dim sValidaIntegracion As Boolean

Dim nPropiedadVisible As Integer

Sub comandaAuto()
     
    'Obteniendo pedidos por imprimir automatico
    ' 20012021
    Dim xSql As String
    Dim i As Integer
    Dim xPedido As String
    xPedido = ""
    Dim xUsuarioAutorizaenvio, xPerfilUsusario As String
    xUsuarioAutorizaenvio = "GIAN"
    xPerfilUsusario = "01"
     xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                         "nInsumo, nGasto, nManoObra, toperador.tDetallado AS Operador " & _
                         "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                         "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
                         
    Isql = "select top 1 tCodigoPedido , tSalon from MPEDIDO WHERE tTiporecepcion=3 and testadopedido<>'03'  and isnull(nReimpresion,0) = 0 and ttipopedido = '01' order by fregistro desc "
    Set RsPedidosCarta = Lib.OpenRecordset(Isql, Cn)
     

     
    If Not RsPedidosCarta.EOF Then
    
        xPedido = RsPedidosCarta!tCodigoPedido
        Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
         
        'Producto Propiedad
        Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, dbo.TPRODUCTOPROPIEDAD.nCantidad, t1.Descripcion AS Descripcion, Operador " & _
               "FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo and dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto and dbo.TPRODUCTOPROPIEDAD.tenlace= t1.tenlace " & _
               "where tCodigoPedido = '" & xPedido & "'"
        Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
           
        'Combo Propiedad
        Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador " & _
               "FROM dbo.TCOMBOPROPIEDAD INNER JOIN (" & xSql & ") T1 ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace " & _
               "where tCodigoPedido = '" & xPedido & "'"
         Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
        
        Isql = "select * from vPedido " & _
           "Where Codigo = '" & xPedido & "' And lImprimeArea = 1 and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 )) " & _
           "ORDER BY nOrden, tItem, nOrdenCombo, tetiqueta,combo"
           '"ORDER BY nOrden, tItem, nOrdenCombo, tItemCombo"
                           
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                
        If Not RsImpresion.EOF Then
           RsArea.MoveFirst
           For i = 1 To RsArea.RecordCount
               RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
               If RsArea!tIcono = "" Or RsPedidosCarta!tSalon = RsArea!tIcono Or ((IsNull(RsPedidosCarta!tSalon) Or RsPedidosCarta!tSalon = "") And RsArea!nValor = 1) Then
                  If RsImpresion.RecordCount <> 0 Then
                     RsImpresion.MoveFirst
                     ImprimePedidoAuto RsImpresion, "N", RsArea!timpresora, RsArea!Area, True, RsProductoPropiedad, RsComboPropiedad, "Inforest", xPedido
                  End If
               End If
               RsArea.MoveNext
           Next i
           
        Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1  where tCodigoPedido='" & xPedido & "'"
        Cn.Execute "Update DPEDIDO Set lImprime = 1, tUsuarioAutorizaEnvio = '" & xUsuarioAutorizaenvio & "', tPerfilUsuarioAutoriza = '" & xPerfilUsusario & "', fEnvio = getdate(), nEnvio = 1 where tCodigoPedido = '" & xPedido & "' and lImprime<>1"
        Cn.Execute "Update CPEDIDO Set lImprime = 1 where tCodigoPedido = '" & xPedido & "'"
        
        End If
        
     
    
    End If
End Sub

Private Sub ImprimeFormatoA()
                    Dim Reporte As New dsrBoleta

                    If RsTipoDocumento!lImprimeImageCab Then
                       iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                    End If
                
                    Reporte.DiscardSavedData
                    Reporte.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                    
                    If xImpresioDE = "B" Then
                       Reporte.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                    ElseIf xImpresioDE = "F" Then
                       Reporte.Text13.SetText "FACTURA ELECTRONICA"
                    End If
                    
                    Reporte.Text8.SetText sRazonSocial
                    Reporte.ReportTitle = sDireccion
                    Reporte.Text15.SetText sTelefono
                    Reporte.Text14.SetText sFax
                    Reporte.Text16.SetText sRUC
                    Reporte.Text50.SetText sWeb
                    
                    If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn) = 1 Then
                    Reporte.ReportComments = tTextoAgenteRetencion
                    End If
                    
                    xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                    Reporte.Text4.SetText xMontoTexto
                    Reporte.Text31.SetText xImpresionFE

                    If lFacturacionE And lFECarbajal Then
                        If lQRFE = False And lImpresionCodigoBarras = False Then
                            'Reporte.ImagenHash.Suppress
                            Reporte.CodigoHash.SetText cadenaCodigoHash
                        Else
                            'Reporte.CodigoHash.Suppress
                        End If
                    End If
'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = Reporte
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal

                    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Reporte.PaperOrientation = crPortrait
                    Reporte.PrintOut False, 1, False, 1, 1
                    '----------------
                    
                    If iImagenCab Then
                       Kill App.path & "\cliente.jpg"
                    End If
End Sub


Private Sub ImprimeFormatoAConsumo()

                        Dim ReporteC As New dsrBoletaC
                        
                        If RsTipoDocumento!lImprimeImageCab Then
                           iImagenCab = Generar_Imagen(Cn, "select iImagenCabDoc As imagen from TCAJA where tCaja='" & sCaja & "'", "imagen", "\cliente.jpg")
                        End If
                        
                        ReporteC.DiscardSavedData
                        ReporteC.Database.SetDataSource RsImpresion 'RsImpDocumentoE
                                                            
                        If xImpresioDE = "B" Then
                           ReporteC.Text13.SetText "BOLETA DE VENTA ELECTRONICA"
                        ElseIf xImpresioDE = "F" Then
                           ReporteC.Text13.SetText "FACTURA ELECTRONICA"
                        End If
                        
                        ReporteC.Text8.SetText sRazonSocial
                        ReporteC.ReportTitle = sDireccion
                        ReporteC.Text15.SetText sTelefono
                        ReporteC.Text33.SetText sFax
                        ReporteC.Text16.SetText sRUC
                        ReporteC.Text50.SetText sWeb
                        
                        ReporteC.Text31.SetText sDetalleConsumo
                        
                        If Calcular(" SELECT case when  lImpresionRetencion=1 then 1 else 0 end  as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn) = 1 Then
                           ReporteC.ReportComments = tTextoAgenteRetencion
                        End If
                        
                        xMontoTexto = "SON: " & NumeroCadena(str(RsImpresion!nVenta)) & " " & sMonedaN
                        ReporteC.Text4.SetText xMontoTexto
                        ReporteC.Text32.SetText xImpresionFE

                        If lFacturacionE And lFECarbajal Then
                            If lQRFE = False And lImpresionCodigoBarras = False Then
                                'ReporteC.ImagenHash.Suppress
                                ReporteC.CodigoHash.SetText cadenaCodigoHash
                            Else
                                'ReporteC.CodigoHash.Suppress
                            End If
                        End If

'                        frmEmite.CRViewer.DisplayGroupTree = False
'                        frmEmite.CRViewer.ReportSource = ReporteC
'                        frmEmite.CRViewer.ViewReport
'                        frmEmite.Show vbModal
                        
                        ReporteC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ReporteC.PaperOrientation = crPortrait
                        ReporteC.PrintOut False, 1, False, 1, 1
                        '----------------
                                                          
                    
                        If iImagenCab Then
                           Kill App.path & "\cliente.jpg"
                        End If
End Sub

'----------------------------
'-----------------------------
Public Sub InsertaProductoRotulado(codigoProducto As String, xCantidad As Double, codigoEtiqueta As String)
    'CESAR ROTULADO
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    'ORDEN
    Dim RsOrd As Recordset
    Dim nOrden As Integer
    'DETALLE
    Dim lProductoMultiArea As Boolean
    Dim tsubalmacen As String
    Dim tAreaProduccion As String
    'OFERTA
    Dim tOferta As String
    Dim nOferta As Double
    tOferta = ""
    nOferta = 0
    
    RsProducto.Filter = adFilterNone
    RsProducto.MoveFirst
    RsProducto.Find "Codigo = '" & Trim(codigoProducto) & "'"
    
    
    If Not RsProducto.EOF Then

        nRecargo = 0
        nDescuento = 0
        nValor = 0
        nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
        nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
        nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)

        lImp1 = RsProducto!lImpuesto1
        lImp2 = RsProducto!lImpuesto2
        lImp3 = RsProducto!lImpuesto3

        If sTipoPedido = "02" Then
           If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto4
              lImp2 = RsProducto!lImpuesto5
              lImp3 = RsProducto!lImpuesto6
           End If
        ElseIf sTipoPedido = "03" Then
           If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto7
              lImp2 = RsProducto!lImpuesto8
              lImp3 = RsProducto!lImpuesto9
           End If
        ElseIf sTipoPedido = "04" Then
           If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto10
              lImp2 = RsProducto!lImpuesto11
              lImp3 = RsProducto!lImpuesto12
           End If
        ElseIf sTipoPedido = "05" Then
           If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
              nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
           Else
              nOficial = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
              nValor = 0
              nValor = nValor + IIf(RsProducto!lImpuesto13, nPorcentaje1, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto14, nPorcentaje2, 0)
              nValor = nValor + IIf(RsProducto!lImpuesto15, nPorcentaje3, 0)
              lImp1 = RsProducto!lImpuesto13
              lImp2 = RsProducto!lImpuesto14
              lImp3 = RsProducto!lImpuesto15
           End If
        Else
           nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
        End If
        
        'PRECIO OFICIAL
        nOficial = IIf(RsProducto!tMONEDA = "02", nOficial * nTC, nOficial)
        
        'DESCUENTO
        If xDescuento <> 0 And RsProducto!lDescuento Then
              nPVenta = nOficial - (nOficial * xDescuento / 100)
              nDescuento = nOficial - nPVenta
        Else
              nPVenta = nOficial - nOferta
              nDescuento = nOficial - nPVenta
        End If


        Select Case pais
            Case "001" 'Bolivia
                    nValor = (nValor / 100)
                    nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3

            Case Else 'Peru, Ecuador
                    nValor = 1 + (nValor / 100)
                    nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        End Select

        Dim nInsumo As Double
        Dim nGasto As Double
        Dim nMObra As Double
        
            If sTipoPedido = "01" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
               nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
               nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
            ElseIf sTipoPedido = "02" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
               nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
               nMObra = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
            ElseIf sTipoPedido = "03" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
               nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
               nMObra = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
            ElseIf sTipoPedido = "04" Then
               nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
               nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
               nMObra = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
            Else
               nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
               nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
               nMObra = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
            End If
            
            sItem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
            'CALCULAR ITEM
            If RsDetalle.RecordCount = 0 Then
               'sitem = "001"
               nOrden = 1
            Else
               'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
               If lOrden Then
                  Set RsOrd = Lib.OpenRecordset("select nOrden, lImprime from DPEDIDO where tCodigoPedido='" & sPedido & "' Order by nOrden DESC", Cn)
                  If RsOrd.RecordCount > 0 Then
                     If IIf(IsNull(RsOrd!lImprime), False, RsOrd!lImprime) Then
                        nOrden = RsOrd!nOrden + 1
                     Else
                        nOrden = RsOrd!nOrden
                     End If
                  Else
                     nOrden = 1
                  End If
               Else
                  nOrden = RsProducto!nOrden
               End If
            End If
            
            
            'TSUBALMACEN
            lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    
            If lProductoMultiArea = False Then
                tsubalmacen = ""
            Else
                tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
                
                tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
                
                If tsubalmacen = "0" Then
                    tsubalmacen = ""
                End If
            
            End If
            
            If lBloqInafecto And Calcular("select count(*) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn) > 0 And Round(nPVenta, 2) > 0 Then
                Dim IgvPED As Double
                IgvPED = CDbl(Calcular("select top 1 round(nprecioimpuesto1,2) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn))
                If IgvPED > 0 Then
                    If Round(nImpuesto1, 2) <= 0 Then
                        MsgBox ("No es posible ingresar el producto, no se puede ingresar Items inafecto y afecto en un mismo pedido!! "), vbInformation
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                Else
                    If Round(nImpuesto1, 2) > 0 Then
                        MsgBox ("No es posible ingresar el producto, no se puede ingresar Items inafecto y afecto en un mismo pedido!! "), vbInformation
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If
            
            'INSERT DPEDIDO
            ' cambios GCAA 24092020
               Isql = "insert into DPEDIDO " & _
                      "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                      "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2," & _
                      "nPrecioImpuesto3, nPrecioVenta, nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                      "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, tEstadoItem, tComanda, fRegistro, " & _
                      "tMozoD, tUsuarioD, nInsumo, nGasto, nManoObra, nOrden, tUnidadNegocio, tOferta, tsubalmacen, tCodigoEtiqueta, fdiacontable, tCajaD,tAplicaExoneracion,tAplicaInafectacion  ) " & _
                      "Values(   '" & sPedido & "', " _
                              & "'" & sTipoPedido & "', '" & sItem & "', '" & codigoProducto & "', " _
                              & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                              & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                              & nPBase & ", " & nRecargo & ", " & nDescuento & ", " & nOficial & ", " _
                              & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                              & nPVenta & ", " & xCantidad & ", " & nPVenta * xCantidad & ", " _
                              & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                              & "0, '" & RsProducto!tArea & "', " _
                              & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                              & IIf(RsProducto!lCombinacion, -1, 0) & "," _
                              & IIf(IsNull(RsProducto!nCombinacion), 0, RsProducto!nCombinacion) & "," _
                              & "'N', '" & sComanda & "', getDate(), '" & sMozoD & "', '" & sUsuario & "', " _
                              & nInsumo & ", " _
                              & nGasto & ", " _
                              & nMObra & ", " _
                              & nOrden & ", '" & sUnidadNegocio & "','" & tOferta & "','" & tsubalmacen & "','" & codigoEtiqueta & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "','" & sCaja & "', " _
                              & IIf(RsProducto!tAplicaExoneracion, -1, 0) & " , " _
                              & IIf(RsProducto!tAplicaInafectacion, -1, 0) & ")"
               Cn.Execute Isql
            
               RsDetalle.Requery


    Else
       MsgBox "Producto no encontrado", vbCritical, sMensaje
    End If

End Sub



Public Function verificaCantidadDeItemsCombos(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal numeroDeItemCombos As Double, ByVal ncantidadNueva As Double) As Boolean
    Dim cantidadMaximoNueva As Double
    Dim cantidadMaximoPosible As Double
    Dim X As Integer
    Dim oRstRecorriendoCombo As New Recordset
    Dim oRstCantidadNUnicos As New Recordset
    
    verificaCantidadDeItemsCombos = False
    cantidadMaximoNueva = numeroDeItemCombos * ncantidadNueva
    cantidadMaximoPosible = 0
    Set oRstRecorriendoCombo = Lib.OpenRecordset("select tcombo.ncantidad ,cpedido.tproductocombo from cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tcodigopedido='" & tCodigoPedido & "' and cpedido.titem='" & tItem & "' and lfijo=1", Cn)
    If Not (oRstRecorriendoCombo.EOF Or oRstRecorriendoCombo.BOF) Then
        oRstRecorriendoCombo.MoveFirst
        For X = 0 To oRstRecorriendoCombo.RecordCount - 1
            cantidadMaximoPosible = cantidadMaximoPosible + (oRstRecorriendoCombo!nCantidad * ncantidadNueva)
            oRstRecorriendoCombo.MoveNext
        Next X
    End If
    
    Set oRstCantidadNUnicos = Lib.OpenRecordset("select isnull(sum(cpedido.ncantidad),0) from cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where tcodigopedido='" & tCodigoPedido & "' and titem='" & tItem & "' and tcombo.lfijo=0 ", Cn)
    If Not (oRstCantidadNUnicos.EOF Or oRstCantidadNUnicos.BOF) Then
        cantidadMaximoPosible = cantidadMaximoPosible + oRstCantidadNUnicos.Fields(0)
    End If
    
    If cantidadMaximoNueva >= cantidadMaximoPosible Then
        verificaCantidadDeItemsCombos = True
    End If
    
End Function


Private Sub BtnSalir_Click()
 FrameComboItem.Visible = False
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    HabilitaTimerColor (False)
    Dim i As Integer
    Sw = False
    Select Case Index
           Case Is = 0 'Grupo
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                sTipo = ""
                Isql = "select * from vGrupo where lActivo = 1 Order by Descripcion"
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Grupo", 2, "Descripcion", 5800, 0, 0, "", _
                                                                "Boton", 2, "nBoton", 1200, 1, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                   For i = 1 To 5
                       cmdGrupo(i).backColor = vbButtonFace
                       cmdSubGrupo(i).backColor = vbButtonFace
                   Next i
                   cmdSubGrupo(6).backColor = vbButtonFace
                   cmdSubGrupo(7).backColor = vbButtonFace
                   
                   sGrupo = sCodigo
                   xGrupo = sCodigo
                   RsGrupo.MoveFirst
                   RsGrupo.Find ("Codigo ='" & sCodigo & "'")
                   If Not RsGrupo.EOF And RsGrupo!nBoton > 0 Then
                      cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
                   End If
                   RsSubgrupo.Filter = ("tGrupo = '" & xGrupo & "'")
                   Call AsignaBoton(7, RsSubgrupo, cmdSubGrupo())
                                      
                   If RsSubgrupo.RecordCount <> 0 Then
                      RsSubgrupo.MoveFirst
                      RsSubgrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & sGrupo & "' and lActivo=1", Cn) & "'"
                      sSubGrupo = RsSubgrupo!codigo
                      xSubGrupo = RsSubgrupo!codigo
                      If Not RsSubgrupo.EOF And RsSubgrupo!nBoton > 0 Then
                         cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
                      End If
                   Else
                      sSubGrupo = ""
                      xSubGrupo = ""
                   End If
                   
                   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
                   Call AsignaBotonProducto(20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio)
                End If
                
                txtBarra.SetFocus
           
           Case Is = 1 'SubGrupo
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                sTipo = ""
                Isql = "select * from vSubGrupo where lActivo = 1 and tGrupo ='" & xGrupo & "' Order by Descripcion"
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "SubGrupo", 2, "Descripcion", 5800, 0, 0, "", _
                                                                "Boton", 2, "nBoton", 1200, 1, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                         
                If wEnter Then
                   For i = 1 To 7
                       cmdSubGrupo(i).backColor = vbButtonFace
                   Next i
                
                   sSubGrupo = sCodigo
                   xSubGrupo = sCodigo
                   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
                   
                   RsSubgrupo.MoveFirst
                   RsSubgrupo.Find ("Codigo ='" & sCodigo & "'")
                   If Not RsSubgrupo.EOF And RsSubgrupo!nBoton > 0 Then
                         cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
                   End If
                   Call AsignaBotonProducto(20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio)
                End If
                txtBarra.SetFocus
           
           Case Is = 2 'Producto
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                sTemp = ""
                sTipo = ""
                
                'cn
                'Isql = "select * from vProducto where lActivo = 1 and tSubGrupo ='" & xSubGrupo & "' and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                Isql = "select * from vProducto where lActivo = 1 AND TUNIDADNEGOCIO='" & sUnidadNegocio & "' and tSubGrupo ='" & xSubGrupo & "' and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                                "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                                "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                                "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                                "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
   
                If wEnter Then
                    'INSUMOCRITICO23
                    If validadIngresoProducto(sCodigo) = False Then
                        Exit Sub
                    End If
                    'INSUMOCRITICO23
            
                
                   sProducto = sCodigo
                   ListarOperadoresConFiltro (sProducto) 'Oscar Ortega
                   RsProducto.MoveFirst
                   RsProducto.Find "Codigo = '" & sProducto & "'"
                                                
                   If vComanda Then
                      sTipo = "Comanda"
                      sDescrip = sComanda
                      frmNumPad.Caption = "Comanda"
                      frmNumPad.Show vbModal
                      sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
                       
                      If wComanda And sComanda = "" Then
                         MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
                         Exit Sub
                      End If
                    
                      If Not wEnter Then
                         Exit Sub
                      End If
                   End If
                    
                   If lBal And RsProducto!lBalanza Then
                      Dim nResultado As Double
                      nResultado = Pesar(nBalanzaPuerto)
                      
                      If (lActivaTresDecimales) Then
                        nResultado = Format(nResultado, "#,##0.000")
                      Else
                        nResultado = Format(nResultado, "#,##0.00")
                      End If
                      
                      If nResultado > 0 Then
                         InsertaProducto nResultado
                      End If
                   Else
                   nCantidad = 1
                      InsertaProducto 1
                   End If
            
                    If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                       lPropiedad = True
                    End If
                End If
                txtBarra.SetFocus
                
           Case Is = 3 'Mozo
                sTemp = ""
                sTipo = ""
                Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by Descripcion"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   sMozo = sCodigo
                   txtMozo.Caption = sDescrip
                End If
                
           Case Is = 4 'Motorizado
                sTemp = ""
                Isql = "select * from vMotorizado where lActivo = 1 Order by Descripcion"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   sMotorizado = sCodigo
                   txtMotorizado.Caption = sDescrip
                End If
                txtBarra.SetFocus
                           
           Case Is = 5 'Punto de Venta
                sTipo = "Infhotel"
                sTemp = ""
                Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
                       " From tPuntoVenta " & _
                       " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   sPuntoVenta = sCodigo
                   txtPuntoVenta.Caption = sDescrip
                End If
                txtBarra.SetFocus

           Case Is = 6 'Operadores
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                            
                sTipo = ""
                sTemp = ""
                'Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
                ListarOperadoresConFiltro (sProducto)
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                    Screen.MousePointer = vbHourglass
                    For i = 1 To 13
                        cmdOperador(i).backColor = vbButtonFace
                    Next i
                    RsOperador.MoveFirst
                    RsOperador.Find "Codigo='" & sCodigo & "'"
                    If Not RsOperador.EOF And RsOperador!nBoton > 0 Then
                       cmdOperador(RsOperador!nBoton).backColor = vbRed
                    End If
                    AsignaPropiedad
                    xOperador = sCodigo
                    Screen.MousePointer = vbDefault
                End If
                txtBarra.SetFocus

           Case Is = 7 'ProductoCombo
                 If RsCabecera.RecordCount = 0 Then
                    Exit Sub
                 End If
                 
                 sTemp = ""
                 If lComboGeneral Then
                    Isql = "select * from vProducto where lActivo = 1 and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " Order by Descripcion"
                 Else
                    Isql = "SELECT * FROM dbo.vProducto INNER JOIN dbo.TCOMBO ON dbo.vProducto.Codigo = dbo.TCOMBO.tCodigoProducto Where (dbo.vProducto.lActivo = 1) And tCombo='" & sProductoCombo & "' And " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", "lLlevar=1")) & " ORDER BY dbo.vProducto.Descripcion "
                 End If
                 Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                                 "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                                 "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                                 "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                                 "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
                 frmBusquedaRapida.nPredeterm = 1
                 frmBusquedaRapida.Show vbModal
        
                 If wEnter Then
                    sProducto = sCodigo
                    'ListarOperadoresConFiltro (sProducto) 'Oscar Ortega
                    Dim xxx As String
                    xxx = RsProducto.Filter
                    RsProducto.Filter = adFilterNone
                    RsProducto.MoveFirst
                    RsProducto.Find "Codigo = '" & sProducto & "'"
                    
                    nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                       "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'", Cn)
                    If nCCombo < nCombo * RsDetalle!nCantidad Then
                        'Oscar Ortega----------------------------------------------
                         'Obtener Si es LUnico
                         Dim oRsProductoDeCombo As Recordset
                         Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
                         If oRsProductoDeCombo.RecordCount > 0 Then
                             If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                                  'Obtener Suma de cantidades
                                  Dim nCantidadEnElCombo As Integer
                                  nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sPedido, sItem, oRsProductoDeCombo!tEtiqueta)
                                  'Suma de cantidades < que nCantidad
                                  If nCantidadEnElCombo < nCantidad Then
                                      InsertaCombo sProducto
                                  Else
                                      MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                                  End If
                             Else
                                 InsertaCombo sProducto
                             End If
                         Else
                             InsertaCombo sProducto
                         End If
                        '----------------------------------------------------------
                        'InsertaCombo sProducto
                    Else
                       MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
                    End If
                    RsProducto.Filter = IIf(xxx = "0", "", xxx)
                    
                 End If
                 txtBarra.SetFocus
                 
    Case 8 'propiedad
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                            
                sTipo = ""
                sTemp = ""
                
                'Isql = "select * from vOperador where lActivo = 1 Order by Descripcion "
                'ListarOperadoresConFiltro (sProducto)
                
                'CASO CORP.FERNANDEZ
                If BDLink = "1" Then
                    If wAgregaCombo = False Then
                            Isql = "usp_Inforest_ObtienePropiedadesBusqueda '" & xOperador & "','" & sProducto & "','" & sPedido & "','" & sItem & "','" & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & "','" & sLocal & "'"
                    Else
                            Isql = "usp_Inforest_ObtienePropiedadesComboBusqueda '" & xOperador & "','" & sCombo & "','" & sPedido & "','" & sItem & "','" & xItem & "','" & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & "','" & sLocal & "'"
                    End If
                Else
                    If wAgregaCombo = False Then
                            Isql = "usp_Inforest_ObtienePropiedadesBusqueda '" & xOperador & "','" & sProducto & "','" & sPedido & "','" & sItem & "','" & sAlmacenMDB & "','" & sLocal & "'"
                    Else
                            Isql = "usp_Inforest_ObtienePropiedadesComboBusqueda '" & xOperador & "','" & sCombo & "','" & sPedido & "','" & sItem & "','" & xItem & "','" & sAlmacenMDB & "','" & sLocal & "'"
                    End If
                End If
                
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Estado", 2, "tEstado", 1500, 0, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 5500, 0, 0, "")
                                                                
                frmBusquedaRapida.nPredeterm = 2
                frmBusquedaRapida.tipoBusqueda = "BusquedaVenta"
                frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
           '         Screen.MousePointer = vbHourglass
                    If wAgregaCombo = False Then
                            If Calcular("SELECT COUNT(*) AS CODIGO FROM TPRODUCTOPROPIEDAD WHERE tCodigoPedido='" & sPedido & "' AND tItem='" & sItem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sProducto & "'  and tenlace='" & tenlacebusquedaVenta & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                            AgregarPropiedadBusqueda sCodigo, sDescrip
                    Else
                    
                           If Calcular("SELECT COUNT(*) AS CODIGO FROM tcombopropiedad WHERE tCodigoPedido='" & sPedido & "' AND tItem='" & sItem & "' AND tCodigoPropiedad='" & sCodigo & "' AND TPRODUCTO='" & sCombo & "' and titemcombo='" & xItem & "'", Cn) = 0 Then
                                    wAgregarPropiedad = True
                            Else
                                    wAgregarPropiedad = False
                            End If
                            AgregarPropiedadBusqueda sCodigo, sDescrip
                    End If
                    
                    
            '        Screen.MousePointer = vbDefault
                End If
                txtBarra.SetFocus

    Case Is = 9 'Origen de ventas
                sTemp = ""
                Isql = "select CodOrigenVenta as Codigo,* from vOrigenVenta where Activo = 1 and CodCanalVenta='" & sTipoPedido & "' and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Descripcion"
                Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Codigo", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.tipoBusqueda = "OrigenVentas"
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   vOrigenVentas = sCodigo
                    Call validarOrigenVentas
                End If
                txtBarra.SetFocus

        
   End Select
   
   HabilitaTimerColor (True)
End Sub
Public Sub AgregarPropiedadBusqueda(ByVal CodigoPropiedad As String, ByVal DescripcionPropiedad As String)
    If lMCPV Then
        If Not ValidaExistenciaProducto() Then
               MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
               RsCabecera.Requery
                RsDetalle.Requery
               Exit Sub
            End If
    End If

   '  HabilitaTimerColor (False)
   Dim ncantidadPropiedad As Double
   Dim Cantidad As Double

    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    Dim nBotonPropiedad As Double
    nBotonPropiedad = 999
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & DescripcionPropiedad & "'")
     RsPropiedad.Find ("Codigo = '" & CodigoPropiedad & "'")
   
     If Not (RsOperador.EOF Or RsOperador.BOF) Then
        nOperadorPropiedad = Calcular("select isnull(ncontrol,0) as codigo from voperador where codigo='" & RsOperador!codigo & "'", Cn)
     End If
    Dim k As Integer
    For k = 1 To 20
        If cmdPropiedad(k).Caption = DescripcionPropiedad Then
        nBotonPropiedad = k
        Exit For
        End If
    Next k
    
    If nBotonPropiedad <> 999 Then
        If cmdPropiedad(nBotonPropiedad).FontBold = True Then
            cmdPropiedad(nBotonPropiedad).FontBold = False
        Else
            cmdPropiedad(nBotonPropiedad).FontBold = True

        End If
    
    End If
    
    
     
                 
        If wAgregarPropiedad = False Then
          '  cmdPropiedad(nBotonPropiedad).FontBold = False
            If Not RsPropiedad.EOF Then
                If wAgregaCombo Then
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tcombopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)
                    Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
                Else
                    Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tproductopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' and tenlace='" & RsPropiedad!tEnlace & "'", Cn)
                    If RsPropiedad!nPrecio <> 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta - RsPropiedad!nPrecio, sPedido)
                    End If

                    Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "' and tEnlace='" & RsPropiedad!tEnlace & "'"
                End If
                If Cantidad <> 1 Then
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & DescripcionPropiedad & ": (" & Cantidad & "), ", "")
                Else
                           lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & DescripcionPropiedad & ", ", "")
                End If
    '            lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & DescripcionPropiedad & ", ", "")
            End If
        Else
        
            ncantidadPropiedad = 1
            If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
                sTipo = "Prepintado"
            
                sCodigo = ncantidadPropiedad
            
                frmNumPad.Show vbModal
                If wEnter And val(sDescrip) > 0 Then
            
                            ncantidadPropiedad = sDescrip
                        
                End If
            End If
            
            If nOperadorPropiedad > 0 Then
                If wAgregaCombo Then
                   Isql = "SELECT COUNT(dbo.TCOMBOPROPIEDAD.tCodigoPedido) AS codigo " & _
                          "FROM dbo.TCOMBOPROPIEDAD INNER JOIN dbo.TPROPIEDAD ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                          "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and dbo.TCOMBOPROPIEDAD.tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
                   If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                      MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                      Exit Sub
                   End If
                Else
                    Isql = "SELECT COUNT(dbo.TPRODUCTOPROPIEDAD.tCodigoPedido) AS codigo FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN " & _
                           "dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                           "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tOperador='" & RsOperador!codigo & "'"
                    If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                       MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                       Exit Sub
                    End If
                End If
            End If
            
            'If lcover Then
             Dim maxped As Double
             maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                 If maxped > 0 Then
                        If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + ((RsPropiedad!nPrecio * ncantidadPropiedad) * nCantidad)) Then
                          MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                          
    '                                 Screen.MousePointer = vbDefault
    '                                 RsDetalle.MoveLast
    '                                 RsDetalle.Requery
                          'RsProducto.MoveFirst
                          Exit Sub
     
                        End If
                 End If
             'End If

         '   cmdPropiedad(nBotonPropiedad).FontBold = True
            If Not RsPropiedad.EOF Then
               nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
               nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
               nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
               If wAgregaCombo Then
                    Cn.Execute "Insert into TCOMBOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1," & ncantidadPropiedad & " , " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & ", " & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & " , " & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
               Else
                    Cn.Execute "Insert into TPRODUCTOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1, " & ncantidadPropiedad & ",  " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & "," & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
                    If RsPropiedad!nPrecio <> 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad), sPedido)
                    End If
               End If
                If ncantidadPropiedad <> 1 Then
                
                      lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & DescripcionPropiedad & ": (" & ncantidadPropiedad & "), "
                Else
                      lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & DescripcionPropiedad & ", "
                End If
            End If
    End If
    
    If wAgregaCombo Then
       RsComboPropiedad.Requery
    Else
       RsProductoPropiedad.Requery
    End If
    
   ' HabilitaTimerColor (True)
End Sub


Private Sub cmdCabecera_Click(Index As Integer)

    HabilitaTimerColor (False)

    Dim sEstado As String
    txtBarra.SetFocus
    Dim j As Integer
    wCalcula = False
    
    Me.Caption = muestra
    
    Select Case Index
    
    
           Case Is = 0 'Nuevo
           
                idPedidoMesa = ""
                sValidaIntegracion = False
                grdCabecera.Enabled = False
                ActivaFrame fraCabecera
                ActivaCabecera False
                ActivaOpcion False
                fraTipoPedido.Enabled = True
                InicializaVariableCabecera
                cmdCabecera(13).Enabled = True
                cmdCabecera(9).FontBold = False
                'TimerDelivery.Interval = 0
                cmdMesa247.Visible = False
                
                'FrmOrdenCompra.Visible = True
                
                'Llena el Detalle
                RsDetalle.Filter = "tCodigoPedido = '" & sPedido & "'"
                Sw = True
                                              
                Select Case sTipoPedidoPD
                       Case Is = "01"
                            cmdCabecera_Click (6)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "02"
                            cmdCabecera_Click (7)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "03"
                            cmdCabecera_Click (8)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "04"
                            cmdCabecera_Click (17)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Is = "05"
                            cmdCabecera_Click (18)
                            vOrigenVentas = sDefaultOrigenVenta
                       Case Else
                End Select
                
                If lInfhotel Then
                    For j = 1 To 9
                        cmdPunto(j).FontBold = False
                        If cmdPunto(j).Caption = Calcular("select tDescripcion as codigo from tPuntoventa where tPuntoVenta='" & sPuntoVentaInfhotel & "'", CnInfhotel) Then
                           cmdPunto_Click (j)
                        End If
                    Next j
                    sPuntoVenta = sPuntoVentaInfhotel
                End If
                                
                'oo2
                ActivarDesactivarCD (False)
                                
               TimerDelivery.Enabled = False
               
               '------ visor 8" -------
               If lvisor Then
               
                   'Isql = "delete from infovisor where estado = 1"
                   Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
                   
                  ' Isql = "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "',0,1)"
                   Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "',0,1)"
                   
               End If
               
               'If lcover Then
                        cmdCabecera(21).Caption = "Monto Max/P : 0"
               'End If

           Case Is = 1 'Modifica
                'TimerDelivery.Interval = 0
                cmdMesa247.Visible = False
                
                If RsCabecera.RecordCount <> 0 Then
                   If lMCPV Then
                      If Not ValidaExistenciaProducto() Then
                         MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                         RsCabecera.Requery
                         RsDetalle.Requery
                         Exit Sub
                      End If
                   End If

                   ActivaFrame fraCabecera
                   ActivaCabecera False
                   ActivaOpcion False
                   wDetalle = False
                   
                   'FrmOrdenCompra.Visible = True
                   
                   cmdCabecera(13).Enabled = False
                   cmdOpcion(1).Enabled = True
                
                   If lmodificatipoPedido = False Then
                        fraTipoPedido.Enabled = False
                   Else
                        fraTipoPedido.Enabled = True
                   End If
                    
                    'Canales de Venta
                    RsCanalesVenta.MoveFirst
                    RsCanalesVenta.Filter = "Codigo = '" & sTipoPedido & "'"
                    lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
                    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
                    If lMCPV Then
                        lActivaMozo = False
                    End If
                    lObligaMozo = IIf(IsNull(RsCanalesVenta!lObligaMozo), False, RsCanalesVenta!lObligaMozo)
                    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
                    lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
                    lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
                    lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
                    lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
                    lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                     
                    If lActivaMozo Then
                       fraMozo.Visible = True
                       Me.fraMozo.Caption = Mesero
                    Else
                       fraMozo.Visible = False
                    End If
                    
                    If lActivaMotorizado Then
                       fraMotorizado.Visible = True
                    Else
                       fraMotorizado.Visible = False
                    End If
                    If lInfhotel Then
                       For j = 1 To 9
                           cmdPunto(j).FontBold = False
                           If cmdPunto(j).Caption = txtPuntoVenta.Caption Then
                              cmdPunto(j).FontBold = True
                           Else
                              cmdPunto(j).FontBold = False
                           End If
                       Next j
                    End If
                End If
                'If lcover Then
                    cmdCabecera(21).Caption = "Monto Max/P : " & Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
                'End If
                
                TimerDelivery.Enabled = False
                
                
                
           Case Is = 2 'Elimina
                Dim i As Integer
                If sPedido <> "" Then
                    If lMCPV Then
                        If Not ValidaExistenciaProducto() Then
                            MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                            RsCabecera.Requery
                            RsDetalle.Requery
                            Exit Sub
                        End If
                    End If
                    
                     If MESA247 Then
                         If Calcular("SELECT COUNT(*) AS CODIGO FROM MPEDIDO WHERE  tcodigopedido='" & sPedido & "' AND isnull(idPedidoMesa,'')<>'' ", Cn) > 0 Then
                            If MsgBox("El pedido esta enlazado a Mesa247, ¿Desea Continuar?", vbYesNo) = vbNo Then
                               Exit Sub
                            End If
                         End If
                    End If


                   If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0')<>'0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
                      MsgBox "Imposible eliminar pedidos con items Facturados", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                   'GCAA 10/07/2023 - REST-012-C-2023
                   If lBloqueoAnulaPedido Then
                        MsgBox "Esta opcion se encuentra Bloqueada, favor de comunicar al Supervisor", vbExclamation, sMensaje
                        Exit Sub
                   End If
                   
                   If MsgBox("Seguro de Eliminar el Pedido?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                      sUsuarioAutoriza = sUsuario
                      If lPasswordC Then
                         If Supervisor("01") = False Then
                            MsgBox "Clave no permitida", vbExclamation, sMensaje
                            Exit Sub
                         End If
                         sUsuarioAutoriza = sVar1
                         'Restaura variable sVar1 al valor anterior (al usuario logeado) - FERM 29/10/2024
                         sVar1 = sUsuario
                      End If
                                                                                                                                                                                         
                      If lEliminaC Then
                         wCabecera = True
                         ActivaFrame fraEliminacion
                         ActivaCabecera False
                         ActivaOpcion False
                         cmdOpcion(16).Visible = True
                         cmdOpcion(1).Enabled = False
                         cmdOpcion(2).Enabled = False
                      Else
                         sCodigo = ""
                         sDescrip = ""
                         EliminaCabecera
                      End If
                   End If
                End If
                
           Case Is = 3 'Mesa
                Dim xSql As String
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
                If lcover Then
                   If Supervisor("26") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If

                sTipo = "M"
                frmMesa.Show vbModal
                
                If wEnter Then
                   Cn.Execute "Update MPEDIDO set nAdulto = " & val(txtAdulto.Text) & ", nNino = " & val(txtNino.Text) & ", nMesa= " & val(txtJuntar.Text) & "  where tCodigoPedido ='" & sPedido & "'"

                   RsCabecera.Requery
                   RsCabecera.Find "Codigo = '" & sPedido & "'"
                End If
                
                If wMesa Then
                   xSql = "nAdulto = " & val(txtAdulto.Text) & ", nNino = " & val(txtNino.Text) & ", nMesa= " & val(txtJuntar.Text)
                   If wCambioMesa Then
                      If Calcular("select count(tMesa) as Codigo from TPEDIDOMESA where tCodigoPedido='" & sPedido & "'", Cn) > 0 Then
                         MsgBox "No se puede mover la Mesa cuando tiene mesas asociadas", vbCritical, sMensaje
                         Exit Sub
                      End If
                      
                      sEstado = Calcular("select tEstadoMesa as Codigo from TMESA where tCodigoMesa ='" & sCodigo & "'", Cn)
                      If sEstado = "01" Or sEstado = "04" Or sEstado = "0" Then
                         xSql = xSql & ", tMesa = '" & sCodigo & "', tSalon='" & IIf(sCodigo <> "", Calcular("select tsalon as Codigo from tmesa where tCodigoMesa='" & sCodigo & "'", Cn), sSalon) & "'"
                         Cn.Execute "Update TMESA set tEstadoMesa='02' where tCodigoMesa ='" & sCodigo & "'"
                         Cn.Execute "Update TMESA set tEstadoMesa='04' where tCodigoMesa ='" & sTempMesa & "'"
                         sMesa = sCodigo
                                                                                                                          
                         If lPrinter And lCambioMesa Then
                            sDescrip = "Cambio de Mesa/Ubicacion a " & Calcular("SELECT dbo.vSalon.Descripcion + ' - ' + dbo.TMESA.tDetallado AS Codigo FROM dbo.TMESA INNER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo where tCodigoMesa='" & sCodigo & "'", Cn)
                            Dim RsTemp As Recordset
                            frmPrecuentaImpresora.cmdOpcion(0).Enabled = False
                            frmPrecuentaImpresora.Show vbModal
                            If Not wEnter Then
                               Exit Sub
                            End If
                                                           
                            Set RsTemp = Lib.OpenRecordset("SELECT dbo.vTipoPedido.Descripcion AS TipoPedido, dbo.vMozo.Descripcion AS Mozo, dbo.vSalon.Descripcion + ' - ' + dbo.TMESA.tResumido AS Mesa, dbo.MPEDIDO.tObservacion AS Observacion, dbo.MPEDIDO.nAdulto AS Pax FROM dbo.vMozo RIGHT OUTER JOIN dbo.MPEDIDO LEFT OUTER JOIN dbo.TMESA INNER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo LEFT OUTER JOIN dbo.vTipoPedido ON dbo.MPEDIDO.tTipoPedido = dbo.vTipoPedido.Codigo where tCodigoPedido='" & sPedido & "'", Cn)
                            ImprimeMensaje sCodigo, sDescrip, RsTemp
                            Set RsTemp = Nothing
                         End If
                         
                         Cn.Execute "Update MPEDIDO set " & xSql & "  where tCodigoPedido ='" & sPedido & "'"
                         'modifca mesa / salon
                         Call Log_Inforest("PUNTO VENTA", "MODIFICACIÓN DE MESA", sPedido, Error, "", xSql, "Modificación en mesa", "Se realizó la modificación de la mesa asignada", sUsuario)
                    
                         RsCabecera.Requery
                         RsCabecera.Find "Codigo = '" & sPedido & "'"
                         wMesa = False
                      Else
                         MsgBox "Mesa Ocupada", vbExclamation, sMensaje
                         Exit Sub
                      End If
                   End If
                End If
                
           Case Is = 4 'Transferencia
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
                If lPasswordTransferencia Then
                   If Supervisor("15") = False Then
                      MsgBox "Clave no permitida", vbExclamation, sMensaje
                      Exit Sub
                   End If
                  'inicia lg 25/06
                   sUsuarioAutoriza = sVar1
                   'fin
                End If
                
                frmDivision.txtFecha = txtFecha.Caption
                frmDivision.txtTipoPedido.ForeColor = txtTipoPedido.ForeColor
                frmDivision.txtTipoPedido.Caption = txtTipoPedido.Caption
                frmDivision.TxtDescuento.Caption = TxtDescuento.Caption
                frmDivision.txtMozo.Caption = txtMozo.Caption
                frmDivision.txtMotorizado.Caption = txtMotorizado.Caption
                frmDivision.txtObservacion.Caption = txtObservacion.Caption
                'inicia 25/06
                frmDivision.usuarioautoriza sUsuarioAutoriza
                'fin
                
                frmDivision.txtMesa.Caption = grdCabecera.Columns(1).Text
                frmDivision.txtOrden.Caption = grdCabecera.Columns(0).Text
                frmDivision.Show vbModal
                If wEnter Then
                   RsCabecera.Requery
                   RsDetalle.Requery
                   RsCombo.Requery
                   RsProductoPropiedad.Requery
                   If RsCabecera.RecordCount <> 0 Then
                      RsCabecera.Find "Codigo = '" & sPedido & "'"
                   End If
                    'MESA247
                    If MESA247 Then
                        If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                              Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                        End If
                    End If
                    
                End If
                
           Case Is = 5 'Observación
                If RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
           
                frmKeyBoard.Caption = "Nombre / Observación"
                frmKeyBoard.txtResultado.Text = txtObservacion.Caption
                frmKeyBoard.Show vbModal
                
                If wEnter = True Then
                   txtObservacion.Caption = sDescrip
                   Cn.Execute "Update MPEDIDO set tObservacion = '" & sDescrip & "' where tCodigoPedido ='" & sPedido & "'"
                   
                   'registro de log modificaciòn de observaciòn del pedido
                    Call Log_Inforest("PUNTO VENTA", "MODIFICA OBSERVACIÓN DE PEDIDO", sPedido, Error, sDescrip, "Modifica observación: " & sDescrip, "Modifica observación del pedido " & sPedido & " -  ''" & sDescrip & "'' ", "Se ha modificado la observación del pedido(" & sPedido & "): ''" & sDescrip, sUsuario)
   
                   RsCabecera.Requery
                   RsCabecera.Find "Codigo = '" & sPedido & "'"
                   
                   'MESA247
                   If MESA247 Then
                        If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                               Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                        End If
                   End If
                   
                End If
                
           Case Is = 6 'Primer Boton
                xMontoMaximo = 0
                txtTipoPedido.Caption = sBoton1
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "01"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = True
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = False
                cmdCabecera(17).FontBold = False
                cmdCabecera(18).FontBold = False
                vOrigenVentas = ""
                'Call ActivarOrigenVentas
           Case Is = 7 'Segundo boton
'                txtTipoPedido.Caption = sBoton2
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "02"
                
                
                Isql = "select isnull(tCodigoPedidoCD,'')tCodigoPedidoCD from mpedido where tCodigoPedido = '" & sPedido & "'"
                Set RsverCD = Lib.OpenRecordset(Isql, Cn)
                If Not (RsverCD.EOF Or RsverCD.BOF) Then
'                    ' MsgBox (RsverCD!tCodigoPedidoCD)
                    If RsverCD!tCodigoPedidoCD <> "" Then
'                        If sTipoPedido <> "02" Then
'                            InicializaVariableCanales
'                            txtTipoPedido.Caption = sBoton2
'                        End If
                    Else
                        InicializaVariableCanales
                        txtTipoPedido.Caption = sBoton2
                    End If
                Else
                       InicializaVariableCanales
                        txtTipoPedido.Caption = sBoton2
                 End If
                
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = True
                cmdCabecera(8).FontBold = False
                cmdCabecera(17).FontBold = False
                cmdCabecera(18).FontBold = False
                vOrigenVentas = ""
                'Call ActivarOrigenVentas
           Case Is = 8 'Tercer Boton
                txtTipoPedido.Caption = sBoton3
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "03"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = True
                cmdCabecera(17).FontBold = False
                cmdCabecera(18).FontBold = False
                vOrigenVentas = ""
                ' Call ActivarOrigenVentas
           Case Is = 17 'Canal4
                txtTipoPedido.Caption = sBoton4
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "04"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = False
                cmdCabecera(17).FontBold = True
                cmdCabecera(18).FontBold = False
                vOrigenVentas = ""
                ' Call ActivarOrigenVentas
           Case Is = 18 'Canal5
                txtTipoPedido.Caption = sBoton5
                txtTipoPedido.ForeColor = &H800000
                sTipoPedido = "05"
                InicializaVariableCanales
                cmdCabecera(6).FontBold = False
                cmdCabecera(7).FontBold = False
                cmdCabecera(8).FontBold = False
                cmdCabecera(17).FontBold = False
                cmdCabecera(18).FontBold = True
                vOrigenVentas = ""
                'Call ActivarOrigenVentas
           Case Is = 9 'Alta Prioridad
                txtTipoPedido.ForeColor = IIf(txtTipoPedido.ForeColor = vbRed, &H800000, vbRed)
                cmdCabecera(9).FontBold = IIf(cmdCabecera(9).FontBold = True, False, True)
                    
           Case Is = 10 'Programa
                frmPrograma.Show vbModal
                If wEnter = True Then
                   txtFechaProg.Caption = sCodigo
                   txtEnvioAntes.Text = lMinutoEnvioAntes
                Else
                   txtFechaProg.Caption = ""
                End If
                lMinutoEnvioAntes = 0
                
           Case Is = 11 'Observación
                frmKeyBoard.Caption = "Nombre / Observación"
                frmKeyBoard.txtResultado.Text = txtObservacion.Caption
                frmKeyBoard.Show vbModal
                If wEnter = True Then
                   txtObservacion.Caption = sDescrip
                End If
                      
           Case Is = 12 'Cliente
           
           'If Not Sw Then
                If Calcular("select isnull(nDescuento,0) as codigo from MPEDIDO where tCodigoPedido='" & sPedido & "' ", Cn) > 0 Then
                    MsgBox "No se puede cambiar Cliente del Pedido, Pedido con Descuento!! Revertir Descuento!!", vbInformation
                    Exit Sub
                End If
           'End If
          
           
            If lClub = False Then
                 If Not Sw Then
                   sTemp = txtTelefono.Caption
                End If
                
                wEnter = False
                sTipo = sTipoPedido
                frmBusquedaDelivery.txtResultado.Caption = txtTelefono.Caption
                frmBusquedaDelivery.Show vbModal
            Else
                If Not Sw Then
                    frmBusquedaSocio.sCodigoPariente = sCodigoParienteSeleccionado
                    frmBusquedaSocio.sCodigoSocio = txtTelefono.Caption
                    frmBusquedaSocio.sCodigoInvitadoB = sCodigoInvitado
              
                Else
                    frmBusquedaSocio.sCodigoPariente = ""
                    frmBusquedaSocio.sCodigoSocio = ""
                    frmBusquedaSocio.sCodigoInvitadoB = ""
         
                End If

                    frmBusquedaSocio.Show vbModal
            End If
                        
                                
                If wEnter = True Then
                   txtCliente.Caption = ""
                   sCliente = sCodigo
                   
                   'pariente2013
                   If sCodigoParienteSeleccionado <> "" Then
                        txtCliente.Caption = Calcular("select pariente + ' ' +case when lconyugue=1 then ' Conyugue de ' else 'Hijo de ' end as codigo from vPariente where tCodigoPariente='" & sCodigoParienteSeleccionado & "' and tCodigoDelivery='" & sCliente & "'", Cn)
                   End If
                   txtCliente.Caption = txtCliente.Caption & " " & sDescrip
                   'pariente2013
                   
                   Set RsTemporal = Lib.OpenRecordset("Select * from TDELIVERY where tcodigoDelivery='" & sCliente & "'", Cn)
                   If Not RsTemporal.EOF Then
                    If lClub Then
                        txtTelefono.Caption = IIf(IsNull(RsTemporal!tCodigoDelivery), "", RsTemporal!tCodigoDelivery)
                    Else
                        txtTelefono.Caption = IIf(IsNull(RsTemporal!tTelefono), "", RsTemporal!tTelefono)
                    End If
                    
                    
                        Dim estadoTipo As String
                        estadoTipo = Calcular("select isnull(tresumido,'') as codigo from vTipoClienteFrecuente  where Codigo='" & RsTemporal!tTipoCliente & "' ", Cn)
                        estadoTipo = estadoTipo + " - " + Calcular("select isnull(tresumido,'') as codigo from vEstadoFrecuente  where Codigo='" & RsTemporal!tEstadoFrecuente & "' ", Cn)
                        'txtDireccion.Caption = estadoTipo
                        txtDireccion.Caption = IIf(IsNull(RsTemporal!tDireccion), "", RsTemporal!tDireccion)
                        xDescuento = IIf(IsNull(RsTemporal!nDescuento), 0, RsTemporal!nDescuento)
                        sCodigoDescuento = "000"
                     
                   Else
                        txtTelefono.Caption = ""
                        txtDireccion.Caption = ""
                        xDescuento = 0
                        sCodigoDescuento = ""
                   End If
                   If sTienda <> "" Then
                      Set RsTemporal = Lib.OpenRecordset("Select * from TTIENDA where tCodigoDelivery='" & sCliente & "' and tCodigoTienda='" & sTienda & "'", Cn)
                      If Not RsTemporal.EOF Then
                         txtCliente.Caption = txtCliente.Caption + " - " + IIf(IsNull(RsTemporal!tNombre), "", RsTemporal!tNombre)
                         txtDireccion.Caption = IIf(IsNull(RsTemporal!tDireccion), "", RsTemporal!tDireccion)
                      
                         xDescuento = 0 'IIf(IsNull(RsTemporal!nDescuento), 0, RsTemporal!nDescuento)
                         sCodigoDescuento = "000"
                      
                      Else
                         txtCliente.Caption = ""
                         txtDireccion.Caption = ""
                         xDescuento = 0
                         sCodigoDescuento = ""
                      End If
                   End If
                   
                   If sTienda <> "" Then
                      Set RsTemporal = Lib.OpenRecordset("Select * from TTIENDA where tCodigoDelivery='" & sCliente & "' and tCodigoTienda='" & sTienda & "'", Cn)
                      If Not RsTemporal.EOF Then
                         txtCliente.Caption = txtCliente.Caption + " - " + IIf(IsNull(RsTemporal!tNombre), "", RsTemporal!tNombre)
                        ' txtDireccion.Caption = IIf(IsNull(RsTemporal!tDireccion), "", RsTemporal!tDireccion)
                      Else
                         txtCliente.Caption = ""
                         txtDireccion.Caption = ""
                      End If
                   End If
                   
                   
                   
                   'invitado2013
                   If sCodigoInvitado <> "" Then
                      Set RsTemporal = Lib.OpenRecordset("Select isnull(tnombre,'')   + ' ' + isnull(tapellido,'') as invitado from tdeliveryinvitado where tCodigoinvitado='" & sCodigoInvitado & "'", Cn)
                      If Not RsTemporal.EOF Then
                        txtObservacion.Caption = "Inv: " & IIf(IsNull(RsTemporal!Invitado), "", RsTemporal!Invitado)
                         
                      Else
                         txtObservacion.Caption = ""
                      End If
                   End If
                   'invitado2013
                   
                   If xDescuento > 0 Then
                      sCodigoDescuento = "000"
                      sDescripcionDescuento = "DESCUENTO POR CLIENTE DELIVERY FRECUENTE"
                   Else
                      sCodigoDescuento = ""
                      sDescripcionDescuento = ""
                   End If
                   wCalcula = True
                   
                End If

           Case Is = 13 'Mesa
                sTipo = "A"
                
                Dim RsMesax As Recordset
                Set RsMesax = Lib.OpenRecordset("select * from TMESA where tCodigoMesa <> '000' and lActivo= 1 order by tCodigoMesa", Cn)
   
                If RsMesax.RecordCount = 0 Then
                   MsgBox "No existen mesas registradas", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                frmMesa.cmdOpcion(6).Enabled = False
                frmMesa.Show vbModal
                If wMesa Then
                   sEstado = Calcular("select tEstadoMesa as Codigo from TMESA where tCodigoMesa ='" & sCodigo & "'", Cn)
                   If sEstado = "01" Or sEstado = "04" Then
                      sMesa = sCodigo
                   Else
                      MsgBox "Mesa Ocupada", vbExclamation, sMensaje
                      Exit Sub
                   End If
                End If
                           
           Case Is = 14 'Habitación
                sCodigo = ""
                frmCargo.cmdOpcion(5).Enabled = False
                frmCargo.cmdOpcion(0).Enabled = False
                frmCargo.cmdOpcion(16).Enabled = False
                frmCargo.cmdOpcion(17).Enabled = False
                frmCargo.cmdOpcion(18).Enabled = False
                
                frmCargo.Show vbModal
                If wEnter Then
                   If sDescrip = "Habitacion" Then
                      txtObservacion.Caption = "Hab: " & sHabitacion
                      sReserva = ""
                   Else
                      txtObservacion.Caption = "Res: " & sReserva
                      sHabitacion = ""
                   End If
                Else
                   sHabitacion = ""
                   sReserva = ""
                End If
           
           Case Is = 15 'Descuento
                If Supervisor("10") = False Then
                   MsgBox "Clave no permitida", vbExclamation, sMensaje
                   Exit Sub
                End If
                
                
                Dim AnticipoGenera As Boolean
                AnticipoGenera = False
                
                If lActivaAnticipo Then
                    Dim CodigoReservaAnticipo2 As String
                    CodigoReservaAnticipo2 = Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
                    If CodigoReservaAnticipo2 <> "" Then
                        MsgBox "No es posible aplicar descuento, " & vbNewLine & " ya que existe anticipo asignado de la reserva: " & CodigoReservaAnticipo2, vbExclamation, sMensaje
                        Exit Sub
                    End If
                End If
                                
                                
                                
                sUsuarioAutoriza = sVar1
                tAutorizaDescuento = sUsuarioAutoriza
                sTemp = ""
                
                If Not AnticipoGenera Then
                    Isql = "SELECT Codigo, LTRIM(RTRIM(Descripcion)) as Descripcion, case lRatio when 1 then nRatio else 0 END as nRatio, case lRatio when 0 then nRatio else 0 END as nMonto FROM vMotivoDescuento WHERE lActivo = '1' and AplicaAnticipo=0 ORDER BY Descripcion"
                    Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                                "Descripción", 2, "Descripcion", 4300, 0, 0, "", _
                                                                "Porcentaje", 2, "nRatio", 1200, 1, 0, "###,##0.00", _
                                                                "Monto", 2, "nMonto", 1200, 1, 0, "###,##0.00")
                End If
                If AnticipoGenera Then
                    Isql = "Select treserva as Codigo,isnull((select tapellido+tnombre from vdelivery where codigo=treserva.tCodigoDelivery),tObservacion )as Descripcion,0 as nRatio,(Select sum(nventa) from mdocumento where treserva=Treserva.treserva and testadoDocumento<>'04') as nMonto from treserva where testadoreserva='01' and (Select sum(nventa) from mdocumento where treserva=Treserva.treserva and testadodocumento<>'04')>0 ORDER BY Descripcion"
                    Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                                "Descripción", 2, "Descripcion", 4300, 0, 0, "", _
                                                                "Porcentaje", 2, "nRatio", 1200, 1, 0, "###,##0.00", _
                                                                "Monto", 2, "nMonto", 1200, 1, 0, "###,##0.00")
                End If
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                wCalcula = True
                
                
                If wEnter Then
                   If Not AnticipoGenera Then
                    
                       sCodigoDescuento = sCodigo
                       Dim RsDesc As ADODB.Recordset
                       Set RsDesc = Lib.OpenRecordset("select * from vMotivoDescuento where Codigo = '" & sCodigo & "'", Cn)
                       If RsDesc.EOF Then
                          Exit Sub
                       End If
                       RsDesc.MoveFirst
                       xDescuento = RsDesc!nRatio
                       lRatio = RsDesc!lRatio
                       
                       Dim SumTotalPedido As Double
                       SumTotalPedido = Calcular("select sum(d.nventa) as codigo from DPEDIDO d inner join TPRODUCTO p on d.tCodigoProducto = p.tCodigoProducto where d.tCodigoPedido='" & sPedido & "' and p.lDescuento = 1", Cn)
                       
                       'CDbl(txtSuma.Caption)
                       If Not RsDesc!lRatio And (RsDesc!nRatio > SumTotalPedido) Then
                          sCodigoDescuento = ""
                          xDescuento = 0
                          MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                          Exit Sub
                       End If
                       
                       If sCodigo = "000" Then
                            sTipo = "Prepintado"
                            frmNumPad.Show vbModal
                            xDescuento = val(sDescrip)
                            
                            If xDescuento > 99 Then
                              xDescuento = 0
                              MsgBox "Porcentaje no válido", vbCritical, sMensaje
                              Exit Sub
                            End If
                            
                            frmKeyBoard.Caption = "Descripcion del Descuento"
                            frmKeyBoard.Show vbModal
                            sDescripcionDescuento = sDescrip
                       Else
                            If RsDesc!lBloqueo Then
                                 sTipo = "Prepintado"
                                 sCodigo = xDescuento
                                 frmNumPad.Show vbModal
                                 
                                'Cambio de monto o % Descuento
                                xDescuento = val(sDescrip)
                                If (RsDesc!lRatio) Then
                                    If xDescuento > 99 Then
                                        MsgBox "Descuento Incorrecto"
                                        Exit Sub
                                    End If
                                Else
                                    'CDbl(txtSuma.Caption)
                                    If xDescuento > SumTotalPedido Then
                                        MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                                        ValidaDesc = False
                                        Exit Sub
                                    End If
                                End If
                                '----------------------------
                            Else
                                wEnter = False
                            End If
                       End If
                       
                       'Cn.Execute "update treserva set testadoreserva='01' where treserva in (select isnull(treservainf,'') from mpedido where tcodigopedido='" & sPedido & "') "
                       'Cn.Execute "update mpedido set treservainf='' where tcodigopedido='" & sPedido & "'"
                       
                       
                    End If
                    
                    If AnticipoGenera Then
                    
                       Dim RsDesc2 As ADODB.Recordset
                       Set RsDesc2 = Lib.OpenRecordset("select top 1 * from vMotivoDescuento where AplicaAnticipo=1 ", Cn)
                       If RsDesc2.EOF Then
                          Exit Sub
                       End If
                       RsDesc2.MoveFirst
                       xDescuento = RsDesc2!nRatio
                       lRatio = RsDesc2!lRatio
                        If RsDesc2.RecordCount = 0 Or lRatio = True Then
                            MsgBox "La configuracion del Descuento para Anticipo esta mal configurada, favor de Revisar!!! ", vbInformation, sMensaje
                            sCodigoDescuento = ""
                            xDescuento = 0
                            Exit Sub
                        End If
                        
                       Dim montoanticipo As Double
                       montoanticipo = Calcular("select sum(nventa) as codigo from mdocumento where treserva='" & sCodigo & "' and testadodocumento='02'", Cn)
                        
                       Dim SumTotalp As Double
                       SumTotalp = Calcular("select sum(d.nventa) as codigo from DPEDIDO d inner join TPRODUCTO p on d.tCodigoProducto = p.tCodigoProducto where d.tCodigoPedido='" & sPedido & "' and p.lDescuento = 1", Cn)
                       
                       'CDbl(txtSuma.Caption)
                       If (montoanticipo > SumTotalp) Then
                          sCodigoDescuento = ""
                          xDescuento = 0
                          MsgBox "Descuento mayor al Pedido", vbCritical, sMensaje
                          'Cn.Execute "update mpedido set treservainf='' where tcodigopedido='" & sPedido & "'"
                          Exit Sub
                       End If
                       
                        sCodigoDescuento = RsDesc2!codigo
                        'scodigo = "000"
                        xDescuento = CDbl(montoanticipo)
                        lRatio = False
                        Cn.Execute "update treserva set testadoreserva='02' where treserva='" & sCodigo & "'"
                        Cn.Execute "update mpedido set treservainf='" & sCodigo & "' where tcodigopedido='" & sPedido & "'"
                        
                    End If
                Else
                   Exit Sub
                End If
                                                                
           Case Is = 16 'Contactos
                sTipo = "Infhotel"
                sTemp = txtObservacion.Caption
                
                Isql = "SELECT dbo.TCOMPANIA.tCodigoCompania + dbo.TCONTACTO.tCodigoContacto as Codigo, dbo.TCOMPANIA.tRazonSocialCompania, rtrim(tCarnet) + ' - ' +Rtrim(dbo.TCONTACTO.tPaterno) + ' ' + Rtrim(dbo.TCONTACTO.tMaterno) as Descripcion, dbo.TCONTACTO.tNombre, dbo.TCONTACTO.tCarnet " & _
                       "FROM dbo.TCOMPANIA INNER JOIN dbo.TCONTACTO ON dbo.TCOMPANIA.tCodigoCompania = dbo.TCONTACTO.tCodigoCompania " & _
                       "Where lInforest = 1 order by tpaterno, tmaterno, tnombre"
                       
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Compañía", 2, "tRazonSocialCompania", 3400, 0, 0, "", _
                                                                "Apellido", 2, "Descripcion", 3000, 0, 0, "", _
                                                                "Nombre", 2, "tNombre", 1800, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                If wEnter = True Then
                   sCompania = Mid(sCodigo, 1, 5)
                   sContacto = Mid(sCodigo, 6, 4)
                   txtObservacion.Caption = sDescrip
                End If
                
            Case Is = 19 ' entregar A
                frmKeyBoard.Caption = "Entregar A"
                frmKeyBoard.txtResultado.Text = txtEntregar.Caption
                frmKeyBoard.Show vbModal
                If wEnter = True Then
                    txtEntregar.Caption = sDescrip
                End If
            
            Case Is = 20 'Anticipos
                If sPedido = "" Then
                    MsgBox "Aun no se ha generado el pedido, imposible asignar anticipo!!!", vbExclamation, sMensaje
                    cmdCabecera(20).FontBold = False
                    Exit Sub
                End If
                
                'JSJR 21/01/2025
                Dim RsDPedidoCount As Recordset
                Set RsDPedidoCount = Lib.OpenRecordset("select top 1 tCodigoPedido from DPEDIDO where tCodigoPedido = '" + sPedido + "'", Cn)
                
                If RsDPedidoCount.RecordCount = 0 Then
                    MsgBox "No se han encontrado ítems para aplicar el anticipo.", vbExclamation
                    cmdCabecera(20).FontBold = False
                    Exit Sub
                End If
            
            
                Dim CodigoReservaAnticipo As String
                CodigoReservaAnticipo = Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
                If CodigoReservaAnticipo = "" Then
                    cmdCabecera(20).FontBold = False
                    sTemp = ""
                    

                    Isql = "Select treserva as Codigo,isnull((select tapellido+tnombre from vdelivery where codigo=treserva.tCodigoDelivery),tObservacion )as Descripcion,0 as nRatio,(Select sum(nventa) from mdocumento where treserva=Treserva.treserva and testadoDocumento<>'04') as nMonto from treserva where testadoreserva='01' and (Select sum(nventa) from mdocumento where treserva=Treserva.treserva and testadodocumento<>'04')>0 ORDER BY Descripcion"
                    Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                                "Descripción", 2, "Descripcion", 4300, 0, 0, "", _
                                                                "-", 2, "nRatio", 1200, 1, 0, "###,##0.00", _
                                                                "Monto", 2, "nMonto", 1200, 1, 0, "###,##0.00")
                    frmBusquedaRapida.nPredeterm = 1
                    frmBusquedaRapida.Show vbModal
                    If wEnter Then
                        Dim valorResultado As String
                        valorResultado = Calcular(" exec sp_AsignaAnticipo_Pedido '" & sPedido & "','" & sCodigo & "',3", Cn)
                        If valorResultado <> "0" Then
                            MsgBox valorResultado, vbExclamation, sMensaje
                            Exit Sub
                        End If
                    
                        Cn.Execute " exec sp_AsignaAnticipo_Pedido '" & sPedido & "','" & sCodigo & "',1"
                        cmdCabecera(20).FontBold = True
                        RsDetalle.Requery
                    End If
                Else
                    cmdCabecera(20).FontBold = True
                    If MsgBox("¿Esta seguro de desasignar el anticipo " & CodigoReservaAnticipo & " del pedido?", vbYesNo, sMensaje) = vbNo Then
                        Exit Sub
                    End If
                    
                    Cn.Execute " exec sp_AsignaAnticipo_Pedido '" & sPedido & "','" & CodigoReservaAnticipo & "',2"
                    cmdCabecera(20).FontBold = False
                    RsDetalle.Requery
                End If
                SumPrecioVenta (sPedido)
                
                
            Case Is = 21 'Monto Max Pedido Monto Max/P : 0
            
                If lcover Then
                    If val(frmVenta.txtAdulto.Text) = 0 Then
                        MsgBox ("#Pax del pedido no ingresado!"), vbExclamation
                        Exit Sub
                     End If
                End If
                
                sTipo = ""
                frmNumPad.Show vbModal
                xMontoMaximo = val(sDescrip)
'                Dim nAdultoPed As Integer
'                nAdultoPed = Val(Calcular("select nAdulto as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
'
                If lcover Then
                    If (sMontoMinCover * CDbl(frmVenta.txtAdulto.Text)) > xMontoMaximo Then
                        MsgBox "El Monto Minimo por Pedido es Mayor al Monto Maximo Ingresado!", vbInformation
                        xMontoMaximo = 0
                    End If
                End If
                'cmdCabecera(20).FontBold = IIf(cmdCabecera(20).FontBold = True, False, True)
                cmdCabecera(21).Caption = "Monto Max/P : " & CStr(IIf(xMontoMaximo = 0, 0, xMontoMaximo))
                                       
    End Select
    
    HabilitaTimerColor (True)
End Sub

Private Sub cmdCombo_Click(Index As Integer)
   HabilitaTimerColor (False)
   Dim nPos As Integer
   Dim nOrd As Integer
   txtBarra.SetFocus
   Select Case Index
   
          Case Is = 0 ' Salir
               fraCombo.Visible = False
               fraProductoCombo.Visible = False
               wAgregaCombo = False
               ActivaCabecera True
               ActivaOpcion True
               AsignaProducto
               If fraPropiedad.Visible = True Then
                  cmdOpcion_Click (8)
               End If
               RsDetalle.Requery
               
          Case Is = 1 ' Elimina
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
              If obtieneEliminaItemFijoCombo(RsCombo.Fields("tproducto"), RsCombo.Fields("tproductocombo")) = True Then
                    MsgBox "No se puede quitar este producto del combo. Consulte con el Administrador"
                    Exit Sub
               End If
               If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsCombo!lImprime) Then
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuarioAutoriza = sVar1
                  End If
                                                                            
                  If lPrinter = False Or (lPrinter = True And RsCombo!lImprime) Then
                     'Impresion del Pedidos Anulados
                     Isql = "SELECT TPRODUCTO_1.tDetallado AS Producto, dbo.vSalon.tResumido + ' - ' + dbo.TMESA.tResumido AS Mesa, dbo.TPRODUCTOAREA.tArea, dbo.MPEDIDO.tTipoPedido AS TipoPedido, dbo.MPEDIDO.nAdulto, dbo.MPEDIDO.lPrioridad AS Prioridad, dbo.MPEDIDO.tObservacion AS Observacion, dbo.vMozo.Descripcion AS Mozo, dbo.CPEDIDO.nCantidad AS nCombo, dbo.CPEDIDO.tItem, dbo.CPEDIDO.tItemCombo, dbo.CPEDIDO.tObservacion AS tObservacionCombo, TPRODUCTO_2.tDetallado AS Combo, dbo.vDelivery.Cliente " & _
                            "FROM dbo.TPRODUCTO TPRODUCTO_2 LEFT OUTER JOIN dbo.TPRODUCTOAREA ON TPRODUCTO_2.tCodigoProducto = dbo.TPRODUCTOAREA.tCodigoProducto RIGHT OUTER JOIN dbo.TMESA LEFT OUTER JOIN dbo.vSalon ON dbo.TMESA.tSalon = dbo.vSalon.Codigo RIGHT OUTER JOIN dbo.vMozo RIGHT OUTER JOIN dbo.vDelivery RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vDelivery.Codigo = dbo.MPEDIDO.tClienteDelivery LEFT OUTER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido ON dbo.vMozo.Codigo = dbo.MPEDIDO.tMozo ON " & _
                            "dbo.TMESA.tCodigoMesa = dbo.MPEDIDO.tMesa ON TPRODUCTO_2.tCodigoProducto = dbo.CPEDIDO.tProductoCombo LEFT OUTER JOIN dbo.TPRODUCTO TPRODUCTO_1 ON dbo.CPEDIDO.tProducto = TPRODUCTO_1.tCodigoProducto " & _
                            "Where dbo.CPEDIDO.lImprime = 1 And dbo.CPEDIDO.lImprimeArea = 1 and dbo.CPEDIDO.tCodigoPedido = '" & sPedido & "' and dbo.CPEDIDO.tItem ='" & RsCombo!tItem & "' and dbo.CPEDIDO.tItemCombo='" & RsCombo!tItemCombo & "' ORDER BY dbo.CPEDIDO.tItem"
              
                     Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                     Dim i As Integer
                     If RsImpresion.RecordCount = 0 Then
                        LimpiaRs
                     Else
                        RsArea.MoveFirst
                        For i = 1 To RsArea.RecordCount
                            RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
                            If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                               If RsImpresion.RecordCount <> 0 Then
                                  RsImpresion.MoveFirst
                                  ImprimeAnulaCombo RsImpresion, RsArea!timpresora, RsArea!Area, RsComboPropiedad
                               End If
                            End If
                            RsArea.MoveNext
                        Next i
                     End If
                  End If
                  
                  'Oscar Ortega----------------------------------------------
                  Dim RstCombo As Recordset
                  Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta - ((RstCombo!nAumento / nCantidad) * RstCombo!nCantidad), sPedido)
                    End If
                  End If
                  'Fin Oscar Ortega------------------------------------------
                  'KDS
                  If lKDS Then
                    Call KDS_EliminarProductoDeCombo(RsCabecera, sItem, xItem)
                  End If
                  
                 'insumoCOMBO
                 'INSUMOCRITICO23
                    Dim rstItems As New ADODB.Recordset
                    Set rstItems = New ADODB.Recordset
                    Set rstItems = Lib.OpenRecordset("SELECT     dbo.TPRODUCTO.tCodigoInsumo, dbo.DPEDIDO.nCantidad * dbo.CPEDIDO.nCantidad AS nCantidad FROM         dbo.DPEDIDO INNER JOIN                       dbo.CPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.DPEDIDO.tItem = dbo.CPEDIDO.tItem INNER JOIN                       dbo.TPRODUCTO ON dbo.CPEDIDO.tProductoCombo = dbo.TPRODUCTO.tCodigoProducto WHERE    cpedido.tcodigopedido='" & sPedido & "'  and cpedido.tItem ='" & sItem & "' and tItemCombo='" & xItem & "' and  (dbo.TPRODUCTO.lControlInsumoCritico = 1) AND (ISNULL(dbo.TPRODUCTO.tCodigoInsumo, '') <> '') AND (ISNULL(dbo.CPEDIDO.lImprime, 0) = 1) ", Cn)
                
                    If Not (rstItems.EOF Or rstItems.BOF) Then
                        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
                    End If
    
                   Dim cMax As String
                   cMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
                   cMax = Lib.Correlativo(cMax, 3)
                   Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
                          "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
                          "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
                          "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
                          "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fDiaContable) " & _
                          "select '" & sPedido & "' as tCodigoPedido, '" & cMax & "' as tItem, cpedido.tProductocombo, cpedido.tCodigoGRupo, cpedido.tCodigoSubGrupo, " & _
                          "cpedido.nPrecioNeto, cpedido.nImpuesto1/cpedido.ncantidad, cpedido.nImpuesto2/cpedido.ncantidad, cpedido.nImpuesto3/cpedido.ncantidad, cpedido.nVenta/cpedido.ncantidad, " & _
                          "0, cpedido.nPrecioNeto, cpedido.nCantidad, cpedido.nImpuesto1, cpedido.nImpuesto2, cpedido.nImpuesto3, cpedido.nVenta, '', cpedido.lImprime, " & _
                          "'" & RsCabecera!tUsuario & "' as tUsuario, dpedido.fregistro as fRegistro, " & _
                          "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
                          "'Anulado de Combo' as tObservacion, 'Anul. de Combo:' + t.tResumido as tObservacionAnulado, 'N', cpedido.lImprimeArea, '', '000', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
                          "from  dpedido inner join cpedido on dpedido.tCodigoPedido=cpedido.tcodigopedido and  dpedido.tItem = CPEDIDO.tItem inner join tproducto t on t.tcodigoproducto = dpedido.tCodigoProducto " & _
                          "where cpedido.tCodigoPedido = '" & sPedido & "' and cpedido.tItem ='" & sItem & "' and cpedido.tItemCombo='" & xItem & "'"
                   Cn.Execute Isql
                  
                  
                  Cn.Execute "delete from CPEDIDO where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                  Cn.Execute "delete from TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"

                  RsComboPropiedad.Requery
                  RsCombo.Requery
                  If RsCombo.RecordCount > 0 Then
                     RsCombo.MoveLast
                  End If
                  
               End If
   
          Case Is = 2 'Aumentar
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                 "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'", Cn)
               If nCCombo + 1 > nCombo * RsDetalle!nCantidad Then
                  MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
                  Exit Sub
               End If
               'Oscar Ortega----------------------------------------------
                Dim oRsProductoDeCombo As Recordset
                Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                If oRsProductoDeCombo.RecordCount > 0 Then
                    If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         Dim nCantidadEnElCombo As Integer
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sPedido, sItem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidadEnElCombo >= RsDetalle!nCantidad Then
                             MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             Exit Sub
                         End If
                    End If
                End If
                'If lcover Then
                   Isql = "Select t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                    Dim RstCombo2 As Recordset
                    Set RstCombo2 = Lib.OpenRecordset(Isql, Cn)
                    If RstCombo2.RecordCount > 0 Then
                            
                             Dim maxped As Double
                             maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                                 If maxped > 0 Then
                                        If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (RstCombo2!nAumento)) Then
                                          MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                           RsCombo.Requery
                                           'RsCombo.AbsolutePosition = nPos
                                          Exit Sub
                
                                        End If
                                 End If
                             
                    End If
                'End If
                '----------------------------------------------------------
               nPos = RsCombo.AbsolutePosition
               Cn.Execute "update CPEDIDO set nCantidad = " & RsCombo!nCantidad + 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               'Oscar Ortega------------------------------------------------------------
               Isql = "Select t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
               Set RstCombo = Lib.OpenRecordset(Isql, Cn)
               
               If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta + RstCombo!nAumento / nCantidad, sPedido)
                    End If
               End If
               'Fin Oscar Ortega--------------------------------------------------------
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
               
          Case Is = 3 'Disminuir
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If

               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
          
               If grdCombo.Columns(3).Text > 1 Then
                  nPos = RsCombo.AbsolutePosition
                  Cn.Execute "update CPEDIDO set nCantidad = " & RsCombo!nCantidad - 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                  'Oscar Ortega------------------------------------------------------------
                  Isql = "Select t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & RsCombo!tItemCombo & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                      txtSuma.Caption = CambiaPrecio(nPVenta - RstCombo!nAumento / nCantidad, sPedido)
                    End If
                  End If
                  'Fin Oscar Ortega--------------------------------------------------------
                  RsCombo.Requery
                  RsCombo.AbsolutePosition = nPos
               End If
               
          Case Is = 4 'Cantidad
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If

                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If


               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               sTipo = ""
               frmNumPad.Show vbModal
               If wEnter And val(sDescrip) > 0 Then
                  nCantidad = val(sDescrip)
                  nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                                    "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'", Cn)
                  If nCCombo + nCantidad - RsCombo!nCantidad > nCombo * RsDetalle!nCantidad Then
                     MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
                     nCantidad = 1
                     Exit Sub
                  End If
                  'Oscar Ortega----------------------------------------------
                  Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sCombo)
                  If oRsProductoDeCombo.RecordCount > 0 Then
                     If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                         nCantidadEnElCombo = ObtenerSumaCantidadesEnElComboSinEsteProducto(sPedido, sItem, xItem, oRsProductoDeCombo!tEtiqueta)
                         If nCantidad > RsDetalle!nCantidad - nCantidadEnElCombo Then
                             MsgBox "Solo es permitido " & RsDetalle!nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                             nCantidad = 1
                             Exit Sub
                         End If
                     End If
                  End If
                  
                   nCantidad = RsDetalle!nCantidad '-- NO TOCAR
                  '----------------------------------------------------------
                    'If lcover Then
                        Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                        'Dim RstCombo As Recordset
                        Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                        If RstCombo.RecordCount > 0 And RstCombo!nCantidad < val(sDescrip) Then
                                
                                 Dim maxped2 As Double
                                 maxped2 = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                                     If maxped2 > 0 Then
                                            If maxped2 < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (RstCombo!nAumento * (val(sDescrip) - RstCombo!nCantidad))) Then
                                              MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                               RsCombo.Requery
                                               'RsCombo.AbsolutePosition = nPos
                                              Exit Sub
                    
                                            End If
                                     End If
                                 
                        End If
                    'End If

                  'Oscar Ortega----------------------------------------------Cambia Precio
                  
                  Isql = "Select c.nCantidad, t.nAumento From CPEDIDO As c Left Join TCombo as t On c.tProducto = t.tCombo And c.tProductoCombo = t.tCodigoProducto Where c.tCodigoPedido = '" & sPedido & "' And c.tItem = '" & sItem & "' And c.tItemCombo = '" & xItem & "'"
                  Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                  If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta - ((RstCombo!nAumento / nCantidad) * RstCombo!nCantidad), sPedido)
                        txtSuma.Caption = CambiaPrecio(nPVenta + ((RstCombo!nAumento / nCantidad) * val(sDescrip)), sPedido)
                    End If
                    
                  End If
                  'Fin Oscar Ortega------------------------------------------
                  
                  nPos = RsCombo.AbsolutePosition
                  Cn.Execute "update CPEDIDO set nCantidad = " & val(sDescrip) & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                  RsCombo.Requery
                  RsCombo.AbsolutePosition = nPos
               End If
   
          Case Is = 5 'Propiedad Combos
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
          
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
               If RsCombo!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               ActivaFrame fraPropiedad
   
           Case Is = 6  'Orden +
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
                    
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                Cn.Execute "update CPEDIDO set nOrden = " & nOrd + 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                RsCombo.Requery
                RsCombo.AbsolutePosition = nPos
                
           Case Is = 7  'Orden -
                If RsCombo.RecordCount = 0 Then
                   Exit Sub
                End If
           
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
           
                nPos = RsCombo.AbsolutePosition
                nOrd = IIf(IsNull(RsCombo!nOrden), 0, RsCombo!nOrden)
                If nOrd > 1 Then
                   Cn.Execute "update CPEDIDO set nOrden = " & nOrd - 1 & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
                   RsCombo.Requery
                   RsCombo.AbsolutePosition = nPos
                End If
                
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
          Case Is = 8  'Linea Corte
               If RsCombo.RecordCount = 0 Then
                  Exit Sub
               End If
               
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
               nPos = RsCombo.AbsolutePosition
               If IIf(IsNull(RsCombo!lCorte), False, RsCombo!lCorte) Then
                  Cn.Execute "update CPEDIDO set lCorte = 0 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               Else
                  Cn.Execute "update CPEDIDO set lCorte = 1 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & RsCombo!tItemCombo & "'"
               End If
               RsCombo.Requery
               RsCombo.AbsolutePosition = nPos
                
   End Select

    HabilitaTimerColor (True)
End Sub
 

'--
Private Sub cmdDetalle_Click(Index As Integer)

    HabilitaTimerColor (False)
    Dim maxped As Double
    
   Dim nPos As Integer
   txtBarra.SetFocus
   If RsDetalle.RecordCount = 0 Then
      Exit Sub
   End If

   If RsDetalle!tFacturado = "F" Or RsDetalle!tFacturado = "P" Then
      MsgBox "Producto ya Facturado ", vbExclamation, sMensaje
      Exit Sub
   End If

   Select Case Index
          Case Is = 0 ' Elimina
          
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
                'GCAA 10/07/2023 - REST-012-C-2023
               If lBloqueoAnulaItemsPedido And RsDetalle!lImprime Then
                   MsgBox "Esta opcion se encuentra Bloqueada, favor de comunicar al Supervisor", vbExclamation, sMensaje
                   Exit Sub
               End If

                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible eliminar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If

               If MsgBox("Seguro de Eliminar el Producto?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  sUsuarioAutoriza = sUsuario
                   
                  If (lPassword And lPrinter = False) Or (lPassword And lPrinter And RsDetalle!lImprime) Or (IIf(IsNull(RsCabecera!tCodigoPedidoCD), "", RsCabecera!tCodigoPedidoCD) <> "") Then
                     If Supervisor("02") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                     End If
                     sUsuarioAutoriza = sVar1
                     'Restaura variable sVar1 al valor anterior (al usuario logeado) - FERM 29/10/2024
                     sVar1 = sUsuario
                  End If
                                                                            
                  If (lElimina And lPrinter = False) Or (lElimina And lPrinter = True And RsDetalle!lImprime) Then
                     wCabecera = False
                     ActivaFrame fraEliminacion
                     ActivaCabecera False
                     ActivaOpcion False
                     'oo2
                     ActivarDesactivarCD (False)
                     cmdOpcion(1).Enabled = False
                     cmdOpcion(2).Enabled = False
                  Else
                     sCodigo = ""
                     sDescrip = ""
                    ' EliminaItem
                     'KDS
                     If lKDS Then
                         Call KDS_EliminarProducto(RsCabecera, sItem, 0)
                     End If
                     EliminaItem
                  End If
                  
                
              End If
                   
          Case Is = 1 ' Cantidad
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
          
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If

               
               'GCAA 05022020
                If Calcular("select count(*) as codigo FROM cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tCodigoPedido='" & sPedido & "' and titem='" & sItem & "' and isnull(tcombo.lfijo,0)=0 and isnull(tcombo.naumento,0)>0", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & " Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
                
               sTipo = IIf(lFactor, "Porcion", "")
               frmNumPad.Show vbModal
               If wEnter And val(sDescrip) > 0 Then
                    'Oscar Ortega------------
                    Dim oRsDetalleProducto As Recordset
                    Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                    
                    'JSJR 21/01/2025
                    If Len(IIf(IsNull(oRsDetalleProducto!tOferta), "", oRsDetalleProducto!tOferta)) > 0 Then
                       Dim RsValorOfertaCantidad As Recordset
                       Isql = "exec CalcularStockOfertas '" & oRsDetalleProducto!tCodigoProducto & "', '" & oRsDetalleProducto!tOferta & "',1, '" & oRsDetalleProducto!tTipoPedido & "'"
                       Set RsValorOfertaCantidad = Lib.OpenRecordset(Isql, Cn)
                       If IsNull(RsValorOfertaCantidad!stockOferta) Then
                          MsgBox ("" & RsValorOfertaCantidad!respuesta), vbExclamation
                          Exit Sub
                       End If
                    End If
                    
                    'CESAR ROTULADO
                    If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                       MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                       Exit Sub
                    End If
                    
                   'If lcover Then
                        maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                            If maxped > 0 Then
                                If (nCantidad < CDbl(sDescrip)) Then ' Sumar
                                       If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + CDbl((CDbl(sDescrip) - nCantidad) * nPVenta)) Then
                                         MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                         
        '                                 Screen.MousePointer = vbDefault
        '                                 RsDetalle.MoveLast
        '                                 RsDetalle.Requery
                                         'RsProducto.MoveFirst
                                         Exit Sub
                    
                                       End If
                                Else ' restar
'                                        If maxped < (Val(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + Val((nCantidad - Val(sDescrip)) * nPVenta)) Then
'                                         MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
'
'        '                                 Screen.MousePointer = vbDefault
'        '                                 RsDetalle.MoveLast
'        '                                 RsDetalle.Requery
'                                         'RsProducto.MoveFirst
'                                         Exit Sub
'
'                                       End If
                                End If
                            End If
                    'End If
                    
                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                       If verificaCantidadDeItemsCombos(sPedido, sItem, oRsDetalleProducto!nCombinacion, val(sDescrip)) = False Then
                          MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                          Exit Sub
                       End If
                       
                       Dim Xse As Integer
                       If (nCantidad < val(sDescrip)) Then ' Sumar
                           Xse = val(sDescrip) - nCantidad
                           Dim i As Integer
                           For i = 1 To Xse
                               nPos = RsDetalle.AbsolutePosition
                               nCantidad = nCantidad + 1
                               Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                               If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                                  'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad - 1), sPedido)
                                  'Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sitem & "'"
                                   Dim AAcombo As String
                                   AAcombo = CambiaPrecioCombo(((oRsDetalleProducto!nVenta - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nPrecioOficial) / nCantidad, sPedido)
                               Else
                                  Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                               End If
                           Next i
                       Else 'Restar
                           Dim nCantidadMax As Double
                           nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sPedido, sItem, nCantidad)
                             
                           If nCantidad > nCantidadMax Then
                              Xse = nCantidad - val(sDescrip)
                              For i = 1 To Xse
                                  Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                                        nCantidad = nCantidad - 1
                                        
                                        If nCantidad = 0 Then
                                            MsgBox "NO es posible Hacer este proceso para Combos!!!!"
                                            Exit Sub
                                        End If
                                        'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad + 1), sPedido)
                                        'Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sitem & "'"
                                        Dim DDcombo As String
                                        DDcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad, sPedido)
                                    Else
                                        Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                                    End If
                              Next i
                           Else
                              MsgBox ("No puedes reducir la cantidad de combos con elementos únicos" & Chr(13) & "Disminuya primero la cantidad de productos dentro del combo"), vbExclamation
                              Exit Sub
                              End If
                           End If
                           'Cambiar Combo
                           nCantidad = val(sDescrip)
                           nPos = RsDetalle.AbsolutePosition
                           oRsDetalleProducto.Requery
                           RsDetalle.Requery
                    Else
                        nCantidad = val(sDescrip)
                        nPos = RsDetalle.AbsolutePosition
                        Cn.Execute "update DPEDIDO set nCantidad = " & val(sDescrip) & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1 * " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2 * " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3 * " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                        
                        'agrega al log INFOREST segun REQ REST-024-N-2024    VSJ
                        Call Log_Inforest("PUNTO VENTA", "MODIFICACIÓN DE CANTIDAD DEL ITEM DEL PEDIDO", sPedido, Error, "", "Cantidad: " & val(sDescrip), "Modifió la cantidad del item " & sItem, "Se realizó la modificación de la cantidad del item " & sItem & " - Pedido: " & sPedido, sUsuario)
                    
                    End If
                    
                    '------------------------
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                    fxCombo "M", nCantidad, sProducto
                    'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
                    SumPrecioVenta (sPedido)
                    verificatitulo
                    
               End If
               
          Case Is = 2 ' Aumentar
          
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If

               'If lcover Then
                    
                    maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                        If maxped > 0 Then
                               If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + CDbl(nPVenta)) Then
                                 MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                                 
'                                 Screen.MousePointer = vbDefault
'                                 RsDetalle.MoveLast
'                                 RsDetalle.Requery
                                 'RsProducto.MoveFirst
                                 Exit Sub
            
                               End If
                        End If
                'End If
                'GCAA 2020
                If Calcular("select count(*) as codigo FROM cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tCodigoPedido='" & sPedido & "' and titem='" & sItem & "' and isnull(tcombo.lfijo,0)=0 and isnull(tcombo.naumento,0)>0 ", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
               
               nPos = RsDetalle.AbsolutePosition
               nCantidad = nCantidad + 1
               'Oscar Ortega Aumenta Combo---------------
               Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
               
               'JSJR 21/01/2025
               If Len(IIf(IsNull(oRsDetalleProducto!tOferta), "", oRsDetalleProducto!tOferta)) > 0 Then
                  Dim RsValorOfertaIncrementar As Recordset
                  Isql = "exec CalcularStockOfertas '" & oRsDetalleProducto!tCodigoProducto & "', '" & oRsDetalleProducto!tOferta & "',1, '" & oRsDetalleProducto!tTipoPedido & "'"
                  Set RsValorOfertaIncrementar = Lib.OpenRecordset(Isql, Cn)
                  If IsNull(RsValorOfertaIncrementar!stockOferta) Then
                     nCantidad = nCantidad - 1
                     MsgBox ("" & RsValorOfertaIncrementar!respuesta), vbExclamation
                     Exit Sub
                  End If
               End If

               
               
               If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                            If verificaCantidadDeItemsCombos(sPedido, sItem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                                MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                                Exit Sub
                            End If

                            Dim Acombo As String
                            Acombo = CambiaPrecioCombo(((oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento) + oRsDetalleProducto!nVenta) / nCantidad, sPedido)
                            'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad - 1), sPedido)
                            'Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sitem & "'"
               Else
                    'CESAR ROTULADO
                    If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                       MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                    Else
                       Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                        'agrega al log INFOREST segun REQ REST-024-N-2024    VSJ
                        Call Log_Inforest("PUNTO VENTA", "MODIFICACIÓN DE CANTIDAD DE PEDIDO(+)", sPedido, Error, "valor ant. " & nCantidad - 1, "valor act. " & nCantidad, "Aumentó la cantidad", "Se realizó la modificación de la cantidad del pedido en la tabla DPEDIDO", sUsuario)
                    End If
                    
               End If
               '-----------------------------------------
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
               fxCombo "M", nCantidad, sProducto
               'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
               SumPrecioVenta (sPedido)
               verificatitulo
               
               
          Case Is = 3 ' Disminuir
          
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
                If RsDetalle!lImprime = True Then
                    MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                    Exit Sub
                End If
                
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If

                'GCAA 05022020
                 If Calcular("select count(*) as codigo FROM cpedido inner join tcombo on cpedido.tproducto=tcombo.tcombo and cpedido.tproductocombo=tcombo.tcodigoproducto where cpedido.tCodigoPedido='" & sPedido & "' and titem='" & sItem & "' and isnull(tcombo.lfijo,0)=0 and isnull(tcombo.naumento,0)>0 ", Cn) > 0 Then
                    MsgBox ("No es posible modificar la cantidad!!! " & vbNewLine & "Hay al menos 1 item combo con aumenta precio no fijo que restringue el proceso!!!"), vbInformation
                    Exit Sub
                End If
                If grdDetalle.Columns(4).Text > 1 Then
                    nCantidad = nCantidad - 1
                    nPos = RsDetalle.AbsolutePosition
                    'Oscar Ortega Disminuir Combo---------------
                    Set oRsDetalleProducto = ObtenerDetalleProducto(sPedido, sItem)
                    If (IIf(IsNull(oRsDetalleProducto!lCombinacion), 0, oRsDetalleProducto!lCombinacion)) Then
                        If verificaCantidadDeItemsCombos(sPedido, sItem, oRsDetalleProducto!nCombinacion, nCantidad) = False Then
                            MsgBox "No es posible aplicar los cambios. Verifique la cantidad de productos dentro del combo", vbCritical + vbInformation
                            nCantidad = nCantidad + 1 'oo03
                            Exit Sub
                        End If
                        'Dim nCantidadMax As Double
                        nCantidadMax = Obtener_CantidadMaximaDeUnicoEtiqueta(sPedido, sItem, nCantidad + 1)
                        If nCantidad + 1 > nCantidadMax Then
                            'txtSuma.Caption = CambiaPrecio((oRsDetalleProducto!nVenta) / (nCantidad + 1), sPedido)
                            'Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sitem & "'"
                            Dim Dcombo As String
                            Dcombo = CambiaPrecioCombo((oRsDetalleProducto!nVenta - (oRsDetalleProducto!nPrecioOficial - oRsDetalleProducto!nDescuento)) / nCantidad, sPedido)

                        Else
                            MsgBox ("No puedes reducir la cantidad de combos con elementos únicos" & Chr(13) & "Disminuya primero la cantidad de productos dentro del combo"), vbExclamation
                            nCantidad = nCantidad + 1
                            Exit Sub
                        End If
                    Else
                        'CESAR ROTULADO
                        If oRsDetalleProducto!tCodigoEtiqueta <> "" Then
                           MsgBox "No es posible aplicar los cambios", vbCritical + vbInformation
                        Else
                        
                        Call Log_Inforest("PUNTO VENTA", "MODIFICACIÓN DE CANTIDAD DE PEDIDO( - )", sPedido, Error, "valor ant. " & nCantidad + 1, "valor act. " & nCantidad, "Disminutó la cantidad", "Se realizó la modificación de la cantidad del pedido en la tabla DPEDIDO", sUsuario)
                        Cn.Execute "update DPEDIDO set nCantidad = " & nCantidad & ", nVenta = " & nCantidad * nPVenta & ",nImpuesto1 = nPrecioImpuesto1* " & nCantidad & ", nImpuesto2 = nPrecioImpuesto2* " & nCantidad & ", nImpuesto3 = nPrecioImpuesto3* " & nCantidad & " where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
                        
                        End If
                        
                    End If
                    
                  '-----------------------------------------
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                    fxCombo "M", nCantidad, sProducto
                    'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
                    SumPrecioVenta (sPedido)
                End If
                    
                verificatitulo
                                  

          Case Is = 4 ' Propiedad
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
                If lActivaAnticipo Then
                  If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
                      MsgBox "No es posible modificar el producto, anticipo aplicado al pedido!!!", vbExclamation, sMensaje
                      Exit Sub
                  End If
                End If

               'Oscar Ortega------------------------------
               If Obtener_PropiedadesPendientesPorItems(sProducto, sItem) = False Then
                    grdDetalle.Enabled = False
                    grdCabecera.Enabled = False
                Else
                    grdDetalle.Enabled = True
                    grdCabecera.Enabled = True
               End If
               'Fin Oscar Ortega--------------------------
               ActivaFrame fraPropiedad
               ActivaCabecera False
               ActivaOpcion False
               
          Case Is = 5 ' Observacion
          
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
               If wAgregaCombo Then
                  If RsCombo!lImprime = True Then
                     MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                     Exit Sub
                  End If
               Else
                  If RsDetalle!lImprime = True Then
                     MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                     Exit Sub
                  End If
               End If
          
               frmKeyBoard.Caption = "Observación del Producto"
               frmKeyBoard.txtResultado.Text = lblObservacion.Text
               frmKeyBoard.Show vbModal
               If wEnter = True Then
                  If wAgregaCombo Then
                     nPos = RsCombo.AbsolutePosition
                     Cn.Execute "Update CPEDIDO set tObservacion = '" & sDescrip & "' where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "' and tItemCombo='" & xItem & "'"
                  Else
                     nPos = RsDetalle.AbsolutePosition
                     Cn.Execute "Update DPEDIDO set tObservacion = '" & sDescrip & "' where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "'"
                  End If
                  lblObservacion.Text = sDescrip
               End If
                         
          Case Is = 6 ' Linea
               If RsDetalle.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
               
               nPos = RsDetalle.AbsolutePosition
               If IIf(IsNull(RsDetalle!lCorte), False, RsDetalle!lCorte) Then
                  Cn.Execute "update DPEDIDO set lCorte = 0 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
               Else
                  Cn.Execute "update DPEDIDO set lCorte = 1 where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
               End If
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
                              
          Case Is = 7  'Orden +
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
                
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               nPos = RsDetalle.AbsolutePosition
               nOrden = nOrden + 1
               Cn.Execute "update DPEDIDO set nOrden = " & nOrden & " where tCodigoPedido='" & sPedido & "' and tItem ='" & RsDetalle!tItem & "'"
               Call Log_Inforest("PUNTO VENTA", "MODIFICACIÓN DE ORDEN ( + )", sPedido, Error, "", "Num. Orden " & nOrden, "Modificación en el número de orden(-)", "Se realizó la modificación de la orden asignada", sUsuario)
                    
               RsDetalle.Requery
               RsDetalle.AbsolutePosition = nPos
                
          Case Is = 8  'Orden -
                If lMCPV Then
                    If Not ValidaExistenciaProducto() Then
                        MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                        RsCabecera.Requery
                        RsDetalle.Requery
                        Exit Sub
                    End If
                End If
               If RsDetalle!lImprime = True Then
                  MsgBox "El Item ya fue Impreso", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               nPos = RsDetalle.AbsolutePosition
               
               If nOrden > 1 Then
                  If lOrden Then
                     If Calcular("select max(nOrden) as Codigo from DPEDIDO where tCodigoPedido='" & sPedido & "' and lImprime=1", Cn) >= nOrden - 1 Then
                        If MsgBox("Deseas Juntar el Producto a la Orden " & nOrden - 1 & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                           Exit Sub
                        End If
                     End If
                  End If
                  nOrden = nOrden - 1
                  Cn.Execute "update DPEDIDO set nOrden = " & nOrden & " where tCodigoPedido='" & sPedido & "' and tItem ='" & RsDetalle!tItem & "'"
                  Call Log_Inforest("PUNTO VENTA", "MODIFICACIÓN DE ORDEN (-)", sPedido, Error, "", "Num. Orden " & nOrden, "Modificación en el número de orden(-)", "Se realizó la modificación de la orden asignada", sUsuario)
               
                  RsDetalle.Requery
                  RsDetalle.AbsolutePosition = nPos
               End If
               
   End Select
   
   HabilitaTimerColor (True)
End Sub

Private Sub cmdEliminacion_Click(Index As Integer)
    
    HabilitaTimerColor (False)
    
    RsMotivoEliminacion.MoveFirst
    RsMotivoEliminacion.Find ("Descripcion = '" & cmdEliminacion(Index).Caption & "'")
    
    If RsMotivoEliminacion.EOF Then
       RsMotivoEliminacion.MoveFirst
    End If
    
    If RsMotivoEliminacion!codigo = "000" Then
       frmKeyBoard.txtResultado = ""
       frmKeyBoard.Show vbModal
       If Not wEnter Then
          Exit Sub
       End If
       sCodigo = "000"
    Else
       sCodigo = RsMotivoEliminacion!codigo
       sDescrip = ""
    End If
    
    If wCabecera Then
       EliminaCabecera
    Else
       'KDS
        If lKDS Then
            Call KDS_EliminarProducto(RsCabecera, sItem, 0)
        End If
       EliminaItem
    End If
        
    fraEliminacion.Visible = False
    wDetalle = True ' Revisar para que se coloca en true si estaba en false con los usuarios mozos (multicajeros)
    wAgrega = False
    wAgregaCombo = False
    
    'oo2
    ActivarDesactivarCD (True)
    HabilitaTimerColor (True)
End Sub

Private Sub cmdEnvioFe_Click(Index As Integer)
    Select Case Index
                   
       Case Is = 0  'Exit
            
                      
       Case Is = 1  'Grabar
            Unload Me
    End Select
End Sub
Private Sub cmdGrupo_Click(Index As Integer)

    HabilitaTimerColor (False)
    
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 5
       cmdGrupo(i).backColor = vbButtonFace
       cmdSubGrupo(i).backColor = vbButtonFace
   Next i
   cmdSubGrupo(6).backColor = vbButtonFace
   cmdSubGrupo(7).backColor = vbButtonFace
   
   RsGrupo.MoveFirst
   RsGrupo.Find "nboton = " & Trim(str(Index))
   sGrupo = RsGrupo!codigo
   xGrupo = RsGrupo!codigo
   cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
   RsSubgrupo.Filter = "tGrupo = '" & sGrupo & "'"
   AsignaBoton 7, RsSubgrupo, cmdSubGrupo()
   
   If RsSubgrupo.RecordCount = 0 Then
      sSubGrupo = ""
   Else
      RsSubgrupo.MoveFirst
      RsSubgrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & sGrupo & "' and lActivo=1", Cn) & "'"
      sSubGrupo = RsSubgrupo!codigo
      xSubGrupo = RsSubgrupo!codigo
      If RsSubgrupo!nBoton > 0 Then
         cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
      End If
 
      RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
      AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio
   End If
   Screen.MousePointer = vbDefault
   txtBarra.SetFocus
   
   HabilitaTimerColor (True)
End Sub

Private Sub cmdMesa247_Click()

     '----------------
     Dim CodPedido1, NomMesa As String
     sTemp = ""
      Isql = "exec usp_WS_MESA_Inforest '', '',2"
      Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                "Cliente", 2, "Descripcion", 5800, 0, 0, "")
      frmBusquedaRapida.nPredeterm = 1
      frmBusquedaRapida.tipoBusqueda = "CHECKING"
      frmBusquedaRapida.Show vbModal
    '-------------
        CodPedido1 = sCodigo
        NomMesa = Calcular("select tDetallado as codigo from TMESA where tCodigoMesa='" & sMesa & "'", Cn)
    If CodPedido1 <> "" Then
   
         Dim value As String
         
            If RsCabecera.RecordCount > 0 Then
               value = MsgBox("¿Enlazar Pedido:( " & NomMesa & " / " & sPedido & " )a usuario Mesa24/7?" + vbNewLine + "Si elige No Se generara un nuevo Pedido", vbYesNoCancel)
               'value = vbNo
            Else
               value = vbNo
            End If
            
            If value = vbYes Then
               If NomMesa = "0" Then
                   If lObligaMesa And sMesa = "" Then
                      MsgBox "Asigne una Mesa", vbExclamation, sMensaje
                      cmdCabecera_Click (13)
                      'Exit Sub
                   End If
                   If sMesa = "" Then
                       If MsgBox("No hay Mesa seleccionada, ¿Desea continuar?", vbYesNo) = vbNo Then
                         Me.cmdOpcion(10).Visible = True
                         Exit Sub
                       End If
                   Else
                       Cn.Execute "update mpedido set tmesa='" & sMesa & "' where tcodigopedido='" & sPedido & "'"
                       Cn.Execute "update TMESA set tEstadoMesa ='02' where tCodigoMesa='" & sMesa & "'"
                       sMesa = ""
                   End If
               End If
            
               If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and isnull(idPedidoMesa,'')=''  and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then ' and tTurno='" & sTurno & "'
                   MsgBox "El Pedido ya fue Facturado ó Ya esta Enlazado A Usuario de Mesa24/7", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Me.cmdOpcion(10).Visible = True
                   Exit Sub
               End If
                 Isql = "exec usp_WS_MESA_Inforest '" & sPedido & "', '" & CodPedido1 & "',1"
                 Cn.Execute Isql
               GoTo fin
            Else
                If value = vbCancel Then
                    Exit Sub
                End If
            End If
    
       ''''' OBLIGATORIEDAD DE MESA
        sMesa = ""
        If lObligaMesa And sMesa = "" Then
           MsgBox "Asigne una Mesa", vbExclamation, sMensaje
           cmdCabecera_Click (13)
           'Exit Sub
        End If
        If sMesa = "" Then
            If MsgBox("No hay Mesa seleccionada, ¿Desea continuar?", vbYesNo) = vbNo Then
              GoTo fin:
            End If
        End If
        
        Dim delmax As Integer
        Dim cantPed As Integer
        cantPed = 1
        delmax = 0
        Isql = ""
        
        Isql = "select tCodigoPedido from MPEDIDO where tTurno<>'' and tMozo<>'0000' and tCaja<>'' and idPedidoMesa<>'' and tCodigoPedido= '" + sCodigo + "'"
        Set RsPedido = Lib.OpenRecordset(Isql, Cn)

        If RsPedido.RecordCount = 0 And CodPedido1 <> "" Then
            Isql = "Update Mpedido Set tmesa=  '" + sMesa + "', tTurno = '" + sTurno + "', TCaja = '" + sCaja + "' ,tMozo = '" + sMozo + "' Where tCodigoPedido = '" + CodPedido1 + "'"
            Cn.Execute Isql
            Cn.Execute "update TMESA set tEstadoMesa ='02' where tCodigoMesa='" & sMesa & "'"
            sMesa = ""
       End If
    End If

fin:
       sCodigo = ""
        Me.cmdMesa247.Visible = False
        RsCabecera.Requery
        RsDetalle.Requery
        RsCombo.Requery
        RsComboPropiedad.Requery
        RsProductoPropiedad.Requery
        Me.cmdOpcion(10).Visible = True

End Sub

Private Sub cmdMotorizado_Click(Index As Integer)

    HabilitaTimerColor (False)
    
   RsMotorizado.MoveFirst
   RsMotorizado.Find "nboton = " & Trim(str(Index))
   txtMotorizado.Caption = RsMotorizado!descripcion
   sMotorizado = RsMotorizado!codigo
   
   HabilitaTimerColor (True)
End Sub

Private Sub cmdMozo_Click(Index As Integer)

   HabilitaTimerColor (False)

   RsMozo.Filter = "tResumido <> ''"
   RsMozo.Requery
   RsMozo.MoveFirst
   RsMozo.Find "nboton = " & Trim(str(Index))
   txtMozo.Caption = RsMozo!descripcion
   sMozo = RsMozo!codigo
   
   HabilitaTimerColor (True)
End Sub


Private Sub cmdNavegar_Click(Index As Integer)

    HabilitaTimerColor (False)
        
    txtBarra.SetFocus
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, grdCabecera
           Case Is = 1 'PgUp
                MoverPuntero pgup, grdCabecera
           Case Is = 2 'Previo
                MoverPuntero previo, grdCabecera
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, grdCabecera
           Case Is = 4 'PgDn
                MoverPuntero pgdn, grdCabecera
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, grdCabecera
           Case Is = 6 'Primero
                MoverPuntero Primero, grdDetalle
           Case Is = 7 'PgUp
              grdDetalle.MoveRelative -7
               If grdDetalle.BOF Then
                  grdDetalle.MoveFirst
               End If
           Case Is = 8 'Previo
                MoverPuntero previo, grdDetalle
           Case Is = 9 'Siguiente
                MoverPuntero siguiente, grdDetalle
           Case Is = 10 'PgDn
             grdDetalle.MoveRelative 7
               If grdDetalle.EOF Then
                  grdDetalle.MoveLast
               End If
           Case Is = 11 'Ultimo
                MoverPuntero Ultimo, grdDetalle
           Case Is = 12 'Primero
                MoverPuntero Primero, grdCombo
           Case Is = 13 'Previo
                MoverPuntero previo, grdCombo
           Case Is = 14 'Siguiente
                MoverPuntero siguiente, grdCombo
           Case Is = 15 'Ultimo
                MoverPuntero Ultimo, grdCombo
           Case Is = 16 'PgDn
                MoverPuntero pgdn, grdCombo
           Case Is = 17 'PgUp
                MoverPuntero pgup, grdCombo
    End Select
    
    HabilitaTimerColor (True)
End Sub
Sub Mostrar_frameMotorizado()
'TimerDelivery.Interval = 0
                cmdMesa247.Visible = False
                
                If RsCabecera.RecordCount <> 0 Then
                   If lMCPV Then
                      If Not ValidaExistenciaProducto() Then
                         MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                         RsCabecera.Requery
                         RsDetalle.Requery
                         Exit Sub
                      End If
                   End If

                   ActivaFrame fraCabecera
                   ActivaCabecera False
                   ActivaOpcion False
                   wDetalle = False
                   
                   'FrmOrdenCompra.Visible = True
                   
                   cmdCabecera(13).Enabled = False
                   cmdOpcion(1).Enabled = True
                
                   If lmodificatipoPedido = False Then
                        fraTipoPedido.Enabled = False
                   Else
                        fraTipoPedido.Enabled = True
                   End If
                    
                    'Canales de Venta
                    RsCanalesVenta.MoveFirst
                    RsCanalesVenta.Filter = "Codigo = '" & sTipoPedido & "'"
                    lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
                    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
                    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
                    If lMCPV Then
                        'lObligaMozo = False
                        lActivaMozo = False
                    End If
                    lObligaMozo = IIf(IsNull(RsCanalesVenta!lObligaMozo), False, RsCanalesVenta!lObligaMozo)
                    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
                    lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
                    lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
                    lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
                    lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
                    lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
                        
                    If lActivaMozo Then
                       fraMozo.Visible = True
                       Me.fraMozo.Caption = Mesero
                    Else
                       fraMozo.Visible = False
                    End If
                    
                    If lActivaMotorizado Then
                       fraMotorizado.Visible = True
                    Else
                       fraMotorizado.Visible = False
                    End If
                    Dim j As Integer
                    If lInfhotel Then
                       For j = 1 To 9
                           cmdPunto(j).FontBold = False
                           If cmdPunto(j).Caption = txtPuntoVenta.Caption Then
                              cmdPunto(j).FontBold = True
                           Else
                              cmdPunto(j).FontBold = False
                           End If
                       Next j
                    End If
                End If
                'If lcover Then
                    cmdCabecera(21).Caption = "Monto Max/P : " & Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
                'End If
                
                TimerDelivery.Enabled = False
End Sub
Private Sub ActivarOrigenVentas()
On Error GoTo fin
       Me.fraOrigenVentas.Visible = True
       Isql = "select * from vOrigenVenta where Activo = 1 and Visible = 1 and CodCanalVenta='" & sTipoPedido & "' and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Boton"
       Set RsOrigenVentas = Lib.OpenRecordset(Isql, Cn)
       
       If RsOrigenVentas.RecordCount > 0 Then
            AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
        Else
            AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
            Me.fraOrigenVentas.Visible = False
            If lObligaOrigenVenta = True Then
                MsgBox "Obligatoriedad de origen de ventas:" & vbNewLine & "No hay Origenes de venta visibles configuradas, " & vbNewLine & "Consultar con su administrador!!!", vbInformation
                Me.fraOrigenVentas.Visible = True
            End If
            
            If lActivaMotorizado Then
                fraMotorizado.Visible = True
            Else
                fraMotorizado.Visible = False
            End If
       End If

Exit Sub
fin:
End Sub


Private Sub cmdOpcion_Click(Index As Integer)
 
On Error GoTo fin:
    'Diego L. 30/11/2017
    Dim nPropinaCalc As Double


    HabilitaTimerColor (False)

   cmdOpcion(1).Enabled = True
   Dim i As Integer
   Dim oComando As clsComando
   txtBarra.SetFocus
   Select Case Index
                   
       Case Is = 0  'Exit
       
        If lvisor Then
            Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
        End If
        
            Unload Me
            
       Case Is = 1  'Grabar

            If wDetalle Then
               ActivaOpcion True
               wDetalle = False
               ActivaFrame fraPlato
               ActivaCabecera True
               GrabaProducto
            Else
               'Obligatoriedad de Mozo
               If lObligaMozo Then
                  If lMCPV Then
                      ' Reestablece la variable sVar1 al usuario mozo actual - Revisar - Franz FERM 05/11/2024
                      sVar1 = sUsuario
                      sMozo = ObtenerCodigoMozo(sVar1)
                        
                      If sMozo = "" Then
                         MsgBox "Configure un Cajero-Mesero", vbExclamation, sMensaje
                         Exit Sub
                      End If
                  Else
                      If sMozo = "" Or sMozo = "0000" Then
                         MsgBox "Asigne al " & Mesero, vbExclamation, sMensaje
                         Exit Sub
                      End If
                  End If
               End If
               
               'GCAA 04032021
               'If Calcular("select isnull(lcanaldelivery,0) as codigo from tcanalventa where tcodigocanalventa='" & sTipoPedido & "'", Cn) Then
               '               'Origenes de Venta
                   If lObligaOrigenVenta = True Then
                        'If vOrigenVentas = "" And sTipoPedido = "02" Then
                        ' Validamos que si no se ha seleccionado el origen de venta se agrege el defautl por caja, siempre y cuando sea el canal de venta por default
                        If sTipoPedidoPD = sTipoPedido And (vOrigenVentas = "" Or vOrigenVentas = Null) Then
                            vOrigenVentas = sDefaultOrigenVenta
                        End If
                        If vOrigenVentas = "" Or vOrigenVentas = Null Then
                          MsgBox "Debe seleccionar un origen de venta", vbExclamation, sMensaje
                          Call ActivarOrigenVentas
        '                  cmdCabecera_Click (9)
                          Exit Sub
                        End If
                   End If
'                    If vOrigenVentas = "" Then
'                        MsgBox "No se ha asignado un origen de venta!!", vbExclamation, sMensaje
'                    End If
               'End If
               
               'Obligatoriedad de Motorizado
               If lObligaMotorizado Then
                  If sMotorizado = "" Or sMotorizado = "0000" Then
                     MsgBox "Asigne al Motorizado", vbExclamation, sMensaje
                     Exit Sub
                  End If
               End If
               
'               'Obligatoriedad de Mesa
'               If lObligaMesa And sMesa = "" And txtObservacion.Caption = "" Then
'                  MsgBox "Asigne una Mesa", vbExclamation, sMensaje
'                  cmdCabecera_Click (13)
'                  Exit Sub
'               End If

               'Obligatoriedad de Mesa
               If lObligaMesa Then 'And sMesa = "" And txtObservacion.Caption = "" Then
                    If sMesa = "" Then
                      MsgBox "Asigne una Mesa", vbExclamation, sMensaje
                      cmdCabecera_Click (13)
                      Exit Sub
                    End If
               End If
               
               'Obligatoriedad de Cliente Frecuente
               If sCliente = "" And lObligaClienteFrecuente Then
                  MsgBox "Asigne el Cliente Delivery", vbExclamation, sMensaje
                  cmdCabecera_Click (12)
                  Exit Sub
               End If
               
               'Obligatoriedad de Fecha de Entrega
               If Me.txtFechaProg.Caption = "" And lObligaFechaEntrega Then
                  MsgBox "Asigne la Fecha de Entrega", vbExclamation, sMensaje
                  cmdCabecera_Click (10)
                  Exit Sub
               End If
               
               'entregar A
               If lObligaEntregarA = True And Me.txtEntregar.Caption = "" Then
                  MsgBox "Asigne información en Entregar A", vbExclamation, sMensaje
                  cmdCabecera_Click (9)
                  Exit Sub
               End If
               
'               'Origenes de Venta
'               If lObligaOrigenVenta = True Then
'                    If vOrigenVentas = "" And sTipoPedido = "02" Then
'                      MsgBox "Asigne origen de venta", vbExclamation, sMensaje
'    '                  cmdCabecera_Click (9)
'                      Exit Sub
'                    End If
'               End If
               
               
               ActivaOpcion True
               ActivaFrame fraPlato
               ActivaCabecera True
                              
               If wAgrega Then
                  'Actualiza el estado de la mesa
                  If wMesa Then
                     Isql = "Update TMESA set tEstadoMesa='02' where tCodigoMesa ='" & sMesa & "'"
                     Cn.Execute Isql
                  Else
                     sMesa = ""
                  End If
                  wMesa = False
                  wAgrega = False
                                    
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spIns_MPEDIDO", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  
                  ValidaModifica
                  
                  oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sCliente
                  oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                  oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, IIf(txtTipoPedido.ForeColor = &H800000, 0, 1)
                  oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, sTipoAtencion
                  oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, sMesa
                  oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, sMozo
                  oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, sMotorizado
                  oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                  oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 2, sSalon
                  oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
                  oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtObservacion.Caption
                  oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, Right(sUsuario, 15)
                  oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 10, val(txtAdulto.Text)
                  oComando.CreateParameter "@nNino", adInteger, adParamInput, 10, val(txtNino.Text)
                  oComando.CreateParameter "@nMesa", adInteger, adParamInput, 10, val(txtJuntar.Text)
                  oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, sPuntoVenta
                  oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, sHabitacion
                  oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, sReserva
                  oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, sPasajero
                  oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, sCompania
                  oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, sContacto
                  oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
                  oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
                  oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, IIf(sCodigoDescuento = "000", sDescripcionDescuento, "")
                  oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, Right(tAutorizaDescuento, 15)
                  oComando.CreateParameter "@nTiempoDelivery", adInteger, adParamInput, 10, nTiempoDelivery
                  oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, sTienda
                  oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
                  oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:mm"))
                  'invitado2013
                  oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
                     
                  'pariente2013
                  oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado

                  'entregara
                  oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
                  
                  oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, val(txtEnvioAntes.Text)
                  
                  oComando.CreateParameter "@nMontoMaximo", adVarChar, adParamInput, 250, xMontoMaximo
                  
                  oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, sPedido

                  'origen de ventas poner validacion
                  If vOrigenVentas = Null Or vOrigenVentas = "" Then
                    vOrigenVentas = "00"
                  End If
                
                   oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
                   oComando.CreateParameter "@tTotemMesa", adVarChar, adParamInput, 10, ""
                  
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  Else
                     sPedido = oComando.GetParameterValue("@tPedido")
                  End If
                  
                  Cn.Execute " update mpedido set tordenCompra='" & Me.txtOrdenCompra.Text & "' where tCodigoPedido='" & sPedido & "'"
                                                                                     
                  RsCabecera.Requery
                  RsCabecera.MoveFirst
                  RsCabecera.Find "Codigo = '" & sPedido & "'"
               Else
               
                  If wMesa Then
                     Isql = "Update TMESA set tEstadoMesa='02' where tCodigoMesa ='" & sMesa & "'"
                     Cn.Execute Isql
                  End If
                  wMesa = False
                  wAgrega = False
               
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spUpd_MPEDIDO", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                       
                       'Valida si hay datos a modificar, para guardar la modificación en el LOG_INFOREST
                        ValidaModifica
                        
                        If sCodigoDescuento <> "" And ValidaDesc Then
                           CalculaDescuento
                           RsDetalle.Requery
                           wCalcula = False
                        Else
                           sCodigoDescuento = ""
                        End If
                        
                        oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sCliente
                        oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, sTipoPedido
                        oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, IIf(txtTipoPedido.ForeColor = &H800000, 0, 1)
                        oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, sTipoAtencion
                        oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, sMozo
                        oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, sMotorizado
                        oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtObservacion.Caption
                        oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
                        oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, sPuntoVenta
                        oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, sHabitacion
                        oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, sReserva
                        oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, sPasajero
                        oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, sCompania
                        oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, sContacto
                        oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, xDescuento
                        oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, sCodigoDescuento
                        oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, IIf(sCodigoDescuento = "000", sDescripcionDescuento, "")
                        oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, Mid(tAutorizaDescuento, 1, 15)
                        oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, sTienda
                        oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:mm"))
                        'invitado2013
                        oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, sCodigoInvitado
                        'invitado2013
                           
                        'pariente2014
                        oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, sCodigoParienteSeleccionado
                        'pariente2013
                        'oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 10, Format(txtFechaProg.Caption, "dd/MM/yyyyy HH:nn")
                        'entregara
                        oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20))
                  
                        oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, val(txtEnvioAntes.Text)
                        
                        oComando.CreateParameter "@nMontoMaximo", adVarChar, adParamInput, 250, xMontoMaximo
                  
                        oComando.CreateParameter "@tPedido", adVarChar, adParamInput, 10, sPedido
                        
                        'origen de ventas poner validacion
                        If vOrigenVentas = Null Or vOrigenVentas = "" Then
                            vOrigenVentas = "00"
                        End If
                        oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, vOrigenVentas
                        oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, sMesa
                        oComando.CreateParameter "@tTotemMesa", adVarChar, adParamInput, 10, ""
                      
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  
                  Call Log_Inforest("PUNTO VENTA", "MODIFICACIÓN DE CANTIDAD DE PEDIDO( + )", sPedido, Error, "valor ant. " & nCantidad - 1, "valor act. " & nCantidad, "Aumentó la cantidad", "Se realizó la modificación de la cantidad del pedido en la tabla DPEDIDO", sUsuario)
                  
                  Cn.Execute " update mpedido set tordenCompra='" & Me.txtOrdenCompra.Text & "' where tCodigoPedido='" & sPedido & "'"
            
                  RsCabecera.Requery
                  RsCabecera.Find "Codigo = '" & sPedido & "'"
                  
                  'MESA247
                  If MESA247 Then
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                          Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                    End If
                  End If
                    
               End If
            End If
            

            ActivaOpcion True
            grdCabecera.Enabled = True
            
 
            If lvisor Then
              'Isql = "delete from infovisor"
                Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
              'Isql = "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sPedido & "',2)"
                Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sPedido & "',2)"
            End If
            
            'TimerDelivery.Interval = 1000
       Case Is = 2  'Cancelar
            'TimerDelivery.Enabled = True
            ActivaOpcion True
            
            'TimerDelivery.Interval = 1000
            wMesa = False
            
            'Juntar Mesa
            Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='')"
            Cn.Execute "delete from TPEDIDOMESA where tCodigoPedido=''"
            
            If wDetalle Then
               wDetalle = False
               ActivaFrame fraPlato
               ActivaCabecera True
            Else
               If RsCabecera.RecordCount = 0 Then
                  fraCabecera.Visible = False
                  ActivaCabecera False
                  cmdCabecera(0).Visible = True
               Else
                  ActivaCabecera True
                  ActivaFrame fraPlato
               End If
               AsignaPedido
               wAgrega = False
            End If
            grdCabecera.Enabled = True
            
            
            If lvisor Then
               Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
            End If
            
            
       Case Is = 3  'Imp.Pedido
       
       If (lVistaPreviaItemsCombo) Then
        Dim cadena As String
        Dim Itemx As ListItem
        cadena = "usp_TicketComandaVistaPrevia  '" & sPedido & "' "
        Set rsItemCombo = Lib.OpenRecordset(cadena, Cn)
        If Not (rsItemCombo.EOF Or rsItemCombo.BOF) Then
                           
                With lstCombo
                    .View = lvwReport
                    .ListItems.Clear
                    .ColumnHeaders.Clear
                End With
                                 
                Me.MousePointer = vbHourglass
                lstCombo.ColumnHeaders.Add , , "", 4000
                
                rsItemCombo.MoveFirst
                lstCombo.ListItems.Clear
                
                While Not rsItemCombo.EOF
                    Set Itemx = lstCombo.ListItems.Add(, , rsItemCombo.Fields(2) & " - " & rsItemCombo.Fields(3))
                    rsItemCombo.MoveNext
                Wend
                
                If rsItemCombo.EOF Then
                    lstCombo.ListItems.Add , , "-"
                End If
                Me.MousePointer = vbDefault
                'Frame3.Visible = True
                FrameComboItem.Visible = True
         
        End If
     
        If rsItemCombo.RecordCount > 0 Then
           If MsgBox("Despues de Revisar la Vista Previo del Pedido, Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                FrameComboItem.Visible = False
                Exit Sub
           End If
        End If
     
     End If
    
    
'           MsgBox (RsCabecera!tTipoPedido)
'           MsgBox (RsCabecera!tMotorizado)
'
    If RsCabecera.RecordCount = 0 Then
       Exit Sub
    End If
    
    
    If Calcular("select isnull(lObligaMotorizado,0) as codigo from tcanalventa where tcodigocanalventa='" & sTipoPedido & "'", Cn) Then
        If RsCabecera!tTipoPedido = "02" And RsCabecera!tTiporecepcion = 1 Then
            If RsCabecera!tMotorizado = "0000" Then
                MsgBox "Primero debe Asignar una Motorizado", vbExclamation, sMensaje
                Mostrar_frameMotorizado
                Exit Sub
            End If
        End If
    End If
    
            
            
         
         
            
    If lBuscarPedidoBarraMesero Then
        If sMesa = "" Then
           MsgBox "Asigne una Mesa", vbExclamation, sMensaje
           'cmdCabecera_Click (13)
           Exit Sub
        End If

          If lMCPV Then
              sMozo = ObtenerCodigoMozo(sVar1)
              If sMozo = "" Then
                 MsgBox "Configure un Cajero-Mesero", vbExclamation, sMensaje
                 Exit Sub
              End If
          Else
              If sMozo = "" Or sMozo = "0000" Then
                 MsgBox "Asigne al " & Mesero, vbExclamation, sMensaje
                 Exit Sub
              End If
          End If
    End If
    ' Miguel Mendoza - 10062019
    If Not VerificaComboMaximo() Then Exit Sub
    ' *************************
    If lMCPV Then
        If Not ValidaExistenciaProducto() Then
            MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
            RsCabecera.Requery
            RsDetalle.Requery
            Exit Sub
        End If
    End If
    
    If RsArea.RecordCount = 0 Then
       MsgBox "No existe area definida", vbInformation, sMensaje
       Exit Sub
    End If
            
     If lEnvioAutomatico Then
         Dim nMinutosEnvioAntes As Integer
         nMinutosEnvioAntes = Calcular("Select ISNULL(nMinutosAntesEnvio,0) as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
         If nMinutosEnvioAntes <> 0 Then
            MsgBox "El Pedido tiene una programacion de impresión automática", vbInformation, sMensaje
            Exit Sub
         End If
     End If

     If lObservacion And Trim(txtObservacion.Caption) = "" Then
        MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
        cmdCabecera_Click (5)
        If Trim(txtObservacion.Caption) = "" Then
           Exit Sub
        End If
     End If
     
     'ENVIO DE PEDIDO AUTORIZADO
     Dim xClaveAutoriza As String
     Dim xUsuarioAutorizaenvio As String
     Dim xPerfilUsusario As String
     Dim cantProductoAutoriza As Integer
     
     xPerfilUsusario = ""
     xUsuarioAutorizaenvio = ""
     xClaveAutoriza = ""
            
            If lEnvioProduccionUsuario Then
            
            cantProductoAutoriza = Calcular("select count(*) as codigo from DPEDIDO d inner join TPRODUCTO t ON d.tCodigoProducto = t.tCodigoProducto where d.tCodigoPedido='" & sPedido & "' and d.lImprime <> 1 and t.lSolicitaAutorizacionEnvio = 1", Cn)
            
                 If cantProductoAutoriza > 0 And lEnvioProduccionCaja Then
                     frmPassword.Show vbModal
                     If wEnter Then
                         xClaveAutoriza = sDescrip
                         
                         If xClaveAutoriza <> "" Then
                            
                             xUsuarioAutorizaenvio = VerificarAutorizacionEnvio(xClaveAutoriza)
                             
                             If xUsuarioAutorizaenvio = "" Then
                                 Exit Sub
                             Else
                                xPerfilUsusario = Mid(xUsuarioAutorizaenvio, 1, 1)
                                 xUsuarioAutorizaenvio = Mid(xUsuarioAutorizaenvio, 2, Len(xUsuarioAutorizaenvio))
                             End If
                             
                         Else
                             Exit Sub
                         End If
                     Else
                        Exit Sub
                     End If
                     
                End If
    
            End If
            '------------------------
            'Oscar Ortega----------------------------------------------------
            If ExistenPropiedadesPendientesEnPedido(sPedido) Then
                If ExistenPropiedadesPendientesEnCombos(sPedido) Then
                    '----------------------------------------------------------------
                    Screen.MousePointer = vbHourglass
                    
                    'InsumosCriticos ' 23
                    If lPrinter = True Then
                                Dim cmdInsumo          As ADODB.Command
                                Dim X As Integer
                                Dim rstItems As New Recordset
                                Dim cadenaInsumos As String
                                Dim cadenaAEnviar As String
                                Dim resultado As String
                                'cadenaInsumos = "select sum(ncantidad) ncantidad, TPRODUCTO.TCODIGOINSUMO from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto INNER JOIN  dbo.TINSUMO ON dbo.TPRODUCTO.tcodigoInsumo = dbo.TINSUMO.tcodigo  where tcodigopedido='" & sPedido & "' and lcontrolinsumocritico=1 and isnull(limprime,0)=0 AND ISNULL(TCODIGOINSUMO,''  )<>'' and (tinsumo.lactivo=1) group by tcodigoinsumo"
                                cadenaInsumos = "usp_Inforest_ObtenerInsumosCriticos '" & sPedido & "'"
                                Set rstItems = Lib.OpenRecordset(cadenaInsumos, Cn)
                                    If Not (rstItems.EOF Or rstItems.BOF) Then
                                         rstItems.MoveFirst
                                         For X = 0 To rstItems.RecordCount - 1
                                            cadenaAEnviar = cadenaAEnviar + rstItems.Fields(1) + "|" + str(rstItems.Fields(0)) + "$"
                                            rstItems.MoveNext
                                         Next X
                                       
                                        Set cmdInsumo = New ADODB.Command

                                        With cmdInsumo
                                             .ActiveConnection = Cn
                                             .CommandType = adCmdStoredProc
                                             .CommandText = "USP_actualizaStockInsumo"
                                             .Parameters.Refresh
                                             .Parameters("@vi_detalles") = cadenaAEnviar
                                             .Parameters("@vi_numdet") = rstItems.RecordCount
                                             .Parameters("@vch_Salida") = ""
                                        End With
                                        cmdInsumo.Execute
                                        resultado = cmdInsumo.Parameters("@vch_Salida").value
                                        If resultado <> "1" Then:  MsgBox "No hay cantidad disponible de : " & resultado: MoverPuntero Primero, grdDetalle:    Screen.MousePointer = vbDefault: Exit Sub
                                    End If
                    End If
                    'InsumosCriticos

                    'CESAR IMPRESION ALERTA COMBO
                    Dim RsCombox As Recordset
                    Dim CountComb As Integer
                        If lOrden Then
                           Isql = "select vpedido.*, isnull(dpedido.lcombinacion,0) lcombinacion FROM dbo.vPedido INNER JOIN dbo.DPEDIDO ON dbo.vPedido.Codigo = dbo.DPEDIDO.tCodigoPedido AND dbo.vPedido.tItem = dbo.DPEDIDO.tItem " & _
                                  "Where vpedido.Codigo = '" & sPedido & "' and vpedido.nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & sPedido & "' And tArea IS NULL and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))) " & _
                                  "ORDER BY vpedido.nOrden, vpedido.tItem, vpedido.nOrdenCombo,vpedido.tetiqueta,vpedido.combo " ' tItemCombo"
                        Else
                           Isql = "select vpedido.*, isnull(dpedido.lcombinacion,0) lcombinacion FROM dbo.vPedido INNER JOIN dbo.DPEDIDO ON dbo.vPedido.Codigo = dbo.DPEDIDO.tCodigoPedido AND dbo.vPedido.tItem = dbo.DPEDIDO.tItem " & _
                                  "Where vpedido.Codigo = '" & sPedido & "' And vpedido.lImprimeArea = 1 And vpedido.tArea IS NULL and (vpedido.lImprime = 0 or (isnull(vpedido.lImprimeAreaCombo,0) = 1  and isnull(vpedido.lImprimeCombo,0) = 0 ))" & _
                                  "ORDER BY vpedido.nOrden, vpedido.tItem, vpedido.nOrdenCombo,vpedido.tetiqueta,vpedido.combo" 'tItemCombo"
                        End If
                    Set RsCombox = Lib.OpenRecordset(Isql, Cn)
                    
                    If RsCombox.RecordCount > 0 Then
                       If IIf(IsNull(RsCombox!Combo), "", RsCombox!Combo) = "" And RsCombox!lCombinacion = 0 Then
                            MsgBox "El Producto " & RsCombox!Producto & " no tiene un area de impresion configurada", vbExclamation, sMensaje
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        
                                
                       CountComb = Calcular("Select count(*) As Codigo From CPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
                       If CountComb > 0 Then
                          MsgBox "El Producto Combo " & RsCombox!Producto & " contiene Items sin un area de impresion configurada", vbExclamation, sMensaje
                       Else
                          MsgBox "El Producto Combo " & RsCombox!Producto & " no contiene ningun Item como detalle", vbExclamation, sMensaje
                       End If
                       Screen.MousePointer = vbDefault
                       Exit Sub
                       
                    End If
                    '--------------------
                    
                   
                    
                    If lOrden Then
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & sPedido & "' and nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & sPedido & "' and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))) " & _
                              "ORDER BY nOrden, tItem, nOrdenCombo,tetiqueta,combo " ' tItemCombo"
                        
                    Else
                       Isql = "select * from vPedido " & _
                              "Where Codigo = '" & sPedido & "' And lImprimeArea = 1 and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))" & _
                              "ORDER BY nOrden, tItem, nOrdenCombo,tetiqueta,combo" 'tItemCombo"
                    End If
                    
                    
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                    If Not RsImpresion.EOF Then
                        RsArea.MoveFirst
                        For i = 1 To RsArea.RecordCount
                            If lImpComandaBarra Then
                                Dim RsAreaPedidoItem As Recordset
                                Set RsAreaPedidoItem = Lib.OpenRecordset(" select isnull(titem,'000') as tItem from vPedido where Codigo = '" & sPedido & "' And lImprimeArea = 1 group by titem", Cn)
                                Dim Item As Integer
                                For Item = 1 To RsAreaPedidoItem.RecordCount
                                    RsImpresion.Filter = "tArea='" & RsArea!tArea & "' And tItem='" & RsAreaPedidoItem!tItem & "'"
                                    If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                                        If RsImpresion.RecordCount <> 0 Then
                                            RsImpresion.MoveFirst
                                            If lImpComandaBarra Then
                                                Set ImageBarra = Nothing
                                                'BarraCodigo RsAreaPedidoItem!tItem + Right(sPedido, 8), ImageBarra
                                                ImageBarra.Picture = LoadPicture(CrearImagenQR_Comanda("P|" + sPedido + "|" + RsAreaPedidoItem!tItem + "|" + sCaja + "|" + RsArea!tArea + "|" + Format(RsImpresion!fRegistro, "yyyyMMdd HH:mm:ss")))
                                                'Exit Sub
                                            End If
                                            ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest", ImageBarra
                                            'If (lCompatibilidadTVS) Then 'TVS
                                            '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Nuevo Ticket", sPedido)
                                            'End If
                                        End If
                                    End If
                                    RsAreaPedidoItem.MoveNext
                                Next Item
                            Else
                                RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
                                If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                                    If RsImpresion.RecordCount <> 0 Then
                                        RsImpresion.MoveFirst
                                        ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                                        'If (lCompatibilidadTVS) Then 'TVS
                                        '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Nuevo Ticket", sPedido)
                                        'End If
                                    End If
                                End If
                            
                            End If
                            RsArea.MoveNext
                        Next i
                    End If
                    
                    'KDS
                    If lKDS And lKDSInforest = False Then
                       Call KDS_AnadirNuevaOrden(RsCabecera, 0)
                    End If
                    
                    
                   '---------------------------------
                   '--- GCAA 23112021 KDS INFOREST
                   '---------------------------------
                   If lKDSInforest Then
                        'ACTUALIZA
                        Dim oComando_ As clsComando
                        Set oComando_ = New clsComando
                        
                        If Not oComando_.CreateCmdSp("KDS_Usp_spUpdate_DPEDIDO_Orden", Cn) Then
                           Set oComando_ = Nothing
                           Exit Sub
                        End If
                   
                        If lKDSxArea Then
                            oComando_.CreateParameter "@tipo", adVarChar, adParamInput, 10, 2
                        Else
                            oComando_.CreateParameter "@tipo", adVarChar, adParamInput, 10, 1
                        End If
                        
                        oComando_.CreateParameter "@pedido", adVarChar, adParamInput, 10, sPedido
                        oComando_.CreateParameter "@tturnoKDS", adVarChar, adParamInput, 10, sPedido
                        
                        If Not oComando_.GetParamOK Then
                           Set oComando_ = Nothing
                           Exit Sub
                        End If
                    
                        If Not oComando_.ExecSP Then
                           Set oComando_ = Nothing
                           Exit Sub
                         End If
                   End If
                   '---------------------------------
                   '---------------------------------
                    
                    
                    
                    
                    
                    
                    
                    
                    'CESAR----CHEF CONTROL
                    Dim ChefEnvio As Boolean
                    ChefEnvio = Calcular("select ISNULL(lEnvioChef,0) as Codigo FROM TPARAMETRO", Cn)
                    
                    Cn.Execute "Update DPEDIDO Set lNoCantado=0 where tCodigoPedido = '" & sPedido & "' and lImprime<>1"
                    
                    If ChefEnvio Then
                        Cn.Execute "Update DPEDIDO Set lCantadoc=1,fCantadoC=GetDate(), lTipoEnvio=0 where tCodigoPedido = '" & sPedido & "' and lImprime<>1"
                    End If
                
                    Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1  where tCodigoPedido='" & sPedido & "'"
                    Cn.Execute "Update DPEDIDO Set lImprime = 1, tUsuarioAutorizaEnvio = '" & xUsuarioAutorizaenvio & "', tPerfilUsuarioAutoriza = '" & xPerfilUsusario & "', fEnvio = getdate(), nEnvio = 1 where tCodigoPedido = '" & sPedido & "' and lImprime<>1"
                    Cn.Execute "Update CPEDIDO Set lImprime = 1 where tCodigoPedido = '" & sPedido & "'"
                End If
            End If
            Set ImageBarra = Nothing
            RsDetalle.Requery
            RsCombo.Requery
            LimpiaRs
            Me.Caption = muestra

             If lvisor Then
                  Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
                  Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sPedido & "',0)"
            End If

       Case Is = 4  'Pre cuenta
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            
            ' Miguel Mendoza - 10062019
            If Not VerificaComboMaximo() Then Exit Sub
            ' *************************
            
            If lMCPV Then
                If Not ValidaExistenciaProducto() Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
            End If
                        
            If lObservacion And Trim(txtObservacion.Caption) = "" Then
               MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
               cmdCabecera_Click (5)
               If Trim(txtObservacion.Caption) = "" Then
                  Exit Sub
               End If
            End If
            
            Dim nLP As Integer
            nLP = Calcular("select nPrecuenta as Codigo FROM MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
            If nLimitePrecuenta <> 0 And nLimitePrecuenta <= nLP Then
               If Supervisor("14") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
                        
            If lPrecuentaImpresora Then
               frmPrecuentaImpresora.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
            Else
               sCodigo = sPreCuenta
            End If
            
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
               If i > 0 Then
                  MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            
            'Chequea Descuento
            If CDbl(TxtDescuento) > 0 Then
               Dim nTope As Double
               Dim nTotalMes As Double
               
               Dim nConsumo As Double
               Dim aplicaTope As Boolean
               
               nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
               If nTope > 0 Then
                  If Calcular("select lTopePedido as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) Then
                     If CDbl(TxtDescuento) > nTope Then
                        If MsgBox("El Descuento a aplicar Supera El Tope Registrado por Pedido" & Chr(13) & "¿Desea aplicar el Tope de " & sMonN & " " & nTope & "?", vbQuestion + vbYesNo) = vbYes Then
                           CalculaAplicaTope (nTope)
                        Else
                           Exit Sub
                        End If
                     End If
                  Else
                     nTotalMes = Calcular("select sum(DPEDIDO.nDescuento*nCantidad) as Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                                          "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'", Cn)
                     
                     If CDbl(TxtDescuento) + nTotalMes > nTope Then
                        If CDbl(TxtDescuento) < nTope Then
                           If MsgBox("El Descuento a aplicar Supera El Tope Registrado dentro de un mes" & Chr(13) & "¿Desea aplicar el saldo " & sMonN & " " & nTope - nTotalMes & "?", vbQuestion + vbYesNo) = vbYes Then
                              CalculaAplicaTope (nTope - nTotalMes)
                           Else
                              Exit Sub
                           End If
                        Else
                            MsgBox "El Descuento a aplicar Supera El Tope Registrado dentro de un mes", vbExclamation
                            Exit Sub
                        End If
                     End If
                  End If
               End If
               TxtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
            End If
            
            wEnter = False
            Screen.MousePointer = vbHourglass
            Cn.Execute "update MPEDIDO set nPrecuenta = isnull(nPrecuenta,0) + 1  where tCodigoPedido='" & sPedido & "'"
            
            
            'CESAR VALIDACION DELIVERY
            lCanalDelivery = Calcular("select v.lCanalDelivery As Codigo from vTipoPedido v INNER JOIN MPEDIDO m ON Codigo = m.tTipoPedido Where tCodigoPedido='" & sPedido & "'", Cn)
            '------------------------------------
            xTiporecepcion = Calcular("Select isnull(tTiporecepcion,0) as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
            
            If lCanalDelivery Or (WEBAP = True And xTiporecepcion <> 0) Then
               Isql = "select t1.*, t2.fProgramacion from vPreCuentaDelivery T1 inner join MPEDIDO t2 on t1.Codigo = t2.tCodigoPedido  WHERE T1.Codigo='" & sPedido & "' ORDER BY T1.tItem, T1.tItemCombo"
            Else
               If lPrecuenta Then 'Precuenta con mozos
                  Isql = "select * from vPrecuentaDetallada WHERE Codigo='" & sPedido & "'"
               Else
                  If lPrecuentaAgrupada Then
                     Isql = "select * from vPrecuentaAgrupada WHERE Codigo='" & sPedido & "' order by tItem"
                  Else
                     Isql = "select * from vPrecuenta WHERE Codigo='" & sPedido & "' order by tItem"
                  End If
               End If
            End If
            
            Set rstFuente = New ADODB.Recordset
            imageCab.Picture = Nothing
            imagepIE.Picture = Nothing
            Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
            If lImprimeImagCabPrecuenta Then
                imageCab.DataField = "foto"
                Set imageCab.DataSource = rstFuente
            End If
            If lImprimeImagPiePrecuenta Then
                    imagepIE.DataField = "fotoPie"
                    Set imagepIE.DataSource = rstFuente
            End If
            
            If lBuscarPedidoBarraMesero Then
                BarraCodigo sPedido, ImageBarra
                'Exit Sub
            End If
          
            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
               If lCanalDelivery Or (WEBAP = True And xTiporecepcion <> 0) Then
                  ImprimeDelivery RsImpresion, RsProductoPropiedad, RsComboPropiedad, imageCab, imagepIE, ImageBarra
               Else
                  If lPrecuenta Then
                     ImprimePreCuentaDetallada RsImpresion, sCodigo, imageCab, imagepIE, ImageBarra
                  Else
                     If lInfhotel Then
                        ImprimeInfhotel RsImpresion, sCodigo, imageCab, imagepIE
                     Else
                        If lPrecuentaNoValorizada Then
                            ImprimePrecuentaNoValorizada RsImpresion, sCodigo, imageCab, imagepIE, ImageBarra
                        Else
                            'ImprimePreCuentaingles RsImpresion, sCodigo, imageCab, imagepIE
                            ImprimePreCuenta RsImpresion, sCodigo, imageCab, imagepIE, ImageBarra
                        End If
                     End If
                  End If
               End If
               Cn.Execute "update MPEDIDO set fRegCuenta = getDate()  where tCodigoPedido='" & sPedido & "'"
               
               If grdDetalle.EOF Then
                grdDetalle.MoveFirst
               End If
            End If
            LimpiaRs
                                
       Case Is = 5  'Pago
            If RsTipoDocumento.RecordCount = 0 Then
               Exit Sub
            End If
                        
            Screen.MousePointer = vbHourglass
            sVar1 = sTipoPedido
            frmDocumento.Show vbModal
            RsCabecera.Requery
            RsDetalle.Requery
            RsCombo.Requery
            If RsCabecera.EOF Then
               LimpiaCabecera
            End If
                                                
       Case Is = 6  'Código Directo
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            
            If lMCPV Then
                If Not ValidaExistenciaProducto() Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
            End If
            
            sTemp = ""
            
            If Calcular("Select count(*) as Codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' ", Cn) = 0 Then
                
                If sTipoPedido = "01" Then
                         Isql = "select Grupo, Descripcion , nPrecioVenta , nBoton, SubGrupo, Codigo from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "' )and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                ElseIf sTipoPedido = "02" Then
                         Isql = "select Grupo, Descripcion , nPrecioDelivery As nPrecioVenta , nBoton, SubGrupo, Codigo from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "' )and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                ElseIf sTipoPedido = "03" Then
                         Isql = "select Grupo, Descripcion , nPreciollevar As nPrecioVenta , nBoton, SubGrupo, Codigo from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "' )and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                ElseIf sTipoPedido = "04" Then
                         Isql = "select Grupo, Descripcion , nPrecioCanal4 As nPrecioVenta , nBoton, SubGrupo, Codigo from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "' )and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                Else
                         Isql = "select Grupo, Descripcion , nPrecioCanal5 As nPrecioVenta , nBoton, SubGrupo, Codigo from vProducto where lActivo = 1 and (tUnidadNegocio='" & sUnidadNegocio & "' )and " & IIf(sTipoPedido = "01", "lLocal=1", IIf(sTipoPedido = "02", "lDelivery=1", IIf(sTipoPedido = "03", "lLlevar=1", IIf(sTipoPedido = "04", "lcanal4=1", "lcanal5=1")))) & " Order by Descripcion"
                End If
                frmBusquedaRapida.cmdOpcion(4).Visible = False
            Else
                Isql = "select vProducto.Grupo, vProducto.Descripcion , tclienteproducto.nprecio As nPrecioVenta , vProducto.nBoton, vProducto.SubGrupo, vProducto.Codigo from vProducto inner join tclienteproducto on vProducto.codigo=tclienteproducto.tcodigoproducto where vProducto.lActivo = 1 and (vProducto.tUnidadNegocio='" & sUnidadNegocio & "' ) and tclienteproducto.tcodigodelivery='" & sCliente & "' Order by vProducto.Descripcion"
                frmBusquedaRapida.cmdOpcion(4).Visible = True
                frmBusquedaRapida.cmdOpcion(4).FontBold = True
            End If

            Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Grupo", 2, "Grupo", 1600, 0, 0, "", _
                                                            "Producto", 2, "Descripcion", 3600, 0, 0, "", _
                                                            "Precio", 2, "nPrecioVenta", 1000, 1, 0, "###,##0.00", _
                                                            "Bot", 2, "nBoton", 500, 1, 0, "", _
                                                            "SubGrupo", 2, "SubGrupo", 1500, 0, 0, "")
            frmBusquedaRapida.nPredeterm = 1
            frmBusquedaRapida.cmdOpcion(0).Visible = True
            frmBusquedaRapida.Show vbModal
   
            If wEnter Then
                   
                'PROGRAMACION PRECIOS
                 '' REQ REST-049-C-2022 - Osaka Argentina
                 If lActivarIncrementoPrecios Then
                     If (programacionPrecios("Venta")) Then
                         Unload Me
                         Exit Sub
                     End If
                 End If
               
               
                If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
                   MsgBox "El Pedido ya fue facturado", vbExclamation, sMensaje
                   RsCabecera.Requery
                   RsDetalle.Requery
                   Exit Sub
                End If
                        
'                    End If
                'INSUMOCRITICO23
                If validadIngresoProducto(sCodigo) = False Then
                    Exit Sub
                End If
                'INSUMOCRITICO23
            
               sProducto = sCodigo
               ListarOperadoresConFiltro (sProducto) 'Oscar Ortega
               Dim xxx As String
               xxx = RsProducto.Filter
               RsProducto.Filter = adFilterNone
               RsProducto.MoveFirst
               RsProducto.Find "Codigo = '" & sProducto & "'"
               
               If vComanda Then
                  sTipo = "Comanda"
                  sDescrip = sComanda
                  frmNumPad.Caption = "Comanda"
                  frmNumPad.Show vbModal
                
                  sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
                  
                  If wComanda And sComanda = "" Then
                     MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
                     Exit Sub
                  End If
               
                  If Not wEnter Then
                     Exit Sub
                  End If
               End If
                  
               If lBal And RsProducto!lBalanza Then
                  Dim nResultado As Double
                  nResultado = Pesar(nBalanzaPuerto)
                  If (lActivaTresDecimales) Then
                    nResultado = Format(nResultado, "#,##0.000")
                  Else
                    nResultado = Format(nResultado, "#,##0.00")
                  End If
                  
                  If nResultado > 0 Then
                     InsertaProducto nResultado
                  End If
               Else
               nCantidad = 1
                  InsertaProducto 1
               End If
            
               If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                  lPropiedad = True
               End If
               RsProducto.Filter = IIf(xxx = "0", "", xxx)
                                 
            End If
            txtBarra.SetFocus
                        
       Case Is = 7  'Importar Pedido
            sTemp = ""
            
            If lPasswordImportarPedido Then
               If Supervisor("15") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            
            If lFiltroTipoPedido Then
                Isql = "select *, Caso = case vpedidoGrilla.tCaja when '" & sCaja & "' then 'Exportar' ELSE 'Importar' END " & _
                       "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                       "where MPEDIDO.tTipoPedido='" & sTipoPedidoPD & "' and tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                       "vpedidoGrilla.tEstadoPedido ='01' and " & _
                       "(vpedidoGrilla.tCaja <>'" & sCaja & "' or (vpedidoGrilla.tCaja='" & sCaja & "' and len(ltrim(tCajaAnterior))<>0 )) " & _
                       "order by Mesa, vPedidoGrilla.tObservacion"
            Else
                If lMCPV Then
                    Isql = "select *, 'Importar' as Caso " & _
                           "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                           "where tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                           "vpedidoGrilla.tEstadoPedido ='01' and " & _
                           "vpedidoGrilla.tUsuario <>'" & sUsuario & "' " & _
                           "order by Mesa, vPedidoGrilla.tObservacion"
                Else
                    Isql = "select *, Caso = case vpedidoGrilla.tCaja when '" & sCaja & "' then 'Exportar' ELSE 'Importar' END " & _
                           "from vPedidoGrilla INNER JOIN dbo.MPEDIDO ON dbo.vPedidoGrilla.Codigo = dbo.MPEDIDO.tCodigoPedido " & _
                           "where tCodigoPedido not in (select distinct dbo.DPEDIDO.tcodigopedido FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido where  tEstadoPedido='01' and (tFacturado = 'F' or tFacturado='P')) and " & _
                           "vpedidoGrilla.tEstadoPedido ='01' and " & _
                           "(vpedidoGrilla.tCaja <>'" & sCaja & "' or (vpedidoGrilla.tCaja='" & sCaja & "' and len(ltrim(tCajaAnterior))<>0 )) " & _
                           "order by Mesa, vPedidoGrilla.tObservacion"
                End If
            End If
            
            Call ConfGrilla(8, frmBusquedaRapida.grdGrilla, "Función", 2, "Caso", 800, 0, 0, "", _
                                                            "Caja", 2, "tCaja", 550, 2, 0, "", _
                                                            "Mesa", 2, "Mesa", 1000, 0, 0, "", _
                                                            "Observacion", 2, "tObservacion", 1800, 0, 0, "", _
                                                            "Pedido", 2, "Descripcion", 1000, 0, 0, "", _
                                                            "Monto", 2, "Suma", 900, 1, 0, "###,##0.00", _
                                                            Mesero, 2, "Mozo", 1100, 0, 0, "", _
                                                            "Usuario", 2, "tUsuario", 1100, 0, 0, "")
                                                            
            If lBuscaPedidoNumero = True Then
                frmBusquedaRapida.nPredeterm = 4
            Else
                frmBusquedaRapida.nPredeterm = 2
            End If
                
            
            frmBusquedaRapida.Show vbModal
            
            If wEnter Then
               sPedido = sCodigo
                              
               'Refrescar
               Dim RsRefresca As Recordset
               Set RsRefresca = Lib.OpenRecordset("select tTurno, tCaja, tCajaAnterior, tTurnoAnterior from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
               
               If lMCPV Then
                  Cn.Execute "update MPEDIDO set tUsuario = '" & sUsuario & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
               ElseIf RsRefresca!tTurno = "MOZO" Then 'Importar desde Mozos
                  If lInfhotel Then
                     Cn.Execute "update MPEDIDO set tPuntoVenta='" & sPuntoVentaInfhotel & "', tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  Else
                     Cn.Execute "update MPEDIDO set tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                  End If
                  Cn.Execute "update TCAJA set lRefresca=1 where tCaja='" & RsRefresca!tCaja & "'"
                  
                  'MESA247 IMPORTAR
                  If MESA247 Then
                    If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                          Cn.Execute "update MPEDIDO set lEnvioMesaEstado = 1, lEnvioMesa=0 where tCodigoPedido='" & sPedido & "'"
                    End If
                  End If
               ElseIf RsRefresca!tTurno = sTurno And Not IsNull(RsRefresca!tTurnoAnterior) Then  'Exportar Mozo
                  Cn.Execute "update MPEDIDO set tTurno = '" & RsRefresca!tTurnoAnterior & "', tCaja='" & RsRefresca!tCajaAnterior & "' where tCodigoPedido='" & sPedido & "' "
                    'MESA247 EXPORTAR
                    If MESA247 Then
                        If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                            Cn.Execute "update MPEDIDO set lEnvioMesaEstado = 0 where tCodigoPedido='" & sPedido & "'"
                        End If
                    End If
               ElseIf RsRefresca!tTurno = sTurno And IsNull(RsRefresca!tTurnoAnterior) Then 'Exportar
                  Cn.Execute "update TCAJA set lRefresca=1 where tCaja='" & RsRefresca!tCajaAnterior & "'"
                  Cn.Execute "update MPEDIDO set tTurno = 'MOZO', tCaja='" & RsRefresca!tCajaAnterior & "' where tCodigoPedido='" & sPedido & "' "
                  
               Else  'Importar

                  If Calcular("select count(ddocumento.tDocumento) as Codigo from DDOCUMENTO inner join mdocumento on ddocumento.tdocumento= mdocumento.tdocumento where tCodigoPedido='" & sPedido & "' and mdocumento.testadodocumento<>'04'", Cn) > 0 Then
                     MsgBox "Error: No se puede importar pedido con Documentos", vbExclamation, sMensaje
                     Exit Sub
                  Else
                     Cn.Execute "update MPEDIDO set tCaja = '" & sCaja & "', tTurno='" & sTurno & "' where tCodigoPedido='" & sPedido & "' "
                     
                     'MESA247 IMPORTAR
                     If MESA247 Then
                        If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                           Cn.Execute "update MPEDIDO set lEnvioMesaEstado = 1, lEnvioMesa=0 where tCodigoPedido='" & sPedido & "'"
                        End If
                    End If
                    
                  End If
               End If
                              
               RsCabecera.Requery
               RsDetalle.Requery
               RsCombo.Requery
               If RsCabecera.RecordCount <> 0 Then
                  RsCabecera.MoveFirst
               End If
               If RsRefresca!tTurno <> sTurno Then
                  RsCabecera.Find "codigo='" & sPedido & "'"
               End If
               Set RsRefresca = Nothing
               If RsCabecera.EOF Then
                  LimpiaCabecera
               End If
               
               'GCAA 05022020
                If (RsCabecera.RecordCount > 0) And Not RsCabecera.EOF Then
                    Call grdCabecera_RowColChange("", 1)
                    'AsignaPedido
               End If
               
               
            End If
            
       Case Is = 8  'Salir Propiedades
            If wAgregaCombo Then
                'Oscar Ortega---------------------------------
               If ObligaPropiedad(sCombo) = False Then
                    Exit Sub
               End If
               'Fin Oscar Ortega-----------------------------
               cmdOpcion(1).Enabled = False
               RsCombo.Requery
               If RsCombo.RecordCount > 0 Then
                  RsCombo.MoveFirst
                  RsCombo.Find "titemCombo = '" & xItem & "'"
               End If
               ActivaFrame fraProductoCombo
            Else
               'Oscar Ortega---------------------------------
               If ObligaPropiedad(sProducto) = False Then
                    Exit Sub
                Else
                    grdDetalle.Enabled = True
                    grdCabecera.Enabled = True
               End If
               'Fin Oscar Ortega-----------------------------
               RsDetalle.Requery
               RsDetalle.MoveFirst
               RsDetalle.Find "titem = '" & sItem & "'"
               ActivaOpcion True
               ActivaCabecera True
               ActivaFrame fraPlato
            End If
            
       Case Is = 9  'IR
            If Not RsCabecera.EOF Then
               wEnter = False
               Sw = False
               sDescrip = ""
               
                'busquedasocio
               'busquedaxSocio
               sTemp = ""
               lCargaDesdePedido = False
               Call ConfGrilla(6, frmIr.grdGrillaPedido, "Pedido", 2, "Pedido", 1200, 0, 0, "", _
                                        "Tipo Pedido", 2, "TipoPedido", 1300, 0, 0, "", _
                                        "Salon", 2, "Salon", 1700, 0, 0, "", _
                                        "Mesa", 2, "Mesa", 800, 0, 0, "", _
                                        "Cliente", 2, "Cliente", 3800, 0, 0, "", _
                                        "Total", 2, "Venta", 1000, 1, 0, "")
                If lBuscarPedidoFiltrarMesa = True Then
                    frmIr.nPredeterm = 3
                Else
                    frmIr.nPredeterm = 4
                End If
               
               
               
               frmIr.Show vbModal
               
               If wEnter Then
                  RsCabecera.MoveFirst
                  
                  If lCargaDesdePedido = True Then
                     RsCabecera.Find ("Codigo='" & sCodigo & "'")
                  Else
                     RsCabecera.Find ("tMesa='" & sCodigo & "'")
                  End If
                  
                  
                  If RsCabecera.EOF Then
                     RsCabecera.MoveFirst
                     MsgBox "El pedido pertenece a otra caja ", vbExclamation, sMensaje
                  End If
                  grdCabecera.Bookmark = RsCabecera.Bookmark
               End If
               
               If Sw Then
                  RsCabecera.MoveFirst
                  If Len(Trim(sDescrip)) > 0 Then
                     RsCabecera.Find ("tObservacion like '*" & Trim(sDescrip) & "*'")
                     If RsCabecera.EOF = True Then
                        RsCabecera.MoveFirst
                        MsgBox "No se encontro " & sDescrip, vbExclamation, sMensaje
                        Exit Sub
                     End If
                     grdCabecera.Bookmark = RsCabecera.Bookmark
                  End If
                  Sw = False
               End If
            End If
            
       Case Is = 10 ' Refrescar
            Me.cmdOpcion(10).Caption = "Refrescar"
            
            If lIntPatio Then
                Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=3 and testadopedido<>'03' "
                Set RsPedido = Lib.OpenRecordset(Isql, Cn)
        
                If RsPedido.RecordCount > 0 Then
                Dim delmax As Integer
                    Do
                        If delmax < 3 Then
                            Isql = "Update Mpedido Set tusuario='" & sUsuario & "', tTurno = '" + sTurno + "', TCaja = '" + sCaja + "', tTurnoAnterior = '" + sTurno + "', tCajaAnterior = '" + sCaja + "', tSalon = '01' Where tCodigoPedido = '" + RsPedido!tCodigoPedido + "'"
                            Cn.Execute Isql
                        End If
                        delmax = delmax + 1
                        RsPedido.MoveNext
                    Loop While Not RsPedido.EOF
                End If
                Me.cmdOpcion(10).backColor = &H8000000F
            End If
            
            RsCabecera.Requery
            RsDetalle.Requery
            RsProductoPropiedad.Requery
            RsCombo.Requery
            If RsCabecera.EOF Then
               LimpiaCabecera
            End If
            If lvisor Then
               Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
               'Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sPedido & "',1)"
            End If
       
       Case Is = 11 ' Combo
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
       
            If wCombo Then
               fraCombo.Visible = True
               fraProductoCombo.Visible = True
               wAgregaCombo = True
               ActivaCabecera False
               ActivaOpcion False
               If Not RsCombo.EOF Then
                  RsCombo.MoveFirst
               End If
               AsignaProductoCombo
               
               cmdOpcion(1).Visible = False
               cmdOpcion(2).Visible = False
               txtBarra.SetFocus
            End If
            
       Case Is = 12 ' Visualizar Pedido
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            If lMCPV Then
                If Not ValidaExistenciaProducto() Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
            End If
            If lPrecuenta Then
               sTipo = "ConMozo"
            Else
               sTipo = "SinMozo"
            End If
            sTipo = ""
            frmPedido.Show vbModal
            If wEnter Then
               RsDetalle.Requery
               'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
               SumPrecioVenta (sPedido)
            End If
            
       Case Is = 13 ' Cuentas internas
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            
            
            
            
            ' Miguel Mendoza - 10062019
            If Not VerificaComboMaximo() Then Exit Sub
            ' *************************
            If lMCPV Then
                If Not ValidaExistenciaProducto() Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
            End If
       
          
       
            
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
               If i > 0 Then
                  MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
                  RsDetalle.Requery
                  Exit Sub
               End If
            End If
                        
            If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0') <> '0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
               MsgBox "Imposible pasar el pedido a Cuenta Corrientes, pedidos con items Facturados", vbExclamation, sMensaje
               RsDetalle.Requery
               Exit Sub
            End If
                        
            'Chequea Descuento
            If CDbl(TxtDescuento) > 0 Then
               nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
               If nTope > 0 Then
                  If Calcular("select lTopePedido as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) Then
                     If CDbl(TxtDescuento) > nTope Then
                        If MsgBox("El Descuento a aplicar Supera El Tope Registrado por Pedido" & Chr(13) & "¿Desea aplicar el Tope de " & sMonN & " " & nTope & "?", vbQuestion + vbYesNo) = vbYes Then
                           CalculaAplicaTope (nTope)
                        Else
                           RsDetalle.Requery
                           Exit Sub
                        End If
                     End If
                  Else
                     nTotalMes = Calcular("select sum(DPEDIDO.nDescuento*nCantidad) as Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                                          "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'", Cn)
                     
                     If CDbl(TxtDescuento) + nTotalMes > nTope Then
                        If CDbl(TxtDescuento) < nTope Then
                           If MsgBox("El Descuento a aplicar Supera El Tope Registrado dentro de un mes" & Chr(13) & "¿Desea aplicar el saldo " & sMonN & " " & nTope - nTotalMes & "?", vbQuestion + vbYesNo) = vbYes Then
                              CalculaAplicaTope (nTope - nTotalMes)
                           Else
                              RsDetalle.Requery
                              Exit Sub
                           End If
                        Else
                            MsgBox "El Descuento a aplicar Supera El Tope Registrado dentro de un mes", vbExclamation
                            RsDetalle.Requery
                            Exit Sub
                        End If
                     End If
                  End If
               End If
            End If
                        
            sUsuarioAutoriza = sUsuario
            If Supervisor("09") = False Then
               MsgBox "Clave no permitida", vbExclamation, sMensaje
               RsDetalle.Requery
               Exit Sub
            End If
            sUsuarioAutoriza = sVar1
                                                      
            'Chequea si existe platos a facturar
            sTD = "N"
            RsDetalle.MoveFirst
            Do While Not RsDetalle.EOF
               If (Len(Trim(RsDetalle!tFacturado)) = 0 Or IsNull(RsDetalle!tFacturado)) Then
                  sTD = "S"
                  Exit Do
               End If
               RsDetalle.MoveNext
            Loop
    
            If sTD <> "S" Then
               MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
               Exit Sub
            End If
                                                                        
            sTemp = Calcular("select tApellido as Codigo from TDELIVERY where tCodigoDelivery='" & sCliente & "'", Cn)
            
            If sTemp = "0" Then
                sTemp = ""
            End If
            'LGCENTRALPEDIDO
            Isql = "select * from vCompania where lActivo = 1 order by Descripcion"
            Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Id.", 2, "tTelefono", 1100, 0, 0, "", _
                                                            "Cliente", 2, "Descripcion", 3650, 0, 0, "", _
                                                            "Identidad", 2, "tIdentidad", 1150, 0, 0, "", _
                                                            "Linea", 2, "nLinea", 1150, 1, 0, "##,##0.00", _
                                                            "Consumo", 2, "nConsumo", 1150, 1, 0, "##,##0.00")
                                                                                                                       
            frmBusquedaRapida.nPredeterm = 1
            frmBusquedaRapida.Show vbModal
                                                                         
            If Not wEnter Or sCodigo = "" Then
               RsDetalle.Requery
               Exit Sub
            End If

            sCliente = sCodigo
            xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and isnull(tFacturado,'') = '' and tCodigoPedido='" & sPedido & "'", Cn)
            
            Dim vley1 As Double
            vley1 = CDbl(sValorLey1) * (Calcular(" select isnull(sum(dpedido.ncantidad),0) as codigo from dpedido inner join vproducto on dpedido.tcodigoproducto=vproducto.codigo where vproducto.lley1=1 and dpedido.tcodigopedido='" & sPedido & "' ", Cn))
            xSuma = xSuma + vley1
                                                                                                                                   

            'Validacion de escoger segun estadoFrecuente
            Dim lValidaEstado As Boolean
            lValidaEstado = False
            lValidaEstado = Calcular("select ISNULL(tb.nValor,0) as codigo from TDELIVERY t INNER JOIN TTABLA tb on t.tEstadoFrecuente = tb.TCODIGO where  t.tCodigoDelivery='" & sCliente & "' and tb.TTABLA='ESTADOFRECUENTE'", Cn)
            If lValidaEstado Then
                MsgBox "No es posible seleccionar al cliente, estado no permitido", vbCritical, sMensaje
                RsDetalle.Requery
                Exit Sub
            End If
            
            
            'centralizada
'            If lCentral = False Then
               xLinea = Calcular("select nLinea as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", Cn)
               xConsumo = Calcular("select nConsumo as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", Cn)
'            Else
'                Dim conServidor As ADODB.Connection
'                Set conServidor = devuelveConexionCentral(sServidorCentral, bdInforestCentral)
'                If conServidor.State Then
'                   xLinea = Calcular("select isnull(nLinea,0) as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", conServidor)
'                   xConsumo = Calcular("select isnull(nConsumo,0) as Codigo FROM TDELIVERY where TCODIGODELIVERY = '" & sCliente & "'", conServidor)
'                 Else
'                        MsgBox "No es posible conectar con el servidor central" & vbCrLf & "No se puede trabajar con la cuenta corriente", vbCritical, sMensaje
'                        Exit Sub
'                End If
'
'
'            End If
            
            If xSuma > xLinea - xConsumo Then
               MsgBox "El Cliente no tiene linea suficiente " & Chr(13) & _
                      "Linea : " & Format(xLinea, "###,##0.00") & "  Consumo : " & Format(xConsumo, "###,##0.00") & Chr(13) & _
                      "Saldo : " & Format(xLinea - xConsumo, "###,##0.00"), vbCritical, sMensaje
               Exit Sub
            End If

            If MsgBox("Esta seguro de Enviar el Pedido Nro: " & sPedido & _
               Chr(13) & "a Cuentas Corrientes al Cliente " & sDescrip & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
               Exit Sub
            End If
            
            If lInfhotel Then
               xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn)
               If Len(Trim(RsCabecera!tComanda)) = 0 Or IsNull(RsCabecera!tComanda) Then
                  sComandaInfhotel = IIf(IsNull(rsPuntoVenta!nUltimoComanda), "00000000", rsPuntoVenta!nUltimoComanda)
                  sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
                  CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & sPuntoVenta & "'"
                  sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!descripcion, 1, 3))
                  rsPuntoVenta.Requery
                  rsPuntoVenta.MoveFirst
                  rsPuntoVenta.Find "Codigo='" & sPuntoVenta & "'"
                                              
                  'Genero las comandas en Infhotel
                  'Cabecera
                  Isql = "Insert into MCOMANDA " & _
                         "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                         "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                         "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE) " & _
                         "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & xSuma & ", '01', " & _
                         "1, '01', '', '', '', " & _
                         "'" & sCajaInfhotel & "', '', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '01', getdate(), getdate(), '" & xUsuario & "')"
                  CnInfhotel.Execute Isql
               Else
                  sComandaInfhotel = RsCabecera!tComanda
                  CnInfhotel.Execute "update MCOMANDA set TASIGNACION='01', TCODIGORESERVA='', TNUMEROHABITACION='', TCLIENTE='', nTotal= " & xSuma & ", tEstado='01' " & _
                                     "where tComanda ='" & RsCabecera!tComanda & "'"
               End If
               
               'Detalle
               Dim xMovimiento As Integer
               CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & RsCabecera!tComanda & "'"
               xMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
               
               Isql = "Insert into DCOMANDA " & _
                      "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                      "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                      "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                      "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & xMovimiento & ", tComanda as tNotaPedido, tInfhotel as tCodigoItem, " & _
                      "T1.nPrecioVenta as nPrecioUnitario, nCantidad, nVenta as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                      "'' as tCuenta, '" & sCajaInfhotel & "' as tCaja, '' as tDocumento, '" & IIf(sDescrip = "Reserva", "03", "02") & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                      "where tCodigoPedido ='" & sPedido & "'"
               CnInfhotel.Execute Isql
               
               Cn.Execute "update MPEDIDO set tComanda = '" & sComandaInfhotel & "' where tCodigoPedido='" & sPedido & "'"
            End If
                        
            Cn.Execute "Update MPEDIDO set tClienteCtaCte ='" & sCliente & "', tEstadoPedido = '04' where tCodigoPedido='" & sPedido & "'"
            
            'LGCENTRALPEDIDO
'            If lCentral = False Then
             Cn.Execute "Update TDELIVERY set nConsumo = " & xConsumo + xSuma & " where TCODIGODELIVERY ='" & sCliente & "'"
'            Else
'                conServidor.Execute "Update TDELIVERY set nConsumo = " & xConsumo + xSuma & " where TCODIGODELIVERY ='" & sCliente & "'"
'            End If
            'LGCENTRALPEDIDO
            Isql = "select * from vCtaCte " & _
                   "WHERE Codigo='" & sPedido & "'"
            Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
            
            If RsImpresion.RecordCount = 0 Then
               LimpiaRs
               MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
            Else
                'GCAA REQ REST-C009-2021 - PASTELERIA TENTACIONES
                If (lActivarMensajeCtsCorrientes) Then
                     
                     If MsgBox("Desea Imprimir? ", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                       
                     Else
                        ImprimeCtaCte RsImpresion
                     End If
                Else
                    ImprimeCtaCte RsImpresion
                End If
            End If
            
            If lImpComandaBarra And Not lPrinter Then
                cmdOpcion_Click (3)
            End If
            
            LimpiaRs
            
            'Libera la Mesa
            Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
            RsCabecera.Requery
            RsDetalle.Requery
            RsCombo.Requery
            
            If RsCabecera.EOF Then
               LimpiaCabecera
            End If

                        
      Case Is = 14  'Cargos
            If RsDetalle.RecordCount = 0 Then
               Exit Sub
            End If
            ' Miguel Mendoza - 10062019
            If Not VerificaComboMaximo() Then Exit Sub
            ' *************************
            If lInfhotel Then
                If rsPuntoVenta.BOF Or rsPuntoVenta.EOF Then
                    MsgBox "No se ha Asignado un Punto de Venta de Infhotel, Verificar!!!"
                    Exit Sub
                End If
            End If
            If lMCPV Then
                If Not ValidaExistenciaProducto() Then
                    MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                    RsCabecera.Requery
                    RsDetalle.Requery
                    Exit Sub
                End If
            End If
            If RsCabecera.RecordCount = 0 Then
               Exit Sub
            End If
            If lPrinter And lObligaPrinter Then
               i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
               If i > 0 Then
                  MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            If sPuntoVenta = "" Then
               MsgBox "Falta ingresar el punto de venta", vbExclamation, sMensaje
               Exit Sub
            End If
            
            If Calcular("select count(tFacturado) as Codigo from DPEDIDO where tCodigoPedido ='" & sPedido & "' and isnull(tFacturado,'0') <> '0' and len(ltrim(tFacturado)) <> 0", Cn) > 0 Then
               MsgBox "Imposible pasar el pedido a Cargos, pedidos con items Facturados", vbExclamation, sMensaje
               Exit Sub
            End If
                               
            'Chequea si existe platos a facturar
            sTD = "N"
            RsDetalle.MoveFirst
            Do While Not RsDetalle.EOF
               If (Len(Trim(RsDetalle!tFacturado)) = 0 Or IsNull(RsDetalle!tFacturado)) Then
                  sTD = "S"
                  Exit Do
               End If
               RsDetalle.MoveNext
            Loop
    
            If sTD <> "S" Then
               MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
               Exit Sub
            End If

            frmCargo.Show vbModal

            If Not wEnter Then
               Exit Sub
            End If
           
            Dim tItem As Integer
            Dim Correlativo As Integer
            Dim CorrelaProp As Integer
            Dim nMovimiento As Integer
            Dim sAsignado   As String
            Dim MonPuntoventa As String
            
            sCliente = sCodigo
            xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido='" & sPedido & "'", Cn)
            
            If sDescrip = "Infhotel" Then
               If MsgBox("Esta seguro de Generar el Pedido Nro: " & sPedido & " en Infhotel?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               sHabitacion = ""
               sReserva = ""
               sPasajero = ""
            Else
               If MsgBox("Esta seguro de Enviar el Pedido Nro: " & sPedido & _
                  Chr(13) & "a la " & Trim(sDescrip) & " " & IIf(sDescrip = "Reserva", sReserva, sHabitacion) & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
            End If
            
            MonPuntoventa = Calcular("select max(tmoneda) as codigo from vpuntoventa where tpuntoventa='" & sPuntoVenta & "'", CnInfhotel)
            If Len(Trim(RsCabecera!tComanda)) = 0 Or IsNull(RsCabecera!tComanda) Then
               Call PUltimaComanda
               sComandaInfhotel = Calcular("select left(MAX(tComanda),8) as Codigo from MCOMANDA where tPuntoVenta='" & sPuntoVenta & "'", CnInfhotel)
               sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
               CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & sPuntoVenta & "'"
               sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!descripcion, 1, 3))
               rsPuntoVenta.Requery
                              
               'Genero las comandas en Infhotel
               'Cabecera
               If sDescrip = "Infhotel" Then sAsignado = "01"
               If sDescrip = "Reserva" Then sAsignado = "03"
               If sDescrip = "Habitacion" Then sAsignado = "02"
               If sTipoComanda = "01" Then
                    Isql = "Insert into MCOMANDA " & _
                           "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                           "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                           "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE, tIncluido, nRoomSer, nDescuento, tNotaPedido, tCompania, tContacto, lInforest ) " & _
                           "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & IIf(MonPuntoventa = "01", xSuma, xSuma / nTC) & ", '01', " & _
                           "1,'" & sAsignado & "', '" & sReserva & "', '" & sHabitacion & "', '', " & _
                           "'', '', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '" & MonPuntoventa & "', getdate(), getdate(), '" & xUsuario & "','" & sTipoComanda & "', " & IIf(sTipoPedido = "03", nLlevar, 0) & ", " & xDescuento & ", '" & sPedido & "', '" & sCompania & "', '" & sContacto & "', 1)"
                    CnInfhotel.Execute Isql
               Else
                    Isql = "Insert into MCOMANDA " & _
                          "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                          "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                          "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE, tIncluido, nRoomSer, nDescuento, tNotaPedido, tCompania, tContacto, lInforest) " & _
                          "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), 0, '01', " & _
                          "1,'" & sAsignado & "', '" & sReserva & "', '" & sHabitacion & "', '', " & _
                          "'', '', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '" & MonPuntoventa & "', getdate(), getdate(), '" & xUsuario & "','" & sTipoComanda & "', " & IIf(sTipoPedido = "03", nLlevar, 0) & ", " & xDescuento & ", '" & sPedido & "', '" & sCompania & "', '" & sContacto & "', 1)"
                    CnInfhotel.Execute Isql
               End If
            Else
               sComandaInfhotel = RsCabecera!tComanda
               CnInfhotel.Execute "update MCOMANDA set TESTADO='01', TASIGNACION='" & sAsignado & "', TCODIGORESERVA='" & sReserva & "', TNUMEROHABITACION='" & sHabitacion & "', TCLIENTE='" & sPasajero & "', " & _
                                  "tpuntoventa='" & sPuntoVenta & "',tmoneda='" & MonPuntoventa & "',tincluido='" & sTipoComanda & "',ntotal=" & IIf(MonPuntoventa = "01", xSuma, xSuma / nTC) & ",ndescuento=" & xDescuento & " where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & sPuntoVenta & "'"
            End If
            If sDescrip = "Habitacion" Then
               txtObservacion.Caption = "Hab: " & sHabitacion
            ElseIf sDescrip = "Reserva" Then
               txtObservacion.Caption = "Res: " & sReserva
               sHabitacion = ""
            End If
            Isql = "Update MPEDIDO Set " & _
                    "tEstadoPedido ='05', " & _
                    "tComanda ='" & sComandaInfhotel & "', " & _
                    "tPuntoVenta ='" & sPuntoVenta & "', " & _
                    "tReserva ='" & sReserva & "', " & _
                    "tHabitacion ='" & sHabitacion & "', " & _
                    "tObservacion='" & txtObservacion.Caption & "', " & _
                    "tFichaPasajero ='" & sFichaPasajero & "', " & _
                    "tTipoComanda ='" & sTipoComanda & "', " & _
                    "tPasajero ='" & sPasajero & "' " & _
                    "  where tCodigoPedido = '" & sPedido & "'"
            Cn.Execute Isql
            
            'Detalle
            CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & RsCabecera!tComanda & "'"
            nMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
            If sTipoComanda = "01" Then
                Isql = "Insert into DCOMANDA " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & " as nMovimiento, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                       IIf(MonPuntoventa = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad, " & IIf(MonPuntoventa = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                       "'' as tCuenta, '' as tCaja, '' as tDocuemento, '" & sAsignado & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                       "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                       "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
                CnInfhotel.Execute Isql
            Else
                Isql = "Insert into DCOMANDA " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem ," & nMovimiento & " as nMovimiento, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                       " 0 as nPrecioUnitario, nCantidad, 0 as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                       "'' as tCuenta, '' as tCaja, '' as tDocuemento, '" & sAsignado & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                       "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                       "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
                CnInfhotel.Execute Isql
            End If
            
            'Propinas
            If val(sPropina) > 0 Then
                'Grabo la propina en el Mpropina del Inforest
                Isql = "insert into MPROPINA " & _
                      "(tcodigopedido,fregistro,tmoneda,nmonto,tusuario, tComanda) " & _
                      "values('" & sPedido & "',getdate(),'" & sMonPropina & "'," & sPropina & ",'" & xUsuario & "', '" & sComandaInfhotel & "')"
                Cn.Execute Isql
                
                'Grabo la propina en el Detalle de la comanda
                tItem = Calcular("select max(titem) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                Isql = "Insert Into dcomanda " & _
                       "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                       "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                       "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                       "values('" & sComandaInfhotel & "','" & sPuntoVenta & "','00','" & tItem & "','1','" & sPedido & "', " & _
                       "'100000'," & sPropina & ",'1'," & sPropina & ",'','" & sReserva & "','" & sHabitacion & "','', " & _
                       " '','','" & sAsignado & "','" & xUsuario & "',getdate(),getdate())"
                CnInfhotel.Execute Isql
               
                'Grabo la propina en el Mpropina del Infhotel
                CorrelaProp = Calcular("select max(ncorrela) as codigo from mpropina", CnInfhotel) + 1
            
                Isql = "Insert Into MPROPINA " & _
                       "(ncorrela,tcodigoreserva,tnumerohabitacion,tcomanda,tcodigoitem,tmoneda,nmonto,tdocumento,tresponsable,testado,ffecha,tusuario, tPuntoVenta) " & _
                       "values(" & CorrelaProp & ",'" & sReserva & "','" & sHabitacion & "','" & sComandaInfhotel & "','100000','" & sMonPropina & "'," & sPropina & ", " & _
                       "'','" & Mid(sMozo, 2, 3) & "','01',getdate(),'" & xUsuario & "','" & sPuntoVenta & "')"
                CnInfhotel.Execute Isql
                
                Isql = "Update MCOMANDA set ncorrelaprop=" & CorrelaProp & " where tcomanda='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
                CnInfhotel.Execute Isql
            End If
                        
            Dim tcuenta As String
            tcuenta = Calcular("select top 1 tcuenta as codigo from tcuentaasignadareserva where tcodigopasajero='" & sFichaPasajero & "' and tCodigoReserva='" & sReserva & "'", CnInfhotel)
            'tcuenta = Calcular("select top 1 tcuenta as codigo from tcuentaasignadareserva where tcodigopasajero='" & sFichaPasajero & "'", CnInfhotel)
                        
            'Actualiza las Cuentas Corrientes Infhotel
            If sDescrip = "Reserva" And sTipoComanda = "01" Then
               i = Calcular("select max(tNumeroCorrelativo) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'", CnInfhotel)
               Isql = "Insert into TCUENTARESERVA " & _
                      "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad,testado, ttipo,tHotel, " & _
                      " nTotal, tNumeroCorrelativo,tpuntoventa, tItem, tUsuario, tcuenta) " & _
                      "select '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion , getdate() as fFecha, getdate() as hHoraMovimiento, '" & sComandaInfhotel & "' as tComanda, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                      IIf(sMonedaBase = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad,'' as testado,'' as ttipo,'" & sHotel & "' as tHotel," & IIf(sMonedaBase = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, tItem + " & i & ",'" & sPuntoVenta & "' as tpuntoventa ,tItem, '" & xUsuario & "' as tUsuario, '" & tcuenta & "'" & _
                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                      "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
               CnInfhotel.Execute Isql
                             
               If val(sPropina) > 0 Then
                    'Graba la propina en la Cta Cte de la Reserva
                    i = Calcular("select max(titem) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'and tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                    Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTARESERVA where tCodigoReserva='" & sReserva & "'", CnInfhotel) + 1
                    
                    If sMonedaBase = sMonPropina Then
                         nPropinaCalc = sPropina
                    Else
                     If (sMonedaBase = "02" And sMonPropina = "01") Then
                            nPropinaCalc = CDbl(sPropina) / nTC
                     Else
                            nPropinaCalc = CDbl(sPropina) * nTC
                     End If
                    End If
                    
                    
                    Isql = "insert into tcuentareserva " & _
                          "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad,testado,ttipo,tHotel, " & _
                          " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                          " values('" & sReserva & "','" & sHabitacion & "',getdate(),getdate(),'" & sComandaInfhotel & "','" & sPedido & "','100000'," & nPropinaCalc & ",'1','','','00', " & _
                          " " & nPropinaCalc & "," & Correlativo & ",'" & sPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                    CnInfhotel.Execute Isql
                End If

            ElseIf sDescrip = "Habitacion" And sTipoComanda = "01" Then
               i = Calcular("select max(tNumeroCorrelativo) as Codigo from TCUENTAHABITACION where tNumeroHabitacion='" & sHabitacion & "' and tCodigoReserva='" & sReserva & "'", CnInfhotel)
               Isql = "Insert into TCUENTAHABITACION " & _
                      "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento,testado,ttipo, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad, tHotel, " & _
                      " nTotal, tNumeroCorrelativo,tpuntoventa, tItem, tUsuario, tcuenta) " & _
                      "select '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion , getdate() as fFecha, getdate() as hHoraMovimiento,'' as testado,'' as ttipo ,'" & sComandaInfhotel & "' as tComanda, '" & sPedido & "' as tNotaPedido, tInfhotel as tCodigoItem, " & _
                      IIf(sMonedaBase = "01", "T1.nPrecioVenta", "T1.nPrecioVenta / " & nTC) & " as nPrecioUnitario, nCantidad, '" & sHotel & "' as tHotel," & IIf(sMonedaBase = "01", "nVenta", "nVenta / " & nTC) & " as nTotal, tItem + " & i & ",'" & sPuntoVenta & "' as tpuntoventa ,tItem, '" & xUsuario & "' as tUsuario, '" & tcuenta & "'" & _
                      "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                      "where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'"
               CnInfhotel.Execute Isql
               
               If val(sPropina) > 0 Then
                  'Graba la propina en la Cta Cte de la Habitacion
                  i = Calcular("select max(titem) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & sReserva & "'and tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
                  Correlativo = Calcular("select max(tnumerocorrelativo) as Codigo from TCUENTAHABITACION where tCodigoReserva='" & sReserva & "'", CnInfhotel) + 1
                  
                    If sMonedaBase = sMonPropina Then
                         nPropinaCalc = sPropina
                    Else
                     If (sMonedaBase = "02" And sMonPropina = "01") Then
                            nPropinaCalc = CDbl(sPropina) / nTC
                     Else
                            nPropinaCalc = CDbl(sPropina) * nTC
                     End If
                    End If
                  
                  Isql = "Insert into TCUENTAHABITACION " & _
                         "(tCodigoReserva, tNumeroHabitacion, fFecha, hHoraMovimiento,testado,ttipo, tComanda, tNotaPedido, tCodigoItem, nPrecioUnitario, nCantidad, tHotel, " & _
                         " nTotal, tNumeroCorrelativo,tpuntoventa ,tItem, tUsuario,ncorrelaprop) " & _
                         " Values('" & sReserva & "','" & sHabitacion & "',getdate(),getdate(),'','','" & sComandaInfhotel & "','" & sPedido & "','100000'," & nPropinaCalc & ",'1','00', " & _
                         " " & nPropinaCalc & "," & Correlativo & ",'" & sPuntoVenta & "'," & i & ",'" & xUsuario & "'," & CorrelaProp & ")"
                  CnInfhotel.Execute Isql
               End If
            End If

           'CarlosD 13/11/2006
            CorrelativoC = 0
            CorrelativoC = val(Calcular("select max(nmovimiento) as codigo from wmcomanda where tcomanda='" & sComandaInfhotel & "' and tpuntoventa='" & sPuntoVenta & "' and thotel='" & sHotel & "'", CnInfhotel)) + 1
            
            Isql = "INSERT INTO WMCOMANDA([tComanda],[tPuntoVenta],[fFecha],[hHora],[tMoneda],[nTotal],[tEstado],[tEmitido], [tAsignacion],[tCodigoReserva],[tNumeroHabitacion],[tUsuario],[tCodigoCompania],[tCliente],[tCodigoFuncionario],[tHotel],[NTCAMBIO],[NMOVIMIENTO],[NDESCUENTO],[NROOMSER],[TMOZO],[tIncluido],[tMesa], tNotaPedido, tCompania, tContacto ) " & _
                   "Values (" & _
                    "'" & sComandaInfhotel & "', " & _
                    "'" & sPuntoVenta & "', " & _
                    "Getdate(), " & _
                    "GetDate(), " & _
                    "'" & MonPuntoventa & "'," & _
                    "" & IIf(MonPuntoventa = "01", xSuma, xSuma / nTC) & ", " & _
                    "'01',0, " & _
                    "'" & sAsignado & "', " & _
                    "'" & sReserva & "', " & _
                    "'" & sHabitacion & "', " & _
                    "'" & xUsuario & "', " & _
                    "'', " & _
                    "'" & sCliente & "', " & _
                    "'', " & _
                    "'" & sHotel & "', " & _
                    "'" & nTC & "', '" & CorrelativoC & "', " & nDescuento & ",'','','" & sTipoComanda & "', '" & sMesa & "', '" & sPedido & "', '" & sCompania & "', '" & sContacto & "')"
                CnInfhotel.Execute Isql

            Isql = "INSERT INTO WDCOMANDA SELECT * From dbo.DComanda where tComanda='" & sComandaInfhotel & "' and tPuntoVenta='" & sPuntoVenta & "'"
            CnInfhotel.Execute Isql
            
            'Libera la Mesa
            Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
            'Juntar Mesa
            Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
            
            wEnter = False
            
            Dim xPrecuenta As Boolean
            xPrecuenta = False
            If lObligaPrecuenta Then
               xPrecuenta = True
            Else
               If MsgBox("Deseas imprimir la Precuenta?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                  xPrecuenta = True
               End If
            End If
            
            If xPrecuenta Then
                Screen.MousePointer = vbHourglass
                'Imprime Precuenta
                If lPrecuentaImpresora Then
                   frmPrecuentaImpresora.Show vbModal
                   If Not wEnter Then
                      sPropina = ""
                      sTipoComanda = ""
                      Exit Sub
                   End If
                Else
                   sCodigo = sPreCuenta
                End If
                                        
                If lPrecuentaAgrupada Then
                   Isql = "select * from vPrecuentaAgrupada WHERE Codigo='" & sPedido & "' order by tItem"
                Else
                   Isql = "select * from vPrecuenta WHERE Codigo='" & sPedido & "' order by tItem"
                End If
    
                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                ImprimeInfhotel RsImpresion, sCodigo
                Cn.Execute "update MPEDIDO set fRegCuenta = getDate() where tCodigoPedido='" & sPedido & "'"
                LimpiaRs
            End If
            
            sPropina = ""
            sTipoComanda = ""
            
            RsCabecera.Requery
            RsDetalle.Requery
            RsCombo.Requery
            
            If RsCabecera.RecordCount = 0 Then
               sPedido = ""
               sHabitacion = ""
               sReserva = ""
               sPasajero = ""
               ActivaCabecera False
               cmdCabecera(0).Visible = True
               fraPlato.Visible = False
            End If
                                                
       Case Is = 15 ' Caja Rapida
            Screen.MousePointer = vbHourglass
            sPedido = ""
            frmCajaRapida.Show vbModal
            If wEnter Then
               Unload Me
            Else
               RsCabecera.Requery
               RsDetalle.Requery
               RsCombo.Requery
            End If
                                        
       Case Is = 16 ' Cancelar del Motivo de Eliminacion
            ActivaFrame fraPlato
            ActivaCabecera True
            ActivaOpcion True
            'oo2
            ActivarDesactivarCD (True)
            
       Case Is = 17 ' Ofertas
                If Supervisor("10") = False Then
                   MsgBox "Clave no permitida", vbExclamation, sMensaje
                   Exit Sub
                End If
                                
                Dim sCriterio As String
                Dim nOferta As Double
                
                If sTipoPedido = "01" Then
                   sCriterio = " and lLocal=1"
                ElseIf sTipoPedido = "02" Then
                   sCriterio = " and lDelivery=1"
                ElseIf sTipoPedido = "03" Then
                   sCriterio = " and lLlevar=1"
                ElseIf sTipoPedido = "04" Then
                   sCriterio = " and lCanal4=1"
                Else
                   sCriterio = " and lCanal5=1"
                End If
                
                Isql = "SELECT tOferta as Codigo, tNombre as Descripcion, " & _
                       "case when nRatio>0 then 'Descuento del ' + str(nRatio,2) + '%' else " & _
                       "case when nMonto>0 then 'Descuento de ' + str(nMonto,2) + ' " & sMonedaN & "' else 'Producto al Precio de ' + ' " & sMonN & " ' + str(nPrecio,2) end end as Oferta " & _
                       "From dbo.TOFERTA WHERE lAutomatica=0 and tCodigoProducto = '" & sProducto & "' and lActivo=1 " & _
                       " and (substring(tFrecuencia," & Weekday(FechaServidor(), vbMonday) & "+1,1) = '1' or (substring(tFrecuencia,1,1)='1') and MONTH(fFecha) = " & Month(FechaServidor()) & " AND DAY(fFecha)= " & Day(FechaServidor()) & ") and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "'" & _
                       " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "')) " & sCriterio
                
                Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                "Oferta", 2, "Descripcion", 3000, 0, 0, "", _
                                                                "Descripcion de la Oferta", 2, "Oferta", 4200, 0, 0, "")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                wCalcula = True
                                
                                
                If wEnter Then
                    Dim stockOferta, stockFinal As Integer
                                        
                   'Valida stock de Oferta
                    Isql = "exec CalcularStockOfertas '" & sProducto & "', '" & sCodigo & "', " & nCantidad & ", '" & sTipoPedido & "'"
                    
                    Set ValorOferta = Lib.OpenRecordset(Isql, Cn)
                    respuestaOferta = ValorOferta!respuesta
                    
                    If UCase(respuestaOferta) = "EXITO" Then
                    
                        stockOferta = ValorOferta!stockOferta
                        
                        If stockOferta = 0 Then
                            'MsgBox "Actualmente, dispones de cantidad ilimitada para este producto", vbExclamation, sMensaje
                        Else
                            stockFinal = stockOferta - nCantidad
                        
                            MsgBox "Actualmente, dispone de una cantidad de " + CStr(stockFinal) + " unidades disponibles en oferta.", vbExclamation, sMensaje
                        End If
                        
                    Else
                        MsgBox respuestaOferta, vbExclamation, sMensaje
                        Exit Sub
                    End If
                    
                    'Fin Valida stock de Oferta
                                    
                   Isql = "select * from TOFERTA where tOferta='" & sCodigo & "' and tCodigoProducto='" & sProducto & "'"
                   Set RsOferta = Lib.OpenRecordset(Isql, Cn)
                   If RsOferta.RecordCount > 0 Then
                      RsOferta.MoveFirst
                      
                      If RsOferta!nPrecio > 0 Then
                         If RsOferta!nPrecio > nOficial Then
                             MsgBox "El monto de la oferta supera el precio del producto", vbExclamation, sMensaje
                             Exit Sub
                         Else
                             nOferta = nOficial - IIf(IsNull(RsOferta!nPrecio), 0, RsOferta!nPrecio)
                         End If
                         
                      ElseIf RsOferta!nMonto > 0 Then
                         If RsOferta!nMonto > nOficial Then
                             MsgBox "El monto de la oferta supera el precio del producto", vbExclamation, sMensaje
                             Exit Sub
                         Else
                             nOferta = RsOferta!nMonto
                         End If
                         
                      Else
                         nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
                      End If
                      
                   End If
                   
                   nPVenta = nOficial - nOferta
                   nDescuento = nOficial - nPVenta

                   Dim Acumulado As Double
                    'extranjero
                    Select Case pais 'ok
                        Case "001" 'Bolivia
                                    Acumulado = 0
                                    Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                                    Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                                    Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                                    Acumulado = (Acumulado / 100)
                                    nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta * nPorcentaje1 / 100, 0)
                                    nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta * nPorcentaje2 / 100, 0)
                                    nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta * nPorcentaje3 / 100, 0)
                                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                                    
                        Case Else 'Peru, Ecuador
                                    Acumulado = 0
                                    Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                                    Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                                    Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                                    Acumulado = 1 + (Acumulado / 100)
                                    nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                                    nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                                    nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                                    
                    End Select
                   Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                           "nDescuento = " & nDescuento & ", " & _
                           "nRecargo = " & nRecargo & ", " & _
                           "nPrecioOficial = " & nOficial & ", " & _
                           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                           "nPrecioVenta = " & nPVenta & ", " & _
                           "nventa = " & nPVenta * nCantidad & ", " & _
                           "nCantidad = " & nCantidad & ", " & _
                           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
                           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
                           "nImpuesto3 = " & nImpuesto3 * nCantidad & ", tOferta='" & sCodigo & "', tAutorizaOferta='" & sVar1 & "' " & _
                           "where tItem = '" & sItem & "' and tCodigoPedido ='" & sPedido & "'"
                   Cn.Execute Isql
                   
                   nPos = RsDetalle.AbsolutePosition
                   RsDetalle.Requery
                   RsDetalle.AbsolutePosition = nPos
                   'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
                    SumPrecioVenta (sPedido)
                   'MESA247
                    If MESA247 Then
                         If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                              Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                         End If
                    End If
                
                End If
       
   End Select
   
   HabilitaTimerColor (True)
   Exit Sub
fin:
    MsgBox "No se Completo el Proceso: " + Error, vbInformation, sMensaje
    Call Log_Inforest("PUNTO VENTA", "cmdOpcion_Click", CStr(Index), "", sPedido, Error, "", "FALLA PROCESO", sUsuario)
End Sub

Public Sub ValidaModifica()
 On Error GoTo fin
    If sPedido <> "" Then

        Set RsPedido = Lib.OpenRecordset("SELECT * FROM MPEDIDO WHERE  tCodigoPedido= '" & sPedido & "'", Cn)
        Dim valorActual As String
        'Modifica Cliente Delivery
        valorActual = IIf(IsNull(RsPedido!tClienteDelivery), "", RsPedido!tClienteDelivery)
        If valorActual <> sCliente Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA CLIENTE DELIVERY", sPedido, Error, "", "", "Modificación del Cliente Delivery. Valor modificado:" & valorActual & " - Valor Actual: " & sCliente, "Modificación del cliente delivery del pedido " & sPedido, sUsuario)
        End If
    
        'Modifica Tipo Pedido
        valorActual = IIf(IsNull(RsPedido!tTipoPedido), "", RsPedido!tTipoPedido)
        If valorActual <> sTipoPedido Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA TIPO PEDIDO", sPedido, Error, "", "", "Modificación del tipo de pedido. Valor modificado:" & valorActual & " - Valor actual: " & sTipoPedido, "Modificación del tipo pedido (" & sPedido & ") ", sUsuario)
        End If
    
        'Modifica lPrioridad
        valorActual = IIf(IsNull(RsPedido!lPrioridad), "", RsPedido!lPrioridad)
        If valorActual <> IIf(txtTipoPedido.ForeColor = &H800000, False, True) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PRIORIDAD", sPedido, Error, "", "", "Modificación de la perioridad del pedido. Valor modificado:" & valorActual & " - Valor actual: " & IIf(txtTipoPedido.ForeColor = &H800000, 0, 1), "Modificación de la prioridad del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        'Modifica Tipo Atencion
        valorActual = IIf(IsNull(RsPedido!tTipoAtencion), "", RsPedido!tTipoAtencion)
        If valorActual <> sTipoAtencion Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA TIPO ATENCIÓN", sPedido, Error, "", "", "Modificación del tipo de atención. Valor modificado:" & valorActual & " - Valor actual: " & sTipoAtencion, "Modificación del tipo atención del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        'Modifica Mozo
        valorActual = IIf(IsNull(RsPedido!tMozo), "", RsPedido!tMozo)
        If valorActual <> sMozo Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MOZO ASIGNADO", sPedido, Error, "", "", "Modificación del Mozo. Valor modificado:" & valorActual & " - Valor actual: " & sMozo, "Modificación del mozo del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tMotorizado), "", RsPedido!tMotorizado)
        If valorActual <> sMotorizado Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MOTORIZADO", sPedido, Error, "", "", "Modificación del Motorizado. Valor modificado:" & valorActual & " - Valor actual: " & sMotorizado, "Modificación del motorizado del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tObservacion), "", RsPedido!tObservacion)
        If valorActual <> txtObservacion.Caption Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA OBSERVACIÓN", sPedido, Error, "", "", "Modificación de la observación. Valor modificado:" & valorActual & " - Valor actual: " & txtObservacion.Caption, "Modificación de la observación del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tcodigoinvitado), "", RsPedido!tcodigoinvitado)
        If valorActual <> sCodigoInvitado Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA INVITADO", sPedido, Error, "", "", "Modifica invitado. Valor modificado:" & valorActual & " - Valor actual: " & sCodigoInvitado, "Modificación invitado del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tPuntoVenta), "", RsPedido!tPuntoVenta)
        If valorActual <> sPuntoVenta Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PUNTO VENTA", sPedido, Error, "", "", "Modificación del punto de venta. Valor modificado:" & valorActual & " - Valor actual: " & sPuntoVenta, "Modificación del punto de venta del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tHabitacion), "", RsPedido!tHabitacion)
        If valorActual <> sHabitacion Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA HABITACIÓN", sPedido, Error, "", "", "Modificación de Nro Habitación. Valor modificado:" & valorActual & " - Valor actual: " & sHabitacion, "Modificación del número de habitación del pedido (" & sPedido & ") ", sUsuario)
        End If
            
        valorActual = IIf(IsNull(RsPedido!TRESERVA), "", RsPedido!TRESERVA)
        If valorActual <> sReserva Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA RESERVA", sPedido, Error, "", "", "Modificación de Nro Reserva. Valor modificado:" & valorActual & " - Valor actual: " & sReserva, "Modificación del número de reserva del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tPasajero), "", RsPedido!tPasajero)
        If valorActual <> sPasajero Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PASAJERO", sPedido, Error, "", "", "Modificación de Pasajero. Valor modificado:" & valorActual & " - Valor actual: " & sPasajero, "Modificación del pasajero del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tCompania), "", RsPedido!tCompania)
        If valorActual <> sCompania Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA COMPANIA", sPedido, Error, "", "", "Modificación de Compania. Valor modificado:" & valorActual & " - Valor actual: " & sCompania, "Modificación de compania del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tContacto), "", RsPedido!tContacto)
        If valorActual <> sContacto Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA CONTACTO", sPedido, Error, "", "", "Modificación de contacto. Valor modificado:" & valorActual & " - Valor actual: " & sContacto, "Modificación del contacto del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!nDescuento), "", RsPedido!nDescuento)
        If valorActual <> xDescuento Or RsPedido!tDescuento <> sCodigoDescuento Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA DESCUENTO", sPedido, Error, "", "", "Modificación de Descuento. Valor modificado:" & valorActual & "(Cod. desct" & RsPedido!tDescuento & ") - Valor actual: " & xDescuento & "(Cod. Descuento" & sCodigoDescuento & ")", "Modificación de descuento del pedido (" & sPedido & ") ", sUsuario)
        End If
        
        
        valorActual = IIf(IsNull(RsPedido!tUsuarioDescuento), "", RsPedido!tUsuarioDescuento)
        If valorActual <> Mid(tAutorizaDescuento, 1, 15) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA AUTORIZA DESCUENTO", sPedido, Error, "", "", "Modifica autoriza descuento. Valor modificado:" & valorActual & " - Valor actual: " & Mid(tAutorizaDescuento, 1, 15), "Modificación de autoriza descuenta ( Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tObservacionDescuento), "", RsPedido!tObservacionDescuento)
        If valorActual <> IIf(sCodigoDescuento = "000", sDescripcionDescuento, "") Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA OBS DESCUENTO", sPedido, Error, "", "", "Modificación de obs. de descuento. Valor modificado:" & valorActual & " - Valor actual: " & IIf(sCodigoDescuento = "000", sDescripcionDescuento, ""), "Modificación de obs. de descuento (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tTienda), "", RsPedido!tTienda)
        If valorActual <> sTienda Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA TIENDA", sPedido, Error, "", "", "Modificación de tienda. Valor modificado:" & valorActual & " - Valor actual: " & sTienda, "Modificación de tienda (Ped. " & sPedido & ") ", sUsuario)
        End If
    
        valorActual = IIf(IsNull(RsPedido!fProgramacion), "", RsPedido!fProgramacion)
        If Format(valorActual, "dd/MM/yyyy HH:mm") <> IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:mm")) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA FECHA PROGRAMACIÓN", sPedido, Error, "", "", "Modificación fecha programación. Valor modificado:" & valorActual & " - Valor actual: " & IIf(txtFechaProg.Caption = "", Null, Format(txtFechaProg.Caption, "dd/MM/yyyy HH:mm")), "Modifica fecha programación (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tcodigopariente), "", RsPedido!tcodigopariente)
        If valorActual <> sCodigoParienteSeleccionado Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PARIENTE", sPedido, Error, "", "", "Modificación pariente. Valor modificado:" & valorActual & " - Valor actual: " & sCodigoParienteSeleccionado, "Modificación pariente (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        valorActual = IIf(IsNull(RsPedido!tentregara), "", RsPedido!tentregara)
        If valorActual <> IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20)) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA ENTREGAR A ", sPedido, Error, "", "", "Modificación entregar a. Valor modificado:" & valorActual & " - Valor actual: " & IIf(Len(txtEntregar.Caption) = 0, "", Left(Me.txtEntregar.Caption, 20)), "Modificación 'entregar a' (Ped. " & sPedido & ") ", sUsuario)
        End If
    
        valorActual = IIf(IsNull(RsPedido!nMinutosAntesEnvio), 0, RsPedido!nMinutosAntesEnvio)
        If valorActual <> val(txtEnvioAntes.Text) Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MINUTOS ANTES ENVIO", sPedido, Error, "", "", "Modificación de tienda. Valor modificado:" & valorActual & " - Valor actual: " & val(txtEnvioAntes.Text), "Modificación de minutos antes del envío (Ped. " & sPedido & ") ", sUsuario)
        End If
    
    
        valorActual = IIf(IsNull(RsPedido!CodigoOrigenVenta), "", RsPedido!CodigoOrigenVenta)
        If valorActual <> vOrigenVentas Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA ORIGEN VENTA", sPedido, Error, "", "", "Modificación de Origen Venta. Valor modificado:" & valorActual & " - Valor actual: " & vOrigenVentas, "Modificación de origen de venta (Ped. " & sPedido & ") ", sUsuario)
        End If
    
        valorActual = IIf(IsNull(RsPedido!tMesa), "", RsPedido!tMesa)
        If valorActual <> sMesa Then
            Call Log_Inforest("PUNTO VENTA", "MODIFICA MESA", sPedido, Error, "", "", "Modificación de mesa. Valor modificado:" & valorActual & " - Valor actual: " & sMesa, "Modificación de mesa (Ped. " & sPedido & ") ", sUsuario)
        End If
        
        If wAgrega Then
        
            Dim nSalon As String
            nSalon = Calcular("select tSalon as Codigo from TMESA where tCodigoMesa='" & sMesa & "' ", Cn)
            valorActual = IIf(IsNull(RsPedido!tSalon), "", RsPedido!tSalon)
            If valorActual <> sSalon Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA SALON", sPedido, Error, "", "", "Modificación de salon. Valor modificado:" & valorActual & " - Valor actual: " & sSalon, "Modificación de Salón (Ped. " & sPedido & ") ", sUsuario)
            End If
        
            valorActual = IIf(IsNull(RsPedido!tCaja), "", RsPedido!tCaja)
            If valorActual <> sCaja Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA NRO DE CAJA", sPedido, Error, "", "", "Modificación de NRO DE CAJA. Valor modificado:" & valorActual & " - Valor actual: " & sCaja, "Modificación de nro de caja (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!tTurno), "", RsPedido!tTurno)
            If valorActual <> sTurno Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA TURNO", sPedido, Error, "", "", "Modificación de turno. Valor modificado:" & valorActual & " - Valor actual: " & sTurno, "Modificación de Turno (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!tUsuario), "", RsPedido!tUsuario)
            If valorActual <> sUsuario Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA USUARIO", sPedido, Error, "", "", "Modificación de usuario. Valor modificado:" & valorActual & " - Valor actual: " & sUsuario, "Modificación de usuario (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!nAdulto), "", RsPedido!nAdulto)
            If valorActual <> val(txtAdulto.Text) Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA NRO ADULTO", sPedido, Error, "", "", "Modificación de nro de adulto. Valor modificado:" & valorActual & " - Valor actual: " & txtAdulto.Text, "Modificación de nro de adulto (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!nNino), "", RsPedido!nNino)
            If valorActual <> val(txtNino.Text) Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA NRO DE NINIO", sPedido, Error, "", "", "Modificación de nro de ninios. Valor modificado:" & valorActual & " - Valor actual: " & txtNino.Text, "Modificación de nro de ninios (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!nMesa), "", RsPedido!nMesa)
            If valorActual <> val(txtJuntar.Text) Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA JUNTAR MESA", sPedido, Error, "", "", "Modificación junta mesa. Valor modificado:" & valorActual & " - Valor actual: " & val(txtJuntar.Text), "Modificación junta mesa. (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!nTiempoDelivery), "", RsPedido!nTiempoDelivery)
            If valorActual <> nTiempoDelivery Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA TIEMPO DELIVERY", sPedido, Error, "", "", "Modificación tiempo delivery. Valor modificado:" & valorActual & " - Valor actual: " & nTiempoDelivery, "Modificación tiempo delivery (Ped. " & sPedido & ") ", sUsuario)
            End If
            
            valorActual = IIf(IsNull(RsPedido!fdiacontable), "", RsPedido!fdiacontable)
            If valorActual <> obtieneDiaContable Then
                Call Log_Inforest("PUNTO VENTA", "MODIFICA DIA CONTABLE", sPedido, Error, "", "", "Modificación del día contable. Valor modificado:" & valorActual & " - Valor actual: " & obtieneDiaContable, "Modificación del día contable (Ped. " & sPedido & ") ", sUsuario)
            End If
        
        End If
    End If
fin:
End Sub


Public Function VerificarAutorizacionEnvio(ByVal tClave As String) As String
    Dim RsUsuario As New ADODB.Recordset
    Dim RsMozo As New ADODB.Recordset
    Dim tCodigoPersona As String
    Dim tDescPersona As String
    Dim ttipoPersona As String 'USUARIO 'MESERO
    
    tDescPersona = ""
    ttipoPersona = ""
    
    Set RsUsuario = Lib.OpenRecordset("SELECT u.tCodigoUsuario, u.tResumido FROM TUSUARIO u INNER JOIN TGRUPOUSUARIO g ON u.tGrupoUsuario = g.tGrupoUsuario  WHERE u.tPassword = '" & Encapsula(tClave) & "' and u.lActivo = 1 and g.lopcion24 = 1", Cn)
    
    If RsUsuario.RecordCount > 0 Then
        If Not (RsUsuario.EOF Or RsUsuario.BOF) Then
            RsUsuario.MoveFirst
            tCodigoPersona = RsUsuario.Fields("tCodigoUsuario")
            tDescPersona = "U" + RsUsuario.Fields("tResumido")
            ttipoPersona = "USUARIO"
        End If
    End If
   
    Set RsMozo = Lib.OpenRecordset("SELECT * FROM vmozo WHERE tvalor='" & Encapsula(tClave) & "' and lactivo=1 and lAutorizaEnviosMozo=1", Cn)
    If RsMozo.RecordCount > 0 Then
        If Not (RsMozo.EOF Or RsMozo.BOF) Then
            RsMozo.MoveFirst
            tCodigoPersona = RsMozo.Fields("Codigo")
            tDescPersona = "M" + RsMozo.Fields("tResumido")
            ttipoPersona = Mesero
        End If
    End If
    
    VerificarAutorizacionEnvio = tDescPersona
  
End Function


Private Sub cmdOperador_Click(Index As Integer)

    HabilitaTimerColor (False)

   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 13
       cmdOperador(i).backColor = vbButtonFace
   Next i
   RsOperador.MoveFirst
   RsOperador.Find "nboton = " & Trim(str(Index))
        xOperador = RsOperador!codigo
   nOperadorPropiedad = RsOperador!nControl
   cmdOperador(Index).backColor = vbRed
   If wAgregaCombo Then
      AsignaComboPropiedad
   Else
      AsignaPropiedad
   End If
   Screen.MousePointer = vbDefault
   
   HabilitaTimerColor (True)
End Sub

Private Sub cmdOrdenCompra_Click()
On Error GoTo fin:
    wEnter = False
    frmKeyBoard.Caption = "OrdeCompra"
    frmKeyBoard.txtResultado.Text = Trim(txtOrdenCompra.Text)
    frmKeyBoard.Show vbModal
    If wEnter = True Then
       txtOrdenCompra.Text = Trim(sDescrip)
    End If
    
Exit Sub
fin:

End Sub

Private Sub cmdOrigen_Click(Index As Integer)
' origen de ventas
   HabilitaTimerColor (False)
    
   RsOrigenVentas.MoveFirst
   RsOrigenVentas.Find "boton = " & Trim(str(Index))
   vOrigenVentas = RsOrigenVentas!CodOrigenVenta
   'sMotorizado = RsMotorizado!codigo
   HabilitaTimerColor (True)
   validarOrigenVentas
      
End Sub

Private Function validarOrigenVentas()
' origen de ventas
    lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
    'entregarA
    lObligaEntregarA = IIf(IsNull(RsCanalesVenta!lObligaEntregarA), False, RsCanalesVenta!lObligaEntregarA)
    
    'origen de ventas
     lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    
    If lMCPV Then
        'lObligaMozo = False
        lActivaMozo = False
    End If
    
    lObligaMozo = IIf(IsNull(RsCanalesVenta!lObligaMozo), False, RsCanalesVenta!lObligaMozo)
    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
    lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
    lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
    lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
    lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
    lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
    
    If lActivaMozo Then
        fraMozo.Visible = True
        Me.fraMozo.Caption = Mesero
    Else
        fraMozo.Visible = False
    End If
                    
    If lActivaMotorizado Then
        fraMotorizado.Visible = True
    Else
        fraMotorizado.Visible = False
    End If
    
    Me.fraOrigenVentas.Visible = False
    
    
    
End Function

Private Sub cmdPedidoPagoEnviar_Click()
On Error GoTo fin:
    frmPedidoPagado.Visible = False
    If Trim(Me.lblPedidoPagado.Caption) <> "" Then
        sPedido = Trim(Me.lblPedidoPagado.Caption)
        Call EnviarPedido(Trim(Me.lblPedidoPagado.Caption))
    End If
    RsCabecera.Requery
    RsDetalle.Requery
    RsCombo.Requery
    LimpiaRs
    sPedido = ""
Exit Sub
fin:
    sPedido = ""
    RsCabecera.Requery
    RsDetalle.Requery
    RsCombo.Requery
    LimpiaRs

    MsgBox "Error: " & Error
End Sub

Private Sub EnviarPedido(ByVal CodigoPedido As String)
On Error GoTo fin
    

    If lOrden Then
       Isql = "select * from vPedido " & _
              "Where Codigo = '" & CodigoPedido & "' and nOrden in (select nOrden from DPEDIDO where tCodigoPedido='" & CodigoPedido & "' and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))) " & _
              "ORDER BY nOrden, tItem, nOrdenCombo,tetiqueta,combo " ' tItemCombo"
        
    Else
       Isql = "select * from vPedido " & _
              "Where Codigo = '" & CodigoPedido & "' And lImprimeArea = 1 and (lImprime = 0 or (isnull(lImprimeAreaCombo,0) = 1  and isnull(lImprimeCombo,0) = 0 ))" & _
              "ORDER BY nOrden, tItem, nOrdenCombo,tetiqueta,combo" 'tItemCombo"
    End If
    
    
    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
    If Not RsImpresion.EOF Then
        RsArea.MoveFirst
        Dim i As Integer
        For i = 1 To RsArea.RecordCount
            RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
            If RsArea!tIcono = "" Then
            'If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                End If
            End If
            RsArea.MoveNext
        Next i
    End If
    
    'KDS
    If lKDS Then
       Call KDS_AnadirNuevaOrden(RsCabecera, 0)
    End If
    
    'CESAR----CHEF CONTROL
    Dim ChefEnvio As Boolean
    ChefEnvio = Calcular("select ISNULL(lEnvioChef,0) as Codigo FROM TPARAMETRO", Cn)
    
    Cn.Execute "Update DPEDIDO Set lNoCantado=0 where tCodigoPedido = '" & CodigoPedido & "' and lImprime<>1"
    
    If ChefEnvio Then
        Cn.Execute "Update DPEDIDO Set lCantadoc=1,fCantadoC=GetDate(), lTipoEnvio=0 where tCodigoPedido = '" & CodigoPedido & "' and lImprime<>1"
    End If
    '---------------------------------
   
    Cn.Execute "update MPEDIDO set nReimpresion = isnull(nReimpresion,0) + 1  where tCodigoPedido='" & CodigoPedido & "'"
    Cn.Execute "Update DPEDIDO Set lImprime = 1, tUsuarioAutorizaEnvio = '" & sUsuario & "', tPerfilUsuarioAutoriza = '', fEnvio = getdate(), nEnvio = 1 where tCodigoPedido = '" & CodigoPedido & "' and lImprime<>1"
    Cn.Execute "Update CPEDIDO Set lImprime = 1 where tCodigoPedido = '" & CodigoPedido & "'"


Exit Sub
fin:
    MsgBox "Error: " & Error
End Sub


Private Sub cmdProducto_Click(Index As Integer)

    HabilitaTimerColor (False)
    
    txtBarra.SetFocus
    
    If RsCabecera.RecordCount = 0 Then
       Exit Sub
    End If
        
    If lMCPV Then
        If Not ValidaExistenciaProducto() Then
           MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
           RsCabecera.Requery
           RsDetalle.Requery
           Exit Sub
        End If
        '''''''''''''
        If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
           MsgBox "El Pedido ya fue facturado", vbExclamation, sMensaje
           RsCabecera.Requery
           Exit Sub
        End If
    Else
        If Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tEstadoPedido='01' and tCaja='" & sCaja & "' and tCodigoPedido='" & sPedido & "'", Cn) = 0 Then
           MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
           RsCabecera.Requery
           RsDetalle.Requery
           Exit Sub
        End If
    End If
    
    'JSJR 21/01/2025
    'cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
    xSubGrupo = RsSubgrupo!codigo
    RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
    
    'RsProducto.Requery
    RsProducto.MoveFirst
    RsProducto.Find "nboton = " & Trim(str(Index))
    sProducto = RsProducto!codigo
    
    
    'PROGRAMACION PRECIOS
    '' REQ REST-049-C-2022 - Osaka Argentina
    If lActivarIncrementoPrecios Then
        If (programacionPrecios("Venta")) Then
            Unload Me
            Exit Sub
        End If
    End If
    
    
    
    'INSUMOCRITICO23
    If validadIngresoProducto(sProducto) = False Then
             Exit Sub
    End If
    'INSUMOCRITICO23
    
    ListarOperadoresConFiltro (sProducto) 'Oscar Ortega
    If vComanda Then
       sTipo = "Comanda"
       sDescrip = sComanda
       frmNumPad.Caption = "Comanda"
       frmNumPad.Show vbModal
       
       sComanda = IIf(sDescrip = "" Or sDescrip = "0", "", Mid("0000000000", 1, 10 - Len(Trim(sDescrip))) & Trim(sDescrip))
       
      If wComanda And sComanda = "" Then
         MsgBox "No ha ingresado el número de la comanda", vbExclamation, sMensaje
         Exit Sub
      End If
      
      If Not wEnter Then
          Exit Sub
       End If
    End If
    
    
    
    
    
    If lPrecuenta Then
       sTemp = ""
       Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by Descripcion"
       Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                       "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
                    
       frmBusquedaRapida.nPredeterm = 1
       frmBusquedaRapida.Show vbModal
       If wEnter = True Then
          sMozoD = sCodigo
       Else
          sMozoD = sMozo
       End If
    End If
    
    If lBal And RsProducto!lBalanza Then
       Dim nResultado As Double
       nResultado = Pesar(nBalanzaPuerto)
       nResultado = Format(nResultado, "#,##0.00")
       If nResultado > 0 Then
          InsertaProducto nResultado
       End If
    Else
    nCantidad = 1
       InsertaProducto 1
    End If
    
    If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
       lPropiedad = True
    End If
    
    
    
    
'''      If (RsProducto!lCombinacion) And (RsProducto!lComboVisible) Then
'''    lComboActivo = True
'''''       MsgBox ("es un combo !!!"), vbInformation
''''        If RsCabecera.RecordCount = 0 Then
''''           Exit Sub
''''        End If
''''
''''        wAgregaCombo = True
''''        ActivaCabecera False
''''        ActivaOpcion False
''''        If Not RsCombo.EOF Then
''''           RsCombo.MoveFirst
''''        End If
''''        AsignaProductoCombo
''''
''''        CmdOpcion(1).Visible = False
''''        CmdOpcion(2).Visible = False
''''        txtBarra.SetFocus
'''''                 fraCombo.Visible = True
'''''                fraProductoCombo.Visible = True
''''        ActivaFrame fraCombo
''''        ActivaFrame fraProductoCombo
'''    End If
    
    HabilitaTimerColor (True)
    If lvisor Then
        Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
        Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sPedido & "',1)"
    End If
    
End Sub

Private Sub cmdProductoCombo_Click(Index As Integer)
    If lMCPV Then
        If Not ValidaExistenciaProducto() Then
            MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
            RsCabecera.Requery
            RsDetalle.Requery
            Exit Sub
        End If
    End If
    
    HabilitaTimerColor (False)
    
    txtBarra.SetFocus
    
    If RsCabecera.RecordCount = 0 Then
       Exit Sub
    End If
        
    Dim xxx As String
    xxx = RsProducto.Filter
    RsProducto.Filter = adFilterNone
    'RsProducto.Requery
    RsProducto.MoveFirst
    RsProducto.Find "tResumido = '" & cmdProductoCombo(Index).Caption & "'"
    sProducto = RsProducto!codigo

    nCCombo = Calcular("select sum(nCantidad) as Codigo " & _
                       "FROM dbo.CPEDIDO WHERE CPEDIDO.tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'", Cn)
    If nCCombo < nCombo * RsDetalle!nCantidad Then
       If lPrecuenta Then
          sTemp = ""
          Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by Descripcion"
          
          Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                          "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
          frmBusquedaRapida.nPredeterm = 1
          frmBusquedaRapida.Show vbModal
          If wEnter = True Then
             sMozoD = sCodigo
          Else
             sMozoD = sMozo
          End If
       End If
       'Oscar Ortega----------------------------------------------
        'Obtener Si es LUnico
        Dim oRsProductoDeCombo As Recordset
        Set oRsProductoDeCombo = Obtener_ProductoDeCombo(RsDetalle!tCodigoProducto, sProducto)
        If oRsProductoDeCombo.RecordCount > 0 Then
            If IIf(IsNull(oRsProductoDeCombo!lUnico), False, oRsProductoDeCombo!lUnico) Then
                 'Obtener Suma de cantidades
                 Dim nCantidadEnElCombo As Integer
                 nCantidadEnElCombo = ObtenerSumaCantidadesEnElCombo(sPedido, sItem, oRsProductoDeCombo!tEtiqueta)
                 'Suma de cantidades < que nCantidad
                 If nCantidadEnElCombo < nCantidad Then
                     InsertaCombo sProducto
                 Else
                     MsgBox "Solo es permitido " & nCantidad & " elemento(s) de tipo " & oRsProductoDeCombo!tEtiqueta, vbExclamation, sMensaje
                 End If
            Else
                InsertaCombo sProducto
            End If
        Else
            InsertaCombo sProducto
        End If
       '----------------------------------------------------------
       'InsertaCombo sProducto
    Else
       MsgBox "La cantidad máxima de items para este producto es de " & nCombo * RsDetalle!nCantidad, vbExclamation, sMensaje
    End If
    RsProducto.Filter = IIf(xxx = "0", "", xxx)

    HabilitaTimerColor (True)
End Sub

Private Sub cmdPropiedad_Click(Index As Integer)
   Dim ncantidadPropiedad As Double
   Dim Cantidad As Double

    If lMCPV Then
        If Not ValidaExistenciaProducto() Then
               MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
               RsCabecera.Requery
                RsDetalle.Requery
               Exit Sub
            End If
    End If

    HabilitaTimerColor (False)

    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    
    RsPropiedad.MoveFirst
    RsPropiedad.Find ("Descripcion = '" & cmdPropiedad(Index).Caption & "'")
    
    If Not (RsOperador.EOF Or RsOperador.BOF) Then
        nOperadorPropiedad = Calcular("select isnull(ncontrol,0) as codigo from voperador where codigo='" & RsOperador!codigo & "'", Cn)
    End If

    If cmdPropiedad(Index).FontBold = True Then
       cmdPropiedad(Index).FontBold = False
       If Not RsPropiedad.EOF Then
          If wAgregaCombo Then
            Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tcombopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and titemcombo='" & xItem & "' and  tproducto='" & sCombo & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' ", Cn)
            If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad), sPedido)
            End If
            Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and tProducto='" & sCombo & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "'"
          Else
            Cantidad = Calcular("select isnull(ncantidad,1) as codigo from tproductopropiedad where tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and tproducto='" & sProducto & "' and tcodigopropiedad='" & RsPropiedad!codigo & "' and tenlace='" & RsPropiedad!tEnlace & "'", Cn)
            If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta - (RsPropiedad!nPrecio * Cantidad), sPedido)
            End If
            Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tProducto='" & sProducto & "' and tCodigoPropiedad='" & RsPropiedad!codigo & "' and tEnlace='" & RsPropiedad!tEnlace & "'"
            Call Log_Inforest("PUNTO VENTA", "ELIMINA PROPIEDAD DE PRODUCTO DEL PEDIDO", sPedido, Error, "", "Eliminación de la propiedad " & RsPropiedad!codigo & " del producto " & sProducto & " (Ped. " & sPedido & ") ", "Eliminación de la propiedad " & RsPropiedad!codigo & "- Prod." & sProducto & " (Ped. " & sPedido & ")", "Se ha eliminado la propiedad del Pedido: " & sPedido, sUsuario)
          
          End If
          'lblResumen.Text = Replace(lblResumen.Text, RsOperador!Descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
          If Cantidad <> 1 Then
            lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ": (" & Cantidad & "), ", "")
          Else
            lblResumen.Text = Replace(lblResumen.Text, RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ", ", "")
          End If
          
       End If
    Else
    
            ncantidadPropiedad = 1
            If RsPropiedad!lsolicitacantidad = 1 Or RsPropiedad!lsolicitacantidad = True Then
                sTipo = "Prepintado"
                
                sCodigo = ncantidadPropiedad
                
                frmNumPad.Show vbModal
                If wEnter And val(sDescrip) > 0 Then
                    ncantidadPropiedad = sDescrip
                End If
            End If

       If nOperadorPropiedad > 0 Then
          If wAgregaCombo Then
             Isql = "SELECT COUNT(dbo.TCOMBOPROPIEDAD.tCodigoPedido) AS codigo " & _
                    "FROM dbo.TCOMBOPROPIEDAD INNER JOIN dbo.TPROPIEDAD ON dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TCOMBOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tItemCombo='" & xItem & "' and dbo.TCOMBOPROPIEDAD.tProducto='" & sCombo & "'  and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          Else
             Isql = "SELECT COUNT(dbo.TPRODUCTOPROPIEDAD.tCodigoPedido) AS codigo FROM dbo.TPRODUCTOPROPIEDAD INNER JOIN " & _
                    "dbo.TPROPIEDAD ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
                    "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "' and tOperador='" & RsOperador!codigo & "'"
             If nOperadorPropiedad <= Calcular(Isql, Cn) Then
                MsgBox "Ha llegado a la Cantidad máxima de " & nOperadorPropiedad & " Propiedad(es) por Operador", vbExclamation, sMensaje
                Exit Sub
             End If
          End If
       End If
        Dim maxped As Double
        maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
        If maxped > 0 Then
               If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + ((RsPropiedad!nPrecio * ncantidadPropiedad) * nCantidad)) Then
                 MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                 Exit Sub
               End If
        End If
       cmdPropiedad(Index).FontBold = True
       If Not RsPropiedad.EOF Then
          nInsumo = IIf(IsNull(RsPropiedad!nInsumo), 0, RsPropiedad!nInsumo)
          nGasto = IIf(IsNull(RsPropiedad!nGasto), 0, RsPropiedad!nGasto)
          nMObra = IIf(IsNull(RsPropiedad!nManoObra), 0, RsPropiedad!nManoObra)
          If wAgregaCombo Then
            Cn.Execute "Insert into TCOMBOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & xItem & "', '" & RsPropiedad!codigo & "', '" & sCombo & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1," & ncantidadPropiedad & " , " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & ", " & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & " , " & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
            If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad), sPedido)
            End If
          Else
            Cn.Execute "Insert into TPRODUCTOPROPIEDAD values ('" & sPedido & "', '" & sItem & "', '" & RsPropiedad!codigo & "', '" & sProducto & "', '" & RsPropiedad!tEnlace & "', " & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nInsumo * -1, ncantidadPropiedad * nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nGasto * -1, ncantidadPropiedad * nGasto) & "," & IIf(RsPropiedad!codigo = "9999", ncantidadPropiedad * nMObra * -1, ncantidadPropiedad * nMObra) & ", 1, " & ncantidadPropiedad & ",  " & IIf(RsPropiedad!codigo = "9999", nInsumo * -1, nInsumo) & "," & IIf(RsPropiedad!codigo = "9999", nGasto * -1, nGasto) & "," & IIf(RsPropiedad!codigo = "9999", nMObra * -1, nMObra) & ")"
            Call Log_Inforest("PUNTO VENTA", "MODIFICA PROPIEDAD DE PRODUCTO DEL PEDIDO", sPedido, Error, "", "Modifica propiedad " & RsPropiedad!codigo & " del producto " & sProducto & " (Ped. " & sPedido & ") ", "Modifica propiedad " & RsPropiedad!codigo & "- Prod." & sProducto & " (Ped. " & sPedido & ")", "Se ha modificado la propiedad del Pedido: " & sPedido, sUsuario)
                      
             If RsPropiedad!nPrecio <> 0 Then
                txtSuma.Caption = CambiaPrecio(nPVenta + (RsPropiedad!nPrecio * ncantidadPropiedad), sPedido)
             End If
          End If
          If ncantidadPropiedad <> 1 Then
                lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ": (" & ncantidadPropiedad & "), "
          Else
                lblResumen.Text = lblResumen.Text & RsOperador!descripcion & " " & cmdPropiedad(Index).Caption & ", "
          End If
       End If
    End If
    
       
    If wAgregaCombo Then
       RsComboPropiedad.Requery
    Else
       RsProductoPropiedad.Requery
    End If
    
    HabilitaTimerColor (True)
End Sub


Private Sub cmdPunto_Click(Index As Integer)
   HabilitaTimerColor (False)
   Dim i As Integer
   For i = 1 To 9
       cmdPunto(i).FontBold = False
   Next i
   cmdPunto(Index).FontBold = True
   rsPuntoVenta.MoveFirst
   rsPuntoVenta.Find "Descripcion = '" & cmdPunto(Index).Caption & "'"
   txtPuntoVenta.Caption = rsPuntoVenta!descripcion
   sPuntoVenta = rsPuntoVenta!codigo
   HabilitaTimerColor (True)
End Sub

Private Sub cmdSubGrupo_Click(Index As Integer)

   HabilitaTimerColor (False)

   Dim i As Integer
   Screen.MousePointer = vbHourglass
   For i = 1 To 7
       cmdSubGrupo(i).backColor = vbButtonFace
   Next i
   
   RsSubgrupo.MoveFirst
   RsSubgrupo.Find "nboton = " & Trim(str(Index))
   xSubGrupo = RsSubgrupo!codigo
   cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio
   Screen.MousePointer = vbDefault
   txtBarra.SetFocus
   
   HabilitaTimerColor (True)
End Sub



Private Sub cmdTipoDocumento_Click(Index As Integer)
 On Error GoTo fin

    xTotal = txtSuma
    
    Dim xUltimoCorrelativo As String
    
    nTotalPR = txtSuma.Caption
    
    Dim imgcodigobarra As Image
    
    HabilitaTimerColor (False)
    
    lDocumEquivaPrecuenta = False
    
    Dim RscadenaCodigoHash As Recordset
    
    Dim fso1 As Object
    
    Dim sSerie As String
    Dim sCorrela As String
    Dim sPrefijo As String
    Dim RsSuma As Recordset
    Dim sTipoDocumento As String
    Dim sImp As String
    Dim wConsumo As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim nRespuesta As Integer
    Dim nTotal As Double
    Dim sClienteFactura As String
    Dim CD2 As Boolean
    lImprimeAlternativa = False
    sUsuarioAutoriza = sUsuario
    
    'Ecuador
    Dim sAutorizacion As String
    Dim fInicio As Date
    Dim fCaducidad As Date
    
    Dim msgError As String
    msgError = ""
    
    wConsumo = False
    sDetalleConsumo = ""
    sVar1 = sTipoPedido
    txtBarra.SetFocus
    
    tAutorizacion = ""
    tCodigoControl = ""
    tDosificacion = ""
    tIdentidadNIT = ""
    
    
    lblPaso1.Visible = True
    lblPaso2.Visible = True
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    imgProceso(3).Visible = False
    FrameFeSpring.Visible = False
    
    
    
    'CAJA CONTINGENCIA -- JOSE LARIOS
    Isql = "select isnull(lCajaContingencia,0) as lCajaContingencia From TCAJA Where tCaja = '" & sCaja & "'"
    Set Rscajacontigencia = Lib.OpenRecordset(Isql, Cn)
    scajacontingencia = Rscajacontigencia!lCajaContingencia
    If scajacontingencia = True Then
        frmFechaEmision.Show vbModal
        sfemision = frmFechaEmision.femision
    Else
        sfemision = FechaServidor()
    End If

    'FIN DE CAJA CONTINGENCIA
    
    ' verifica
    If RsDetalle.RecordCount = 0 Then
        MsgBox "Error: No existen Productos a Facturar", vbInformation, sMensaje
        Exit Sub
    End If
    
        
    ' Miguel Mendoza - 10062019
    If Not VerificaComboMaximo() Then Exit Sub
    ' *************************
            
    'Donde se conteneran todas las validaciones del proceso de Documento del inforest
    If Validaciones("Documento") = False Then
        RsDetalle.Requery
        Exit Sub
    End If

    'Verifica si existe item por facturar
    X = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)
    If X <= 0 Then
       MsgBox "Error: No existen Productos a Facturar", vbCritical, sMensaje
'       Validaciones = False
       Exit Sub
    End If

    lCanalDelivery = Calcular("select v.lCanalDelivery As Codigo from vTipoPedido v INNER JOIN MPEDIDO m ON Codigo = m.tTipoPedido Where tCodigoPedido='" & sPedido & "'", Cn)
    
    'Chequea Descuento
    If Revisa_Descuento() = False Then
        Exit Sub
    End If
    
    xTiporecepcion = Calcular("Select isnull(tTiporecepcion,0) as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
    
    If (CD = True Or WEBAP = True Or xTiporecepcion) And xTiporecepcion <> 0 Then
        
        If xTiporecepcion = 1 Then
            
                RsCanalesVenta.Filter = "Codigo = '" & sTipoPedido & "'"
                lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                
                If lCanalDelivery Then
        
                    Dim CnCD As Connection
                    'Configuración
                    Dim sUserNameCD As String
                    Dim sUserPasswordCD As String
                
                    Set CnCD = New Connection
                    CnCD.Provider = "SQLOLEDB"
                    CnCD.CursorLocation = adUseServer
                    CnCD.ConnectionString = "User ID=" & sUserName & _
                                            ";password=" & sUserPassword & _
                                            ";Data Source=" & sRutaCD & _
                                            ";Initial Catalog=" & sMDBCD
                    CnCD.CommandTimeout = 250
                    CnCD.Open
            
                    Isql = "Select tCodigoPedidoCD From MPEDIDO where tCodigoPedido = '" + sPedido + "'"
                    
                    If Not IsNull(Lib.OpenRecordset(Isql, Cn)!tCodigoPedidoCD) Then
                        CD2 = True
                                                                                     
                                                                                                       
                                                                           
                        
                        'AJMM 21/08/2025
                        If Calcular("Select  count(*) as codigo from DPEDIDO Where tCodigoPedido='" & sPedido & "' and isnull(lRegistroVenta,0) = 0", Cn) >= 1 Then
                            tCodigoPedidoCD = Lib.OpenRecordset(Isql, Cn)!tCodigoPedidoCD
                            Set RsClienteSelecc = Nothing
                            Isql = "USP_CD_SELECCIONAR_MPEDIDO_ENVIADOS_X_CODIGO '" + tCodigoPedidoCD + "'"
                            Set RsClienteSelecc = Lib.OpenRecordset(Isql, CnCD)
                            
                            Select Case RsClienteSelecc!tEmision
                                      
                                               
                                                                                   
            
                                    
                                                             
                                            
                                      
                                Case "000" 'Cortesia
                                'GCAA 09032021
                                    If cmdTipoDocumento(Index).Caption = "CORTESIA" Or cmdTipoDocumento(Index).Caption = "CUENTA INTERNA" Then ''pp
                
                                    Else
                                        MsgBox ("Elija Cortesia.")
                                        Exit Sub
                                    End If
                                Case "001" 'Factura
                                    If cmdTipoDocumento(Index).Caption = "FACTURA" Then
                
                                    Else
                                        MsgBox ("Elija Factura.")
                                        Exit Sub
                                    End If
                                Case "002" 'Ticket o Boleta
                                    If cmdTipoDocumento(Index).Caption = "TICKET" Or cmdTipoDocumento(Index).Caption = "BOLETA" Then
                                        'MsgBox ("Elija Factura.")
                                        'Exit Sub
                                    Else
                                        MsgBox ("Elija Ticket o Boleta.")
                                        Exit Sub
                                    End If
                                Case Else
                                '...
                            End Select
                        Else
                        
                        End If
                    Else
                        CD2 = False
                    End If
                    
                Else
                    CD2 = False
                End If
        
        ElseIf xTiporecepcion = 2 Or xTiporecepcion = 3 Then
        
                RsCanalesVenta.Filter = "Codigo = '" & sTipoPedido & "'"
                lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
                CD2 = True
                Isql = "exec usp_WS_Inforest '" & sPedido & "', 4 "
                Set RsClienteSelecc = Lib.OpenRecordset(Isql, Cn)
                
                Select Case RsClienteSelecc!TipoDocumento
                    Case "00" 'Cortesia
                        'GCAA 09032021
                        If cmdTipoDocumento(Index).Caption = "CORTESIA" Or cmdTipoDocumento(Index).Caption = "CUENTA INTERNA" Then  ''pp
    
                        Else
                            MsgBox ("Elija Cortesia.")
                            Exit Sub
                        End If
                    Case "01" 'Factura
                        If cmdTipoDocumento(Index).Caption = "FACTURA" Or cmdTipoDocumento(Index).Caption = "FACTURA ELECTRONICA" Then
    
                        Else
                            MsgBox ("Elija Factura.")
                            Exit Sub
                        End If
                    Case "02" 'Ticket o Boleta
                        If cmdTipoDocumento(Index).Caption = "TICKET" Or cmdTipoDocumento(Index).Caption = "BOLETA" Or cmdTipoDocumento(Index).Caption = "BOLETA ELECTRONICA" Then
        
                        Else
                            MsgBox ("Elija Ticket o Boleta.")
                            Exit Sub
                        End If
                    Case Else

                End Select

        End If
               
    End If
    
    

    'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
    
    'JAPN 03012024 Mensaje de confirmacion de documentos
    If RsTipoDocumento!lConfirmacionDoc Then
        If MsgBox("Selecciono la opcion " & RsTipoDocumento!descripcion & ", desea continuar?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
           Exit Sub
        End If
    End If
    
    
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    Else
       xlTipoDocumento = Calcular("Select lValidaRuc As Codigo From TTIPODOCUMENTO Where tCodigoTipoDocumento = '" & RsTipoDocumento!TTipoEmision & "'", Cn)
    End If

 'validacionMontoMinimo
    nMontoPedidoFacturarMInimo = 0
    nMontoPedidoFacturarMInimo = Calcular("select isnull(nMontoMinimo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturarMInimo > 0 Then
        If nMontoPedidoFacturarMInimo >= CDbl(txtSuma.Caption) Then
            MsgBox "El Monto a Facturar no llega al Minimo Permitido al Tipo de Documento"
            Exit Sub
        End If
    End If

    'validacionMontoMaximo
    nMontoPedidoFacturar = 0
    nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturar > 0 Then
        If nMontoPedidoFacturar <= CDbl(txtSuma.Caption) Then
            MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
            Exit Sub
        End If
    End If
    
    
    If lInfhotel Then
        If rsPuntoVenta.BOF Or rsPuntoVenta.EOF Then
            MsgBox "No se ha Asignado un Punto de Venta de Infhotel, Verificar!!!"
            Exit Sub
        End If
    End If
    'Consistencia Cortesia
    sCortesia = ""
    If RsTipoDocumento!TTipoEmision = "00" Then
        If CD2 Then
            If Supervisor("04") = False Then
                MsgBox "Clave no permitida", vbExclamation, sMensaje
                Exit Sub
            End If
            sCodigo = RsClienteSelecc!tCortesia
            wEnter = True
        Else
            If Supervisor("04") = False Then
                MsgBox "Clave no permitida", vbExclamation, sMensaje
                Exit Sub
            End If

            sUsuarioAutoriza = sVar1
            sTemp = ""
            Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
            Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                            "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
            xtipodoc = RsTipoDocumento!prefijo
            frmBusquedaRapida.nPredeterm = 1
            frmBusquedaRapida.Show vbModal
            
        End If

        If wEnter = True Then
            sCortesia = sCodigo
        Else
            Exit Sub
        End If
    End If

    'CLAROSA 2014-MAYO
    If RsTipoDocumento!TTipoEmision = "00" And sCortesia <> "" Then
        Dim nTotalCortesiaActual As Double
        Dim nTopeCortesia As Double
        Dim nTotalDocActual As Double
        nTopeCortesia = Calcular("select isnull(tope,0) as codigo from vcortesia where codigo='" & sCortesia & "'", Cn)
        If nTopeCortesia > 0 Then
                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where ttipodocumento='00' and tcortesia='" & sCortesia & "' and YEAR(FREGISTRO)=YEAR(GETDATE()) AND month(fregistro)=month(getdate()) ", Cn)
                nTotalDocActual = Calcular("select sum(nVenta) as Codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)
                    If nTotalCortesiaActual + nTotalDocActual > nTopeCortesia Then
                            MsgBox "Con esta Emisión se supera el Tope Mensual asignado para la Cortesia " & UCase(sDescrip) & vbCrLf & "Tope Mensual: " & nTopeCortesia & ". Ya Asignado : " & nTotalCortesiaActual, vbCritical
                            Exit Sub
                    End If
        End If
    End If

    'Por Consumo
    'CENTRAL DELIVERY
    If CD2 Then
        If RsTipoDocumento!TTipoEmision <> "00" Then
            If lConsumo1 = True Then
                Isql = "Select tCodigoPedidoCD From MPEDIDO where tCodigoPedido = '" + sPedido + "'"
                If Not IsNull(Lib.OpenRecordset(Isql, Cn)!tCodigoPedidoCD) Then
                    Isql = "USP_CD_SELECCIONAR_MPEDIDO_ENVIADOS_X_CODIGO '" + Lib.OpenRecordset(Isql, Cn)!tCodigoPedidoCD + "'"
                    Set RsClienteSelecc = Lib.OpenRecordset(Isql, CnCD)
                    sDetalleConsumo = RsClienteSelecc!tConsumo
                    If sDetalleConsumo = "" Then
                        wConsumo = False
                    Else
                        wConsumo = True
                    End If
                End If
            End If
        End If
    Else
        If RsTipoDocumento!TTipoEmision <> "00" Then
           If lConsumo1 = True Then
              nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
              If nRespuesta = vbYes Then
                 frmKeyBoard.txtResultado = tTextoConsumo
                 frmKeyBoard.Show vbModal
                 If sDescrip = "" Or Not wEnter Then
                    MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
                    Exit Sub
                 End If
                 sDetalleConsumo = sDescrip
                 wConsumo = True
              ElseIf nRespuesta = vbCancel Then
                 Exit Sub
              End If
            End If
        End If
    End If
    
    'Verifica si no ha sido importado
    If lMCPV Then
       i = Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tCodigoPedido ='" & sPedido & "'", Cn)
    Else
       i = Calcular("select count(tCodigoPedido) as codigo from MPEDIDO where tCodigoPedido ='" & sPedido & "' and tCaja ='" & sCaja & "'", Cn)
    End If
    If i <= 0 Then
       MsgBox "Error: Pedido jalado por otra caja, favor de Refrescar", vbCritical, sMensaje
       Exit Sub
    End If
    
    TimpresionDolaresDelivery = False
    '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sCliente & "'", Cn) Then
        If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If

        '----------- VISOR 8 -----------
        If lvisor Then
             Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
             Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado,tipodocu,rucfactura) values(1,'" & sCaja & "','" & sPedido & "',2,'','')"
             'frmBusquedaRapida.tipoBusqueda = "VISOR"
         End If
        '------------------
    '---------- verificar cover y monto de pedido-------------------------------------------------------------
            If lcover And RsTipoDocumento!TTipoEmision <> "00" Then
                    Dim TotPedido As Double
                    Dim paxped As Integer
                    Dim difped As Double
                    Dim xxx As String
                    xxx = RsProducto.Filter
                    TotPedido = val(Calcular("select sum(nVenta) as Codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn))
                    paxped = Calcular("select nadulto as codigo from mpedido WHERE tCodigoPedido='" & sPedido & "'", Cn)
                    If (sMontoMinCover * paxped) > TotPedido Then
                     RsProducto.Filter = adFilterNone
                     RsProducto.MoveFirst
                     RsProducto.Find "Codigo = '" & sCodItemCover & "'"
                            If Not RsProducto.EOF Then
                            sProducto = RsProducto!codigo
                                difped = (sMontoMinCover * paxped) - TotPedido
                                InsertaProducto difped
                            Else
            
                            End If
                    End If
                
                End If
     
             '-- cover
            RsProducto.Filter = IIf(xxx = "0", "", xxx)
    
    '----VALIDA BLOQUEO DE DOCUMENTOS POR CONSUMO CON BONIFICACIONES / 03012020 / GCAA -------------------
    
    
    If lbloqueadocconsumo = True And wConsumo = True Then
    Dim Xx As Integer
        Xx = Calcular("select count( round(nPrecioVenta,2))as codigo from dpedido where round(nPrecioVenta,2) = 0 and tCodigoPedido ='" & sPedido & "'", Cn)
         ' X = Calcular("select sum(nprecioventa) from MPEDIDO where tCodigoPedido ='" & sPedido & "'", Cn)
        If Xx > 0 Then
            MsgBox "Información: No se puede emitir el documento por consumo con bonificaciones (item con precio 0)", vbInformation, sMensaje
            RsDetalle.Requery
            Exit Sub
        End If
        
    End If
    
        If lBloqueoAfectoInafecto And wConsumo Then
            Dim inafecto As Integer
            Dim afecto As Integer
            inafecto = Calcular("select count(*) as codigo from dpedido where round(nprecioimpuesto1,2)=0 and tCodigoPedido ='" & sPedido & "'", Cn)
            afecto = Calcular("select count(*) as codigo from dpedido where round(nprecioimpuesto1,2)>0 and tCodigoPedido ='" & sPedido & "'", Cn)
            
            If inafecto > 0 And afecto > 0 Then
                MsgBox "Bloqueo de emision: Documento por consumo con afectos e inafectos", vbInformation, sMensaje
                RsDetalle.Requery
                Exit Sub
            End If
            
        End If
             
    '------------------------------------------------------------------------------------------------------
    
    'TIPO EMISION
    If Not wConsumo And RsTipoDocumento!tFormulario <> "01" Then
       '----------------------------------------------------------------------
       '  ************************* Emisión Variable *************************
       '----------------------------------------------------------------------
    
       'Calcula el total de la cabecera
       nTotal = Calcular("select sum(nVenta) as Codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)
       
       If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nTotal Or RsTipoDocumento!Monto = 0) Then
            sTemp = ""
            Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sCliente & "','" & RsTipoDocumento!TTipoEmision & "', " & nTotal & ", '" & sCaja & "'"

            '------VALIDA CORREO----------
            sTipoDocum = RsTipoDocumento!TTipoEmision
            
            frmBusquedaRapida.cmdOpcion(1).Enabled = True
            frmBusquedaRapida.cmdOpcion(2).Enabled = True
            frmBusquedaRapida.cmdOpcion(3).Enabled = True
            frmBusquedaRapida.nPredeterm = 1
          
            Select Case pais ' ok
                Case "001" 'Bolivia
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                      "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                      "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
            End Select
          
            sTemp = Calcular("SELECT dbo.TCLIENTE.tIdentidad AS codigo FROM dbo.TDELIVERY INNER JOIN dbo.TCLIENTE ON dbo.TDELIVERY.tCodigoCliente = dbo.TCLIENTE.tCodigoCliente WHERE tCodigoDelivery ='" & sCliente & "'", Cn)
            sTemp = IIf(sTemp = "0", "", sTemp)
            sClienteFactura = ""
                If lvisor Then
                     'frmBusquedaRapida.sTipoDocumBusq = sTipoDocum
                     'frmBusquedaRapida.tipoBusqueda = "VISOR"
                End If
            xtipodoc = RsTipoDocumento!prefijo
            frmBusquedaRapida.Show vbModal
            
            
                '------- PARA VISOR DE 8 "
            If lvisor And sCodigo <> "" Then
                    Dim Ruc As String
'                    Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
'                    'Cn.Execute "update infovisor set estado=2 , TipoDocu='" & sTipoDocum & "', Rucfactura='" & sCodigo & "' where pedido = '" & sPedido & "'"
'                    RUC = Calcular(" SELECT TOP 1 TIDENTIDAD AS CODIGO  FROM TCLIENTE WHERE TCODIGOCLIENTE ='" & sCodigo & "'", Cn)
'                    Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado,tipodocu,rucfactura) values(1,'" & sCaja & "','" & sPedido & "',2,'" & sTipoDocum & "','" & RUC & "')"
'                FrmMensajeVisor.Show vbModal
'                If FrmMensajeVisor.RESULTADOVISOR = False Then
'                    If FrmMensajeVisor.OPERADORMENSAJE = 1 Then
'                         MsgBox ("!EL DOCUMENTO HA SIDO CANCELADO POR EL CLIENTE¡")
'                    End If
'                    Exit Sub
'                End If
            End If

            If lvisor And sCodigo = "" Then
                    Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
                    Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado,tipodocu,rucfactura) values(1,'" & sCaja & "','" & sPedido & "',4,'" & sTipoDocum & "','" & Ruc & "')"
            End If

            '--------------------
            If wEnter = True And sCodigo <> "" Then
                    If pais = "000" Then
                      'IMPRIME DNI
                      Dim RsTc As ADODB.Recordset
                      Set RsTc = New ADODB.Recordset
                      Set RsTc = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sCodigo & "'", Cn)
                      If Not (RsTc.EOF Or RsTc.BOF) Then
                          RsTc.MoveFirst
                          If RsTc.Fields(0) <> "ok" Then
                              MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                              Exit Sub
                          End If
                      End If
                    End If
                    
                    If lCanalDelivery Then
                       Cn.Execute "Update TDELIVERY set tCodigoCliente = '" & sCodigo & "' where tCodigoDelivery ='" & sCliente & "'"
                    End If
                    sTemp = sCliente
                    sClienteFactura = sCodigo
                    
                    '-------VALIDA CORREO------------
                    lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                    
                    If lValidaEmail = True Then
                      sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                    
                      If sEmail = "" Then
                         MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                         Exit Sub
                      End If
                    End If
            Else
               Exit Sub
            End If
       End If
              
       'Genera y Actualiza los Numero de Documento
       RsDetalle.MoveFirst
       
       If RsTipoDocumento!tFormulario = "03" Then
          nFItem = nItemV
       End If

       For i = 1 To IIf(X Mod nFItem = 0, Int(X / nFItem), Int(X / nFItem) + 1)
           RsTipoDocumento.Requery
           RsTipoDocumento.MoveFirst
           RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
           If RsTipoDocumento.EOF Then
              MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
              Exit Sub
           End If

           sSerie = RsTipoDocumento!tSerie
           sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
           sPrefijo = RsTipoDocumento!prefijo
           sTipoDocumento = RsTipoDocumento!TTipoEmision
           sImp = RsTipoDocumento!timpresora
           sDocumento = sPrefijo & sSerie & sCorrela
           sResumen = RsTipoDocumento!lResumen
                                        
            Select Case pais
                Case "001" 'Bolivia
                        tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                        tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                        If tAutorizacion <> "" And tDosificacion <> "" Then
                            Else
                            MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                            Exit Sub
                        End If
                        
                Case "002" 'Ecuador
                    tAutorizacion = RsTipoDocumento!tNumeroAutorizacion
                
                Case Else 'Peru
                    tAutorizacion = ""
                    tCodigoControl = ""
                    tDosificacion = ""
                    
            End Select
         

           'Genera el Detalle de DDOCUMENTO
           Dim xClave As String
           If lDocumentoAgrupado Then
                RsDetalle.Sort = "tCodigoProducto"
                Dim oCantReg As Integer
                oCantReg = 0
                Dim oRsAgrupado As Recordset
                Dim oRsProximoProducto As Recordset
                Do While oCantReg < nFItem ' oo + 1
                     Isql = "select tCodigoProducto From DPedido Where tDocumento = '" & sDocumento & "' Group By tCodigoProducto"
                     Set oRsAgrupado = Lib.OpenRecordset(Isql, Cn)
                     oCantReg = oRsAgrupado.RecordCount
                     If (oCantReg = nFItem) Then
                        Isql = "Select tCodigoProducto From DPEDIDO where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido = '" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0 or isnull(tfacturado,'')='') "
                        Set oRsProximoProducto = Lib.OpenRecordset(Isql, Cn)
                        oRsAgrupado.Find ("tCodigoProducto = '" & oRsProximoProducto!tCodigoProducto & "'")
                        If oRsAgrupado.EOF Then
                            oCantReg = oCantReg + 1
                        Else
                            xClave = RsDetalle!tItem
                            Isql = "Update DPEDIDO set tDocumento = '" & sDocumento & "' where tItem = '" & xClave & "' and tCodigoPedido = '" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0 or isnull(tfacturado,'')='') "
                            Cn.Execute Isql
                            RsDetalle.MoveFirst
                            RsDetalle.Find ("tItem ='" & xClave & "'")
                            RsDetalle.MoveNext
                            If RsDetalle.EOF Then
                                oCantReg = nFItem
                                i = IIf(X Mod nFItem = 0, Int(X / nFItem), Int(X / nFItem) + 1)
                            End If
                        End If
                     Else
                        xClave = RsDetalle!tItem
                        Isql = "Update DPEDIDO set tDocumento = '" & sDocumento & "' where tItem = '" & xClave & "' and tCodigoPedido = '" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0 or isnull(tfacturado,'')='') "
                        Cn.Execute Isql
                        RsDetalle.MoveFirst
                        RsDetalle.Find ("tItem ='" & xClave & "'")
                        RsDetalle.MoveNext
                        If RsDetalle.EOF Then
                            oCantReg = nFItem + 1
                            i = IIf(X Mod nFItem = 0, Int(X / nFItem), Int(X / nFItem) + 1)
                        End If
                     End If
                Loop
                RsDetalle.Sort = "tCodigoPedido , tItem"
           Else
                For j = 1 To nFItem
                    xClave = RsDetalle!tItem
                    Isql = "Update DPEDIDO set tDocumento = '" & sDocumento & "' where tItem = '" & xClave & "' and tCodigoPedido = '" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado)) = 0 or isnull(tfacturado,'')='') "
                    Cn.Execute Isql
                    RsDetalle.MoveFirst
                    RsDetalle.Find ("tItem ='" & xClave & "'")
                    RsDetalle.MoveNext
                    If RsDetalle.EOF Then
                       Exit For
                    End If
                Next j
           End If

            Set RsSuma = Lib.OpenRecordset("select  sum(nVenta) as nVenta from DPEDIDO where tDocumento ='" & sDocumento & "' group by tDocumento", Cn)
            If RsSuma.RecordCount > 0 Then
                nCargo = Round(RsSuma!nVenta, 2)
            Else
                Exit Sub
            End If
            Select Case pais ' ok
                Case "001"
                    tCodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sClienteFactura, nCargo)
            End Select
           
           ''ELDCQ (INSERTA DOCUMENTO)
           If Not INSERTADOCUMENTO(sPedido, sDocumento, sTipoDocumento, IIf(sClienteFactura = "", "", sClienteFactura), "01", sCaja, sTurno, sSalon, sUsuario, sUsuarioAutoriza, Format(obtieneDiaContable, "yyyy/MM/dd"), IIf(lAplicablePedido, "", sCodigoDescuento), sDetalleConsumo, IIf(TimpresionDolaresDelivery, 1, 0), tAutorizacion, tCodigoControl, "", Format("01/01/1900", "yyyy/mm/dd"), Format("01/01/1900", "yyyy/mm/dd"), "", 2, sfemision) Then
              GoTo fin:
           End If

           Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"

           wEnter = True
           Dim RsCantDocumentos As Recordset ' necesario uso verificar
           If lPagoAntesImpresion And lCanalDelivery = False Then
                If lPagoRapidoPV = False Then
                   sFormulario = ""
                   sTipo = ""
                   frmPago.Show vbModal
                Else
                   lModuloPago = "PuntoVenta"
                   frmPagoRapido.Show vbModal
                End If
                If wEnter = False Then
                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                    Cn.Execute "Update dPEDIDO set tfacturado = '', tdocumento = ''  where tdocumento ='" & sDocumento & "'"
                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    
                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
    
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Exit Sub
                End If
           End If

           If lPagoAntesImpresion = False Then

                   'Actualiza Base de Datos Detalle del Pedido
                   Cn.Execute "Update DPEDIDO set tFacturado = 'F' where tDocumento ='" & sDocumento & "'"
                   Cn.Execute "Update DPEDIDO set tFacturado = 'C' where tDocumento ='" & sDocumento & "' and len(ltrim(tCortesia)) = 4 "
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica=1 where tCodigoPedido ='" & sPedido & "'"
                   
                   'PARA NO FISCALES
                   Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & sPedido & "' and tDocumento ='" & sDocumento & "'"
                   
                   'Imprime documentos
                   If wConsumo = False And lDescripcionAlternativa = True Then
                      If validaImpresionAlternativa(sDocumento) = False Then
                            If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                                  lImprimeAlternativa = True
                            End If
                      End If
                   End If
                    '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                    ' ELDCQ 15/11/2017
                     If lImprimeAlternativa = False Then
                        If lDocumentoAgrupado Then
                            Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
                            'FACTURACION_E_PERU
                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',2"
                        Else
                            Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                            'FACTURACION_E_PERU
                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
                        End If
                    Else
                        If lDocumentoAgrupado Then
                            Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
                            'FACTURACION_E_PERU
                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',6"
                        Else
                            Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
                            'FACTURACION_E_PERU
                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
                        End If
                    End If
                    '---------------------------------------------------------------------------------------------------------
                    'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                    
                    If RsImpresion.RecordCount = 0 Then
                       LimpiaRs
                       MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                    Else
                        'SUNAT
                        numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                        codigoImpresora = sImp
                        'SUNAT--AGREGO CAMPOS A MDOCUMENTO
                        Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "
                       
                       'FACTURACION_E_PERU
                        xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                        xImpresioDE = Mid(sDocumento, 1, 1)
                   
                   
                            'FACTURACION_E_PERU
                            If pais = "000" Then
                               If lFacturacionE Then
                                     If lFEOfisis Then 'OFISIS
                                                '----CABECERA
                                                Set oComandoCabeceraOfisis = New clsComando
                                                If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                            
                                                If Not oComandoCabeceraOfisis.GetParamOK Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoCabeceraOfisis.ExecSP Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                '----FIRMA DOCUMENTO OFISIS
                                                If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                    Set oComandoFirmaDocumentoOfisis = New clsComando
                                                    If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                            
                                                    If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    
                                                    'VALIDAR RESPUESTA CODIGO DE BARRA
                                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                                    Sleep 2500
                                                    If lImpresionCodigoBarras Then
                                                        imageHash.DataField = "foto"
                                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                        Set imageHash.DataSource = RsCodigoHash
                                                    Else
                                                        cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                                    End If
                                                End If
                                     
                                     ElseIf lFESpring Then
                                     
                                     ElseIf lFECarbajal Then
                                            Label2.Caption = "   Proceso de envio de documento a InfoFact......."
                                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                Dim sImporteLetra As String
                                                sImporteLetra = "" 'NumeroCadena(str(RsImpDocumentoE!nVenta)) + " " + sMonedaN
                                                FrameFeSpring.Visible = True
                                                Sleep 1000
                                                If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                                        Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
                                                        Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                                        imgProceso(2).Visible = True
                                                        imgProceso(3).Visible = True
                                                        Sleep 1000
                                                        FrameFeSpring.Visible = False
                                                        Exit Sub
                                                 End If
                                                 imgProceso(0).Visible = True
                                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                 If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                                         Else
                                                             cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                                         End If
                                                     End If
                                                 End If
                                                 imgProceso(1).Visible = True
                                                 Sleep 1000
                                                 FrameFeSpring.Visible = False
                                            End If
                                       
                                     ElseIf lFEpape Then
                                     
                                     
                                     Else 'INFOFACT
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                             If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                                 'Exit Sub
                                             End If
'                                             If RsImpresion!Ruc <> "" Then
'                                                 If Not INSERTAFE(sDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                     'Exit Sub
'                                                 End If
'                                             End If
                                             'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                             fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                             If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                                     Else
                                                         cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                                     End If
                                                 End If
                                             End If
                                         End If
                                    End If
                                End If
                                              
                            End If
                            '---------------------------------------
                             
                            'Configura la Impresora
                            Imprimir (sImp)
                            Printer.FontName = sFont
                            Printer.FontBold = False
            
                            'FORMATO TICKET VARIABLE
                            If RsTipoDocumento!tFormulario = "03" Then
                                If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                                    If wConsumo Then
                                      ImprimeFacturaVariableConsumo RsImpresion, sDetalleConsumo, sEmpresa
                                    Else
                                      ImprimeFacturaVariable RsImpresion, sEmpresa
                                    End If
                                                    
                                   NFactura = sCorrela
                                   lblFactura.Caption = NFactura
                                ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                   If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                      ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                                   End If
                                Else
                                   ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                End If
                            
                            Else
                            
                                'FORMATO VARIABLE
                                If lFacturacionE And RsTipoDocumento!lFacturacionElectronica And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
                                    'FACTURACION_E_PERU
                                    'FORMATO A4
                                    If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                       ImprimeFormatoA
                                       Kill App.path & "\fact.bmp"
                                    Else
                                       ImprimeFormatoA
                                    End If
                                
                                ElseIf lFacturacionE And lFECarbajal Then
                                    ImprimeFormatoA
                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                    If fso1.FileExists(App.path & "\fact.bmp") Then
                                        Kill App.path & "\fact.bmp"
                                    End If
                                ElseIf lFacturacionE = False Then
                                    If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                                       If wConsumo Then
                                          ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                       Else
                                          ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                       End If
                                       NFactura = sCorrela
                                       lblFactura.Caption = NFactura
                                    Else
                                       If wConsumo Then
                                          ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                       Else
                                          ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                       End If
                                    End If
                                End If
                            End If
                            
                            If pais = "002" And lFEEcuador = False Then
                               sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                               GeneraFacturaElectronica sXML, sDocumento
                            End If

                        If lFEEcuador Then
                         If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
                             MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                         End If
                        End If
                    End If
            End If
       Next i
       
       'IMPRESION DE DOCUMENTOS CON PAGOS
       If lPagoAntesImpresion Then
            Set RsCantDocumentos = Lib.OpenRecordset("select distinct D.tDocumento from DDOCUMENTO D inner join MDOCUMENTO M ON M.tDocumento = D.tDocumento where D.tCodigoPedido ='" & sPedido & "' and M.tEstadoDocumento <> '04'", Cn)
            For rdi = 0 To RsCantDocumentos.RecordCount - 1
                   sDocumento = RsCantDocumentos!tDocumento
                   
                   'Actualiza Base de Datos Detalle del Pedido
                   Cn.Execute "Update DPEDIDO set tFacturado = 'P' where tDocumento ='" & RsCantDocumentos!tDocumento & "'"
                   Cn.Execute "Update DPEDIDO set tFacturado = 'C' where tDocumento ='" & RsCantDocumentos!tDocumento & "' and len(ltrim(tCortesia)) = 4 "
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica=1 where tCodigoPedido ='" & sPedido & "'"
                   'PARA NO FISCALES
                   Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where  tCodigoPedido ='" & sPedido & "' and tDocumento ='" & RsCantDocumentos!tDocumento & "'"
                   
                    '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                    ' ELDCQ 15/11/2017
                     If lImprimeAlternativa = False Then
                        If lDocumentoAgrupado Then
                            Isql = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',1"
                            'FACTURACION_E_PERU
                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',2"
                        Else
                            Isql = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',3"
                            'FACTURACION_E_PERU
                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',4"
                        End If
                    Else
                        If lDocumentoAgrupado Then
                            Isql = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',5"
                           'FACTURACION_E_PERU
                            'IsqlFact = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',6"
                        Else
                            Isql = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',7"
                            'FACTURACION_E_PERU
                            'Isql = "EXEC usp_Inforest_Impresion '" & RsCantDocumentos!tDocumento & "',8"
                        End If
                    End If
                    '-----------------------------------------------------------------------------------------------------------
                    'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
                    Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                    
                    If RsImpresion.RecordCount = 0 Then
                       LimpiaRs
                       MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
                    Else
                        'SUNAT
                        numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
                        codigoImpresora = sImp
                        'SUNAT--AGREGO CAMPOS A MDOCUMENTO
                        Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & RsCantDocumentos!tDocumento & "' "
                       
                       'FACTURACION_E_PERU
                        xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & RsCantDocumentos!tDocumento & "')", Cn)
                        xImpresioDE = Mid(RsCantDocumentos!tDocumento, 1, 1)
                   
                   
                            'FACTURACION_E_PERU
                            If pais = "000" Then
                               If lFacturacionE Then
                               
                                     If lFEOfisis Then 'OFISIS
                                                '----CABECERA
                                                Set oComandoCabeceraOfisis = New clsComando
                                                If Not oComandoCabeceraOfisis.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                oComandoCabeceraOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, RsCantDocumentos!tDocumento
                            
                                                If Not oComandoCabeceraOfisis.GetParamOK Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                If Not oComandoCabeceraOfisis.ExecSP Then
                                                     Set oComandoCabeceraOfisis = Nothing
                                                     Exit Sub
                                                End If
                                                
                                                '----FIRMA DOCUMENTO OFISIS
                                                If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                                    Set oComandoFirmaDocumentoOfisis = New clsComando
                                                    If Not oComandoFirmaDocumentoOfisis.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    oComandoFirmaDocumentoOfisis.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, RsCantDocumentos!tDocumento
                            
                                                    If Not oComandoFirmaDocumentoOfisis.GetParamOK Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    If Not oComandoFirmaDocumentoOfisis.ExecSP Then
                                                         Set oComandoFirmaDocumentoOfisis = Nothing
                                                         Exit Sub
                                                    End If
                                                    
                                                    'VALIDAR RESPUESTA CODIGO DE BARRA
                                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(RsCantDocumentos!tDocumento, 8, 8)))
                                                    Sleep 2500
                                                    If lImpresionCodigoBarras Then
                                                        imageHash.DataField = "foto"
                                                        Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(RsCantDocumentos!tDocumento, 1, 1) & "','' ", Cn)
                                                        Set imageHash.DataSource = RsCodigoHash
                                                    
                                                    ElseIf lQRFE Then
                                                        Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                                    Else
                                                        Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                        If RscadenaCodigoHash.RecordCount > 0 Then
                                                            cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                        End If
                                                    End If
                                                End If
                                                    
                                     ElseIf lFESpring Then
                                                    
                                     ElseIf lFECarbajal Then
                                            'ElseIf lFECarbajal Then
                                            Label2.Caption = "   Proceso de envio de documento a InfoFact......."
                                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                                'Dim sImporteLetra As String
                                                sImporteLetra = NumeroCadena(str(RsImpresion!nVenta)) + " " + sMonedaN
                                                FrameFeSpring.Visible = True
                                                lblPaso1.Visible = True
                                                lblPaso2.Visible = True
                                                imgProceso(0).Visible = False
                                                imgProceso(1).Visible = False
                                                imgProceso(2).Visible = False
                                                imgProceso(3).Visible = False
                                                Sleep 1000
                                                If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                                        Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
                                                        Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                                        imgProceso(2).Visible = True
                                                        imgProceso(3).Visible = True
                                                        Sleep 1000
                                                        FrameFeSpring.Visible = False
                                                        Exit Sub
                                                 End If
                                                 imgProceso(0).Visible = True
                                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                                 If tCodigoFE = "000" Then
                                                     If lQRFE Then
                                                         Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                                     Else
                                                         If lImpresionCodigoBarras Then
                                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                                         Else
                                                             cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                                         End If
                                                     End If
                                                 End If
                                                 imgProceso(1).Visible = True
                                                 Sleep 1000
                                                 FrameFeSpring.Visible = False
                                            End If
                                                    
                                     ElseIf lFEpape Then
                                     
                                     
                                     Else 'INFOFACT
                                     
                                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                              If Not INSERTAFE(RsCantDocumentos!tDocumento, "", 1, "") Then '----CABECERA
                                                 'Exit Sub
                                             End If
'                                             If RsImpresion!Ruc <> "" Then
'                                                 If Not INSERTAFE(RsCantDocumentos!tDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                                     Exit Sub
'                                                 End If
'                                             End If
                                             'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                             fDocumento = Mid(RsCantDocumentos!tDocumento, 1, 1) + Mid(RsCantDocumentos!tDocumento, 4, 3) + Mid(RsCantDocumentos!tDocumento, 8, 8)
                                             If tCodigoFE = "000" Then
                                                 If lQRFE Then
                                                     Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                                 Else
                                                     If lImpresionCodigoBarras Then
                                                         Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                                     Else
                                                         cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                                     End If
                                                 End If
                                             End If
                                         End If
                                    End If
                                End If
                            End If
                            '---------------------------------------
                             
                            'Configura la Impresora
                            Imprimir (sImp)
                            Printer.FontName = sFont
                            Printer.FontBold = False
            
                            'FORMATO TICKET VARIABLE
                            If RsTipoDocumento!tFormulario = "03" Then
                                If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                                    If wConsumo Then
                                      ImprimeFacturaVariableConsumo RsImpresion, sDetalleConsumo, sEmpresa
                                    Else
                                      ImprimeFacturaVariable RsImpresion, sEmpresa
                                    End If
                                                    
                                   NFactura = sCorrela
                                   lblFactura.Caption = NFactura
                                ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                                   If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                      ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                                   End If
                                Else
                                   ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                End If
                                   
                            Else
                            
                                'FORMATO VARIABLE
                                If lFacturacionE And RsTipoDocumento!lFacturacionElectronica And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
                                            'FACTURACION_E_PERU
                                            'FORMATO A4
                                            If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                               ImprimeFormatoA
                                               Kill App.path & "\fact.bmp"
                                            Else
                                               ImprimeFormatoA
                                            End If
                                        
                                ElseIf lFacturacionE And lFECarbajal Then
                                        ImprimeFormatoA
                                        Set fso1 = CreateObject("Scripting.FileSystemObject")
                                        If fso1.FileExists(App.path & "\fact.bmp") Then
                                            Kill App.path & "\fact.bmp"
                                        End If
                                        
                                ElseIf lFacturacionE = False Then
                                        If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                                           If wConsumo Then
                                              ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                           Else
                                              ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                                           End If
                                           NFactura = sCorrela
                                           lblFactura.Caption = NFactura
                                        Else
                                           If wConsumo Then
                                              ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                           Else
                                              ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                                           End If
                                        End If
                                        
                                End If
                                
'                                If pais = "002" Then
'                                   sXML = Trim(LeerIni(App.Path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
'                                   GeneraFacturaElectronica sXML, RsCantDocumentos!tDocumento
'                                End If
                                If pais = "002" And lFEEcuador = False Then
                                   sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
                                   GeneraFacturaElectronica sXML, RsCantDocumentos!tDocumento
                                End If

                                        
                            End If
                                                            
                        If lFEEcuador Then
                         If INSERTA_FE_INFOREST(RsCantDocumentos!tDocumento, 1, DateTime.Now) = False Then
                             MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
                         End If
                        End If
                      
                    End If
                    
                 RsCantDocumentos.MoveNext
            Next rdi
       End If


       If i - 1 = 1 Then
            LimpiaRs
            'Cancelacion del Documento
            If RsTipoDocumento!TTipoEmision <> "00" Then
               xTipo = ""
               If lCanalDelivery Then
                  frmPrePago.Show vbModal
                  If wEnter Then
                     cmdOpcion_Click (5)
                  End If
               Else
                  If lPagoAntesImpresion = False Then
                      If lPagoRapidoPV = False Then
                         sFormulario = ""
                         sTipo = ""
                         frmPago.Show vbModal
                      Else
                         lModuloPago = "PuntoVenta"
                         frmPagoRapido.Show vbModal
                      End If
                  End If
               End If
               
               If wEnter Then
                  'Actualiza Base de Datos Cabecera del Pedido
                  Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
                  Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica=1 where tCodigoPedido ='" & sPedido & "'"
                 
                  'Juntar Mesa
                  Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
               Else
                  Cn.Execute "Update TMESA set tEstadoMesa = '05' where tCodigoMesa ='" & sMesa & "'"
               End If
            Else
               'Libera la Mesa
               Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
               
               'Juntar Mesa
               Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
            End If
            
            RsCabecera.Requery
            RsDetalle.Requery
            RsCombo.Requery
       Else
            If lPagoAntesImpresion Then
                'Libera la Mesa
                Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
               
                'Juntar Mesa
                Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
                RsCabecera.Requery
                RsDetalle.Requery
                RsCombo.Requery
            Else
                Cn.Execute "Update TMESA set tEstadoMesa = '05' where tCodigoMesa ='" & sMesa & "'"
                frmDocumento.Show vbModal
            End If
       End If

    
    Else
       '---------------------------------------------------------------------------
       '  ************************* EMISION POR TICKETERA *************************
       '---------------------------------------------------------------------------
       
       'JAPN 08/09/2023
        Dim tDocumentoVenta As String
       
       ' JAPN 8/09/2023
        tDocumentoVenta = IIf(IsNull(RsTipoDocumento!tCodigoExterno), "", RsTipoDocumento!tCodigoExterno)
         
       'Calcula el total de la cabecera
       nTotal = Calcular("select sum(nVenta) as Codigo from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0) and tCodigoPedido ='" & sPedido & "'", Cn)
       
       'erick 26082021
       sClienteFactura = ""
       
       If lActivaAnticipo Then ' RsTipoDocumento!TTipoEmision
                Dim CodigoReservaAnticipo As String
                CodigoReservaAnticipo = Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
                If CodigoReservaAnticipo <> "" Then
                    Dim valorResultado As String
                    valorResultado = Calcular(" exec sp_AsignaAnticipo_Pedido '" & sPedido & "','" & CodigoReservaAnticipo & "',3", Cn)
                    If valorResultado <> "0" Then
                        MsgBox valorResultado, vbExclamation, sMensaje
                        Exit Sub
                    End If
                    
                    Dim valorResultadoTipoDoc As String
                    valorResultadoTipoDoc = Calcular(" exec sp_AsignaAnticipo_Pedido '" & RsTipoDocumento!TTipoEmision & "','" & CodigoReservaAnticipo & "',4", Cn)
                    If valorResultadoTipoDoc <> "0" Then
                        MsgBox valorResultadoTipoDoc, vbExclamation, sMensaje
                        Exit Sub
                    End If
                    
                    
                End If
       End If
       
       
       
       If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nTotal Or RsTipoDocumento!Monto = 0) Then
            If (CD2 = True Or WEBAP = True) And xTiporecepcion <> 0 Then
            
               If xTiporecepcion = 1 Then
               'AJMM 21/08/2025
                    If Calcular("Select  count(*) as codigo from DPEDIDO Where tCodigoPedido='" & sPedido & "' and isnull(lRegistroVenta,0) = 0", Cn) >= 1 Then
                        sClienteFactura = ObtenerClienteFacturaCD(sCliente, sPedido)
                    Else
                        sClienteFactura = ""
                    End If
               ElseIf xTiporecepcion = 2 Or xTiporecepcion = 3 Then
                  sClienteFactura = ObtenerClienteFacturaWebAp(sPedido)
               End If
           End If
          
          If sClienteFactura = "" Then
                'Consistencia Factura
                frmBusquedaRapida.nPredeterm = 1
                Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sCliente & "','" & RsTipoDocumento!TTipoEmision & "', " & nTotal & ", '" & sCaja & "'"
                sTemp = ""
                If Calcular("select COUNT(*) as codigo from TDELIVERYCLIENTE Where tCodigoDelivery='" & sCliente & "'", Cn) = 0 Then
                   frmBusquedaRapida.nPredeterm = 1
                End If
                
                Select Case pais 'ok
                    Case "001" 'Bolivia
                    
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                         "Ident", 2, "tIdentidad", 1800, 2, 0, "", _
                                                                         "Cliente", 2, "Descripcion", 5200, 0, 0, "")
                    Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
                End Select
                
                '------VALIDA CORREO----------
                sTipoDocum = RsTipoDocumento!TTipoEmision
                
                frmBusquedaRapida.cmdOpcion(1).Enabled = True
                frmBusquedaRapida.cmdOpcion(2).Enabled = True
                frmBusquedaRapida.cmdOpcion(3).Enabled = True
                xtipodoc = RsTipoDocumento!prefijo
                
                'JAPN 8 / 9 / 2023 para solicitar datos de boleta
                 
                Dim sBoleta As Boolean
                sBoleta = Calcular("select ISNULL(lBODato,0) as Codigo FROM TPARAMETRO", Cn)
            
                If sBoleta = "Verdadero" Then
                    If xtipodoc = "B" And tDocumentoVenta = "03" Then
                        'Me.FrameSolicitudDatos.Visible = True
                        frmSolicitudBoleta.Show vbModal
                    Else
                        frmBusquedaRapida.Show vbModal
                    End If
                Else
                     frmBusquedaRapida.Show vbModal
                End If
                    
                    
                
            '------- PARA VISOR DE 8 "
            If lvisor And sCodigo = "" Then
                Dim RUC2 As String
                Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
                Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado,tipodocu,rucfactura) values(1,'" & sCaja & "','" & sPedido & "',4,'" & sTipoDocum & "','" & RUC2 & "')"
            End If
            '--------------------
                If wEnter = True And sCodigo <> "" Then
                      sTemp = sCliente
                      sClienteFactura = sCodigo
            
                      If pais = "000" Or pais = "003" Then
                      
                          'IMPRIME DNI
                          Dim RsTc1 As ADODB.Recordset
                          Set RsTc1 = New ADODB.Recordset
                          Set RsTc1 = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sClienteFactura & "'", Cn)
                          If Not (RsTc1.EOF Or RsTc1.BOF) Then
                              RsTc1.MoveFirst
                              If RsTc1.Fields(0) <> "ok" Then
                                  MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                                  Exit Sub
                              End If
                          End If
                      
                      End If
                  
                      lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                      
                      If lValidaEmail = True Then
                         sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                      
                         If sEmail = "" Then
                            MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                            Exit Sub
                         End If
                      End If
                        
                    If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True Then
                        Dim TempUbigeo As String
                        Dim TempUrbaniza As String
                        TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                        If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
                            MsgBox "El cliente no tiene Ubigeo ó Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
                            Exit Sub
                        End If
                    End If

                      Cn.Execute "usp_Inforest_ValidaDeliveryCliente '" & sCliente & "','" & sClienteFactura & "'"
                Else
                   RsDetalle.MoveFirst
                   Exit Sub
                End If
                
            End If
          
       Else
          sClienteFactura = ""
       End If
       

        'Genera y Actualiza los Numero de Documento
        sSerie = RsTipoDocumento!tSerie
        sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
        sPrefijo = RsTipoDocumento!prefijo
        sTipoDocumento = RsTipoDocumento!TTipoEmision
        sImp = RsTipoDocumento!timpresora
        sDocumento = sPrefijo & sSerie & sCorrela
        sResumen = RsTipoDocumento!lResumen
              
'        Set RsSuma = Lib.OpenRecordset("select sum(nVenta) as nVenta from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0 OR ISNULL(tFacturado,'')='') and tCodigoPedido ='" & sPedido & "' group by tCodigoPedido ", Cn)
'        nCargo = Round(RsSuma!nVenta, 2)
        nCargo = SumPrecioVenta(sPedido)
        Select Case pais
            Case "001" 'Bolivia
                    tAutorizacion = obtieneAutorizacionDosificacion(sCaja, "1")
                    tDosificacion = obtieneAutorizacionDosificacion(sCaja, "2")
                    If tAutorizacion <> "" And tDosificacion <> "" Then
                        tCodigoControl = devuelveCodigoControl(sCaja, sCorrela, tAutorizacion, tDosificacion, sClienteFactura, nTotal)
                        If tCodigoControl = "" Then: MsgBox "Error al generar Código de Control", vbCritical, sMensaje: Exit Sub
                    Else
                           MsgBox "Error al obtener Número de Autorización o Dosificación. Verifique.", vbCritical, sMensaje
                           Exit Sub
                    End If
                    
            Case "002" 'Ecuador
                tAutorizacion = IIf(IsNull(RsTipoDocumento!tNumeroAutorizacion), "", RsTipoDocumento!tNumeroAutorizacion)
            Case Else 'Peru, Ecuador
                tCodigoControl = ""
                tDosificacion = ""
        End Select
        
        If sClienteFactura = "" And sTipoDocumento = "02" And RsTipoDocumento!lFacturacionElectronica = True Then
          If sCodigoClienteGeneral = "" Or sCodigoClienteGeneral = "0" Then
                MsgBox "Codigo de cliente general no configurado en el INI.", vbInformation, sMensaje
                Exit Sub
          Else
            sClienteFactura = sCodigoClienteGeneral
          End If
            
            
            
'            If txtTipoDocumento.Caption = "" Then
'                MsgBox "No se encontro el documento configurado con los parametros para emision sin documento!!!.", vbInformation, sMensaje
'                Exit Sub
'            End If
        End If
        ''ELDCQ (INSERTA DOCUMENTO)
        If Not INSERTADOCUMENTO(sPedido, sDocumento, sTipoDocumento, IIf(sClienteFactura = "", "", sClienteFactura), "01", sCaja, sTurno, sSalon, sUsuario, sUsuarioAutoriza, Format(obtieneDiaContable, "yyyy/MM/dd"), IIf(lAplicablePedido, "", sCodigoDescuento), sDetalleConsumo, IIf(TimpresionDolaresDelivery, 1, 0), tAutorizacion, tCodigoControl, sCortesia, Format(IIf(IsNull(RsTipoDocumento!fInicio), "01/01/1900", RsTipoDocumento!fInicio), "yyyy/mm/dd"), Format(IIf(IsNull(RsTipoDocumento!fCaducidad), "01/01/1900", RsTipoDocumento!fCaducidad), "yyyy/mm/dd"), tContribuyenteEspecial, 1, sfemision) Then
            GoTo fin
        End If

        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
        
        If lIntVisaNetQR And sTipoDocumento <> "00" Then
            Set ImageIntegracion = Nothing
            If (EjecutaMotorIntegraciones(sDocumento, sUsuario, 1, 1) = True) Then
                'Sleep (500)
                Dim Ob As New Scripting.FileSystemObject
                If Ob.FileExists(lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp") Then
                    Set ImageIntegracion.Picture = LoadPicture(lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp")
                    Kill lRutaImagenIntegraciones + "\VISAQR" + sRUC + sCaja + sDocumento + ".bmp"
                Else
                 Set ImageIntegracion = Nothing
                End If
            Else
                MsgBox "No se pudo generar el QR Visa Net, el sistema continuara con la operacion!!!", vbInformation
                Set ImageIntegracion = Nothing
            End If
        Else
            Set ImageIntegracion = Nothing
        End If

        'GCAA 03/04/2020 boleta
        'GCAA 08012021 PAGO ANTES DE IMPRESION
        If lPagoAntesImpresion = False Then
            If lIntMercadoPagoQR And sTipoDocumento <> "00" Then
                Set ImageMercadoPago = Nothing
                If (EjecutaMotorIntegraciones(sDocumento, sUsuario, 5, 1) = True) Then
                    'Sleep (500)
                    Dim Obx As New Scripting.FileSystemObject
                    If Obx.FileExists(lRutaImagenIntegraciones + "\MERCADOPAGOQR" + sRUC + sCaja + sDocumento + ".bmp") Then
                        Set ImageMercadoPago.Picture = LoadPicture(lRutaImagenIntegraciones + "\MERCADOPAGOQR" + sRUC + sCaja + sDocumento + ".bmp")
                        Kill lRutaImagenIntegraciones + "\MERCADOPAGOQR" + sRUC + sCaja + sDocumento + ".bmp"
                    Else
                     Set ImageMercadoPago = Nothing
                    End If
                Else
                    MsgBox "No se pudo generar el QR Mercado Pago, el sistema continuara con la operacion!!!", vbInformation
                    Set ImageMercadoPago = Nothing
                End If
            Else
                Set ImageMercadoPago = Nothing
            End If
        End If

        wEnter = True
        If lPagoAntesImpresion Then
              'Cancelacion del Documento
              If RsTipoDocumento!TTipoEmision <> "00" Then
                   Cn.Execute "Update TMESA set tEstadoMesa = '05' where tCodigoMesa ='" & sMesa & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica = 1  where tCodigoPedido ='" & sPedido & "'"
                   xTipo = ""
              
                   If (CD2 = True Or WEBAP = True) And xTiporecepcion <> 0 Then
                                                              
                           If xTiporecepcion = 1 Then
                               Isql = "USP_CD_SELECCIONAR_DPREPAGO_X_PEDIDO '" & tCodigoPedidoCD & "',''"
                               Set RsRecord = Lib.OpenRecordset(Isql, CnCD)
                           ElseIf xTiporecepcion = 2 Or xTiporecepcion = 3 Then
                               Isql = " exec usp_WS_Inforest '" & sPedido & "' ,1" '
                               Set RsRecord = Lib.OpenRecordset(Isql, Cn)
                           End If
                             
                           If xTiporecepcion <> 0 Then 'PrePago ' sVar1 = "02"
                                   If xTiporecepcion = 1 Then
                                           Isql = "Select tCodigoPedidoCD From MPEDIDO WHERE tCodigoPedido = '" + sPedido + "'"
                                           Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
                                           If Not RsCodigoPCD.EOF Then
                                             If Not IsNull(RsCodigoPCD!tCodigoPedidoCD) Then
                                                 Call ModifcarEstadoDeliveryCabecera(RsCodigoPCD!tCodigoPedidoCD, "3", txtMotorizado.Caption)
                                             End If
                                           End If
        
                                           Call frmPrePago.Form_Load
                                           Do While Not RsRecord.EOF
                                               If RsRecord!tMONEDA = "01" Then 'Soles
                                                   Select Case RsRecord!tTipoPago 'pp
                                                       Case "01"
                                                           If Trim(RsRecord!tTarjeta) = "Efectivo" Then
                                                               frmPrePago.PnEN = RsRecord!nMonto
                                                           End If
                                                       Case "02"
                                                           If Not Trim(RsRecord!tCodigoTarjeta) = "" Then
                                                           'GCAA 30072020 - 04092020
                                                               'frmPrePago.cmdTarjeta_Click (RsRecord!tCodigoTarjeta)
                                                                Dim xval As Integer
                                                                Dim xnomtar As String
                                                                Isql = "select top 1 tDetallado, nboton from TTARJETACREDITO WHERE tCodigoTarjeta = '" + Trim(RsRecord!tCodigoTarjeta) + "'"
                                                                Set RsCodigoTar = Lib.OpenRecordset(Isql, Cn)
                                                                
                                                                
                                                                xval = CInt(RsCodigoTar!nBoton)
                                                           
                                                               If xval = 8 Or xval = 0 Then
                                                                    frmPrePago.nIndex = 8
                                                               Else
                                                                    frmPrePago.nIndex = xval
                                                               End If
                                                                                                                                                                                             
                                                               frmPrePago.sTipoTarjeta = RsRecord!tCodigoTarjeta
                                                               frmPrePago.txtNT.Caption = RsRecord!tNumero
                                                               frmPrePago.sTitulo = RsCodigoTar!tDetallado
                                                               frmPrePago.txtVT.Caption = RsRecord!tFechaVencimiento
                                                               frmPrePago.PnET = RsRecord!nMonto
                                                               frmPrePago.PnPT = RsRecord!nPropina
                                                               frmPrePago.cmdOpcion_Click (4)
                                                           End If
                                                       Case "04"
                                                               Call frmPrePago.CargarDocumento(RsRecord!tOtroTipoPago, RsRecord!tBanco, RsRecord!nMonto, RsRecord!tNumero)
                                                   End Select
                                               Else 'Dolares
                                                   If Trim(RsRecord!tTarjeta) = "Efectivo" Then
                                                       frmPrePago.PnEE = RsRecord!nMonto / RsRecord!nTipoCambio
                                                   End If
                                               End If
                                               RsRecord.MoveNext
                                           Loop
                                           frmPrePago.Show vbModal
        
                                           If wEnter Then
                                               cmdOpcion_Click (4)
                                               Cn.Execute "Delete from Dprepago where tdocumento='" & sPedido & "' or tcodigopedido='" & sPedido & "'"
                                           End If
                                           
                                           '18112020
                                           frmVenta.Caption = " Caja : " & sCaja
                                           
                                   ElseIf xTiporecepcion = 2 Or xTiporecepcion = 3 Then   'WEB AP---------------------
                                            Call frmPrePago.Form_Load
                                            Do While Not RsRecord.EOF
                                                    Select Case RsRecord!tipoPago
                                                        Case "01"
                                                                frmPrePago.PnEN = RsRecord!Monto
                                                        Case "02"
                                                            If Not Trim(RsRecord!tarjeta) = "" Then
                                                                'frmPrePago.cmdTarjeta_Click (RsRecord!tarjeta)
                                                                frmPrePago.nIndex = RsRecord!nBoton
                                                                frmPrePago.sTipoTarjeta = RsRecord!tarjeta
                                                                frmPrePago.sTitulo = RsRecord!DesTarjeta
                                                                frmPrePago.txtNT.Caption = RsRecord!Numero
                                                                frmPrePago.nPT = RsRecord!nPropina ' GCAA 25012023
                                                                frmPrePago.PnET = RsRecord!Monto
                                                                frmPrePago.cmdOpcion_Click (4)
                                                            End If
                                                      End Select
                            
                                                RsRecord.MoveNext
                                            Loop
                                            frmPrePago.Show vbModal
                            
                                            If wEnter Then
                                                cmdOpcion_Click (4)
                                                Cn.Execute "Delete from Dprepago where tdocumento='" & sPedido & "' or tcodigopedido='" & sPedido & "'"
                                            End If
                                   End If
                           
                           Else 'PAGO
                                   If lPagoRapidoPV = False Then
                                      sFormulario = ""
                                      sTipo = ""
                                      frmPago.Show vbModal
                                   Else
                                      lModuloPago = "PuntoVenta"
                                      frmPagoRapido.Show vbModal
                                   End If
                                   
                                    If wEnter = False Then
                                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                        
                                        xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                        Exit Sub
                                    End If
                           End If
                   
                   Else 'Sin Central Delivery
                           If lCanalDelivery Then
                            If lIntPagoEfectivo = False Then
                                frmPrePago.Show vbModal
                               If wEnter Then
                                   cmdOpcion_Click (4)
                               End If
                            Else
                                sFormulario = ""
                                sTipo = ""
                                frmPago.Show vbModal
                                
                                If wEnter = False Then
                                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                    Exit Sub
                                End If
                                
                            End If

                           Else
                               If lPagoRapidoPV = False Then
                                  sFormulario = ""
                                  sTipo = ""
                                  frmPago.Show vbModal
                               Else
                                  lModuloPago = "PuntoVenta"
                                  frmPagoRapido.Show vbModal
                               End If
                               
                                If wEnter = False Then
                                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                    
                                    'JSJR 21/01/2025
                                    RsDetalle.MoveFirst
                                    
                                    Exit Sub
                                End If
                                
                           End If
                   End If
               Else
                   'Libera la Mesa por cuenta Interna
                   Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica = 1 where tCodigoPedido ='" & sPedido & "'"
                   
                   'Juntar Mesa
                   Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
               End If
        End If
        '-----------------------
        If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
             If Not FacturarTCPIP(2, sDocumento, 0) Then
                 Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                 Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 0  where tCodigoPedido ='" & sPedido & "'"
                 Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 
                 Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                 xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                 Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"

                GoTo fin
             End If
        End If
        '------------------------
       'Actualiza Base de Datos Detalle del Pedido
        If sPrefijo = "0" Then
            Cn.Execute "Update DPEDIDO set tFacturado = 'C', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tfacturado))=0)"
            Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica=1 where tCodigoPedido = '" & sPedido & "'"
            Cn.Execute "Update MDOCUMENTO set tEstadoDocumento ='02',lreplica=1  where tDocumento = '" & sDocumento & "'"
        Else
            If lPagoAntesImpresion Then
                Cn.Execute "Update DPEDIDO set tFacturado = 'P', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
            Else
                Cn.Execute "Update DPEDIDO set tFacturado = 'F', tDocumento = '" & sDocumento & "' where tCodigoPedido ='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
            End If

            Cn.Execute "Update DPEDIDO set tFacturado = 'C' where tDocumento ='" & sDocumento & "' and len(ltrim(tCortesia)) = 4 "
            Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica=1 where tCodigoPedido ='" & sPedido & "'"
        End If
        
        'PARA NO FISCALES
        Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                      
'        If lIntPagoEfectivo And sTipoDocumento <> "00" Then
'            If (MsgBox("¿Desea emitir Codigo CIP PagoEfectivo en el Documento?", vbYesNo) = vbYes) Then
'               If (EjecutaMotorIntegraciones(sDocumento, sUsuario, 7, 1) = True) Then
'
'               Else
'                    MsgBox "No se pudo generar el codigo CIP PagoEfectivo en la operacion!!!", vbInformation
'                    Set ImageMercadoPago = Nothing
'               End If
'            End If
'        End If

        If lInfhotel Then
           xSuma = Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tDocumento ='" & sDocumento & "' and tCodigoPedido='" & sPedido & "'", Cn)
           If RsTipoDocumento!TTipoEmision = "00" Then
              PrecioTotal = 0
              EstadoC = "01"
           Else
              PrecioTotal = xSuma
              EstadoC = "02"
           End If
   
           If Len(Trim(RsCabecera!tComanda)) = 0 Or IsNull(RsCabecera!tComanda) Then
              sComandaInfhotel = Calcular("select left(MAX(tComanda),8) as Codigo from MCOMANDA where tPuntoVenta='" & sPuntoVenta & "'", CnInfhotel)
              sComandaInfhotel = Lib.Correlativo(sComandaInfhotel, 8)
              CnInfhotel.Execute "Update TPUNTOVENTA Set nUltimoComanda = '" & sComandaInfhotel & "' where tPuntoVenta='" & sPuntoVenta & "'"
              sComandaInfhotel = sComandaInfhotel & "-" & UCase(Mid(rsPuntoVenta!descripcion, 1, 3))
              rsPuntoVenta.Requery
              rsPuntoVenta.MoveFirst
              rsPuntoVenta.Find "Codigo='" & sPuntoVenta & "'"
                                          
              Isql = "Insert into MCOMANDA " & _
                     "(tComanda, tPuntoVenta, tHotel, nMovimiento, fFecha, hHora, nTotal, tEstado, " & _
                     "tEmitido, tAsignacion, tCodigoReserva, tNumeroHabitacion, tCodigoFuncionario, " & _
                     "tCaja, tDocumento, tUsuario, nTCambio, tCodigoCompania, tCliente, tMoneda, fFechaE, hHoraE, tUsuarioE,tnotapedido) " & _
                     "values('" & sComandaInfhotel & "', '" & sPuntoVenta & "', '" & sHotel & "', 1,  getdate(), getdate(), " & PrecioTotal & ", '" & EstadoC & "', " & _
                     "1, '" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', '', '', '" & IIf(RsTipoDocumento!TTipoEmision = "00", Mid(sCortesia, 3, 2), "") & "', " & _
                     "'" & sCajaInfhotel & "', '" & IIf(pais = "002", Mid(sDocumento, 1, 1) + Mid(sDocumento, 3), sDocumento) & "', '" & xUsuario & "', " & nTC & ", '', '" & sPasajero & "', '01', getdate(), getdate(), '" & xUsuario & "', '" & sPedido & "')"
              CnInfhotel.Execute Isql
           Else
              sComandaInfhotel = RsCabecera!tComanda
              CnInfhotel.Execute "update MCOMANDA set TASIGNACION='" & IIf(RsTipoDocumento!TTipoEmision = "00", "05", "01") & "', TCODIGORESERVA='', TNUMEROHABITACION='', TCLIENTE='', nTotal= " & xSuma & ", tEstado='" & EstadoC & "', tDocumento='" & IIf(pais = "002", Mid(sDocumento, 1, 1) + Mid(sDocumento, 3), sDocumento) & "' " & _
                                 "where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & sPuntoVenta & "'"
           End If
           
           'Detalle
           Dim nMovimiento As Integer
           CnInfhotel.Execute "delete from DCOMANDA where tComanda ='" & RsCabecera!tComanda & "' and tPuntoVenta='" & sPuntoVenta & "'"
           nMovimiento = Calcular("select max(nmovimiento) as codigo from dcomanda where tcomanda='" & sComandaInfhotel & "'", CnInfhotel) + 1
            If RsTipoDocumento!TTipoEmision = "00" Then
                       Isql = "Insert into DCOMANDA " & _
                              "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                              "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                              "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                              "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & ",  '" & sPedido & "' as  tNotaPedido, tInfhotel as tCodigoItem, " & _
                              "0, nCantidad, 0, 0, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                              "'' as tCuenta, '" & sCajaInfhotel & "' as tCaja, '" & IIf(pais = "002", Mid(sDocumento, 1, 1) + Mid(sDocumento, 3), sDocumento) & "' as tDocumento, '" & IIf(sDescrip = "Reserva", "03", "02") & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                              "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                              "where tDocumento='" & sDocumento & "' and tCodigoPedido ='" & sPedido & "'"
                       CnInfhotel.Execute Isql
            Else
                       Isql = "Insert into DCOMANDA " & _
                              "(tComanda, tPuntoVenta, tHotel, tItem, nMovimiento, tNotaPedido, tCodigoItem, " & _
                              "nPrecioUnitario, nCantidad, nTotal, nPrecioCos, tCodigoReserva, tNumeroHabitacion, " & _
                              "tCuenta, tCaja, tDocumento, tAsignado, tUsuario, fFecha, hHora) " & _
                              "select '" & sComandaInfhotel & "' as tComanda, '" & sPuntoVenta & "' as tPuntoVenta, '" & sHotel & "' as tHotel, tItem , " & nMovimiento & ",  '" & sPedido & "' as  tNotaPedido, tInfhotel as tCodigoItem, " & _
                              "T1.nPrecioVenta as nPrecioUnitario, nCantidad, nVenta as nTotal, T1.nInsumo+T1.nGasto+T1.nManoObra as nPrecioCos, '" & sReserva & "' as tCodigoReserva, '" & sHabitacion & "' as tNumeroHabitacion, " & _
                              "'' as tCuenta, '" & sCajaInfhotel & "' as tCaja, '" & IIf(pais = "002", Mid(sDocumento, 1, 1) + Mid(sDocumento, 3), sDocumento) & "' as tDocumento, '" & IIf(sDescrip = "Reserva", "03", "02") & "' as tAsignado, '" & xUsuario & "' as  tUsuario, getdate() as fFecha, getdate() as hHoraMovimiento " & _
                              "FROM OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.DPEDIDO) T1 INNER JOIN OPENROWSET('SQLOLEDB','" & sRuta & "';'" & sUserName & "';'" & sUserPassword & "', " & sMDB & ".dbo.TPRODUCTO) T2 ON T1.tCodigoProducto = T2.tCodigoProducto " & _
                              "where tDocumento='" & sDocumento & "' and tCodigoPedido ='" & sPedido & "'"
                       CnInfhotel.Execute Isql
            End If
            Cn.Execute "update MPEDIDO set tComanda = '" & sComandaInfhotel & "'  where tCodigoPedido='" & sPedido & "'"
        End If
    
        '============= nov 2010
        Dim imprimeDolaDocumentos As String
        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
        If imprimeDolaDocumentos = "Verdadero" Then
            lDocumEquivaPrecuenta = True
        Else
            lDocumEquivaPrecuenta = False
        End If
        
            If wConsumo = False And lDescripcionAlternativa = True Then
             If validaImpresionAlternativa(sDocumento) = False Then
                    If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                          lImprimeAlternativa = True
                    End If
              End If
            End If
        '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
        ' ELDCQ 15/11/2017
        'Imprime Documentos
         If lImprimeAlternativa = False Then
            If lDocumentoAgrupado Then
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
'                'FACTURACION_E_PERU
'                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',2"
            Else
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
'                'FACTURACION_E_PERU
'                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
            End If
        Else
            If lDocumentoAgrupado Then
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
'                'FACTURACION_E_PERU
'                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',6"
            Else
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
'               'FACTURACION_E_PERU
'                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
            End If
        End If
       '------------------------------------------------------------------------------------------------------------
        'FACTURACION_E_PERU
        'Set RsImpDocumentoE = Lib.OpenRecordset(IsqlFact, Cn)
        xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
        xImpresioDE = Mid(sDocumento, 1, 1)
        '---------------------------------------
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        Set rstFuente = New ADODB.Recordset
        imageCab.Picture = Nothing
        imagepIE.Picture = Nothing
        Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
        imageCab.DataField = "foto"
        Set imageCab.DataSource = rstFuente
        imagepIE.DataField = "fotoPie"
        Set imagepIE.DataSource = rstFuente

       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
          MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
       Else
        'SUNAT
        numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
        codigoImpresora = sImp
        'SUNAT
        Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "

               'FACTURACION_E_PERU
               If pais = "000" Then
                   If lFacturacionE Then
                   
                       If lFEOfisis Then 'OFISIS
                                        '----CABECERA
                                        Set oComandoCabeceraOfisis1 = New clsComando
                                        If Not oComandoCabeceraOfisis1.CreateCmdSp("USP_FactDocumentoOfisis", Cn) Then
                                             Set oComandoCabeceraOfisis1 = Nothing
                                             Exit Sub
                                        End If
                                        oComandoCabeceraOfisis1.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
            
                                        If Not oComandoCabeceraOfisis1.GetParamOK Then
                                             Set oComandoCabeceraOfisis1 = Nothing
                                             Exit Sub
                                        End If
                                        If Not oComandoCabeceraOfisis1.ExecSP Then
                                             Set oComandoCabeceraOfisis1 = Nothing
                                             Exit Sub
                                        End If
                                        
                                        '----FIRMA DOCUMENTO OFISIS
                                        If RsTipoDocumento!lDocumentoElectronicoOfisis Then
                                            Set oComandoFirmaDocumentoOfisis1 = New clsComando
                                            If Not oComandoFirmaDocumentoOfisis1.CreateCmdSp("USP_FactFirmaDocumentoOfisis", Cn) Then
                                                 Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                 Exit Sub
                                            End If
                                            oComandoFirmaDocumentoOfisis1.CreateParameter "@CodigoDocumento", adVarChar, adParamInput, 20, sDocumento
                
                                            If Not oComandoFirmaDocumentoOfisis1.GetParamOK Then
                                                 Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                 Exit Sub
                                            End If
                                            If Not oComandoFirmaDocumentoOfisis1.ExecSP Then
                                                 Set oComandoFirmaDocumentoOfisis1 = Nothing
                                                 Exit Sub
                                            End If
                                            
                                            'VALIDAR RESPUESTA CODIGO DE BARRA
                                            fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                            Sleep 2500
                                            If lImpresionCodigoBarras Then
                                                imageHash.DataField = "foto"
                                                Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                                Set imageHash.DataSource = RsCodigoHash

                                            ElseIf lQRFE Then
                                                Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                            Else

                                                Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                                If RscadenaCodigoHash.RecordCount > 0 Then
                                                    cadenaCodigoHash = RscadenaCodigoHash!codigo
                                                End If
                                                'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                            End If
                                            
                                        End If
                                                                
                       ElseIf lFESpring Then

                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then

                                frmMensajeFeSpring.sMesa = sMesa
                                frmMensajeFeSpring.sDocumento = sDocumento
                                frmMensajeFeSpring.sTipoDocumento = sTipoDocumento
                                frmMensajeFeSpring.sCaja = sCaja
                                frmMensajeFeSpring.sPedido = sPedido
                                frmMensajeFeSpring.oVenta = 1 ' 1: "Formulario Venta"

                                frmMensajeFeSpring.Show vbModal

                                If frmMensajeFeSpring.lEnvio = False Then
                                    Exit Sub
                                End If

                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                        If frmMensajeFeSpring.lQrInf Then
                                            Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                        Else
                                            Set imageHash.Picture = LoadPicture(ImagenFeSpring(lQRFE, sDocumento))
                                        End If
                                     Else
                                         If lImpresionCodigoBarras Then
                                             
                                         Else
                                             cadenaCodigoHash = ImagenFeSpring(lQRFE, sDocumento)
                                         End If
                                     End If
                                 End If
                             End If
                                                                
                       ElseIf lFECarbajal Then
                            Label2.Caption = "   Proceso de envio de documento a InfoFact......."
                            lblPaso1.Caption = "Enviando información de documento a InfoFact."
                            lblPaso2.Caption = "Obteniendo codigo " & IIf(lQRFE, "QR", IIf(lImpresionCodigoBarras, "de barras", " hash")) & " almacenado."
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                               ' Dim sImporteLetra As String
                                sImporteLetra = NumeroCadena(str(RsImpresion!nVenta)) + " " + sMonedaN
                                FrameFeSpring.Visible = True
                                Sleep 1000
                                If Not INSERTAFE_CARVAJAL(sDocumento, sImporteLetra, 0, 0) Then '----CABECERA
                                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                    imgProceso(2).Visible = True
                                    imgProceso(3).Visible = True
                                    Sleep 1000
                                    FrameFeSpring.Visible = False
                                    Exit Sub
                                 End If
                                 imgProceso(0).Visible = True
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                         Else
                                             cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                         End If
                                     End If
                                 End If
                                 imgProceso(1).Visible = True
                                 Sleep 1000
                                 FrameFeSpring.Visible = False
                            End If
                            
                       ElseIf lFEpape Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
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
                            End If
                       ElseIf lFEBiz Then
                           'GCAA 02122022 - OFFLINE
                           
                           If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If Not INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Date) Then '----CABECERA
                                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                    xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                                    If TipoFacturacion = "5" Then
                                        msgError = "Facturacion USIL"
                                    Else
                                        msgError = "Bizlink"
                                    End If
                                    'MsgBox "Se Produjo un Problema al enviar el Documento" & vbNewLine & "Favor de Refrescar los valores e intentarlo nuevamente!!!", vbInformation, sMensaje
                                    GoTo fin
                                 End If
                                 
                                 
                                 Sleep 1000
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                         Else
                                             cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                         End If
                                     End If
                                 End If
                            End If
                            
                            
                            
                       ElseIf lFEGesa Then
                        If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                            If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                'Exit Sub
                            End If
                            If lQRFE Then
                                Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                            Else
                                If lImpresionCodigoBarras Then
                                    Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                Else
                                    cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                End If
                            End If
                            
                        End If
                       
                       Else 'INFOFACT
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                 If Not INSERTAFE(sDocumento, "", 1, "") Then '----CABECERA
                                    MsgBox "No se pudo enviar el documento a facturacion electronica", vbInformation
                                     'Exit Sub
                                 End If
'                                 If RsImpresion!Ruc <> "" Then
'                                     If Not INSERTAFE(sDocumento, "", 2, RsImpresion!Ruc) Then '----CLIENTE
'                                         Exit Sub
'                                     End If
'                                 End If
                                 'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                 fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                                 If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                         Else
                                             cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                         End If
                                     End If
                                 End If
                             End If
                       End If
                   End If
            End If
            '---------------------------------------
            'Configura la Impresora
            Imprimir (sImp)
            Printer.FontName = sFont
            Printer.FontBold = False
            
            'GCAA AQUI FE ARGENTINA 14032022
            If pais = "003" And lFETusFacturasAPP = True And (Left(sDocumento, 1) = "A" Or Left(sDocumento, 1) = "B") Then
                 'If lPagoAntesImpresion Then
                    If Not (INSERTA_FE_INFOREST_ARGE(sDocumento, 1, DateTime.Date)) Then
                        Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                        Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                        xUltimoCorrelativo = Calcular("select isnull(MAX(tDocumento),'000000000') as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                        xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                        Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                        Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                        Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                        msgError = "Error de conexion con el servicio TUSFACTURASAPP, " & vbNewLine & " revisar el motor de facturacion, conexion de impresora"
                        GoTo fin
                    Else
                        'ELDCQ 27/04/2022
                         Dim Termico As String
                         Termico = Calcular("select isnull(tImprTermica,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                         If Termico <> "" Then
                            Set imageHash.Picture = LoadPicture(CrearImagenQR(Termico))
                            ImprimeFacturaElectronicaARGE RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageMercadoPago
                        Else
                           ' MsgBox "Por el momento los servicios de AFIP estan caidos o hubo un percance en la facturacion el documento intentara enviarse en uno minutos a mas tardar 1 dia, favor indicar al cliente que revise su correo electronico para la recepcion de la factura: " & sDocumento
                           Dim MENSAJERROR As String
                                MENSAJERROR = Calcular("select isnull(error,'') as codigo from DOCUMENTO_ARGENTINA where documento='" & sDocumento & "'", Cn)
                                MsgBox MENSAJERROR, vbInformation
                        End If
                    End If
                'End If
            Else
            
            
            
            If pais = "003" And Calcular("select isnull(tpuertoserial,'') as codigo from timpresora where tcaja='" & sCaja & "' and timpresora='" & sImp & "'", Cn) <> "" Then ' Pais Argentina
                Dim RespDev As Integer
'                If lFETusFacturasAPP Then
'                      If Not (INSERTA_FE_INFOREST_ARGE(sDocumento, 1, DateTime.Date)) Then
'                        RespDev = 0
'                      Else
'                        RespDev = 2
'                      End If
'                Else
                    RespDev = ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 2, RsImpresion)
                'End If
                
                If RespDev = 0 Then
                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    xUltimoCorrelativo = Calcular("select isnull(MAX(tDocumento),'000000000') as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                    msgError = "Error de conexion con la impresora fiscal, " & vbNewLine & " revisar fuente de papel, conexion de impresora"
                    GoTo fin
                ElseIf RespDev = 1 Then
                    Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuario & "', fRegistroAnulado = getdate() , tObservacion ='Anulacion por error en impresora fiscal', lreplica=1  where tDocumento ='" & sDocumento & "'"
'                    xUltimoCorrelativo = Calcular("select isnull(MAX(tDocumento),'000000000') as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
'                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
'                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Cn.Execute "Update DPEDIDO set tFacturado = NULL, tDocumento = NULL where tCodigoPedido ='" & sPedido & "' "
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '01', lReplica = 1 where tCodigoPedido = '" & sPedido & "'"
                    Cn.Execute "UPDATE DPEDIDO SET lregistroventa = NULL where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
                    msgError = "Error en el proceso de impresion, " & vbNewLine & "Corte de energia u otro, revisar fuente de papel, conexion de impresora"
                    GoTo fin
                End If

            Else

                If wConsumo Then
                    If RsTipoDocumento!tFormulario = "01" Then
                         If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                             'FACTURACION ELECTRONICA
                             If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                  ImprimeFacturaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion
                             Else
                                  ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                             End If
                            
                             NFactura = sCorrela
                             lblFactura.Caption = NFactura
                         Else
                             'FACTURACION ELECTRONICA
                             If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                 ImprimeBoletaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion
                             Else
                                 ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                             End If
                         End If
                       
                    Else
                   
                        If lFacturacionE And RsTipoDocumento!lFacturacionElectronica And lFEOfisis = False And lFESpring = False And lFEpape = False And lFECarbajal = False Then
                              If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                  ImprimeFormatoAConsumo
                                  Kill App.path & "\fact.bmp"
                              Else
                                  ImprimeFormatoAConsumo
                              End If
                        ElseIf lFacturacionE And lFECarbajal Then
                              ImprimeFormatoAConsumo
                              
                              Set fso1 = CreateObject("Scripting.FileSystemObject")
                              If fso1.FileExists(App.path & "\fact.bmp") Then
                                  Kill App.path & "\fact.bmp"
                              End If
                        ElseIf lFacturacionE = False Then
                              If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                                 ImprimeFacturaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                                 NFactura = sCorrela
                                 lblFactura.Caption = NFactura
                              Else
                                 ImprimeBoletaConsumoN RsImpresion, sDetalleConsumo, sEmpresa
                              End If
                        End If
                   End If
                Else
                   If RsTipoDocumento!tFormulario = "01" Then
                      If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                              'FACTURACION ELECTRONICA
                              If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                 If pais <> "005" Then
                                    ImprimeFacturaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageMercadoPago
                                 Else
                                    ImprimeFacturaFiscal RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageMercadoPago
                                 End If
                              Else
                                 ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery, ImageIntegracion
                              End If
                            NFactura = sCorrela
                            lblFactura.Caption = NFactura
                              
                      ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                              If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                 ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                              End If
                      Else
                
                              'FACTURACION ELECTRONICA
                              If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                                 If pais <> "005" Then
                                    ImprimeBoletaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageMercadoPago
                                 Else
                                    ImprimeFacturaSimple RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery, ImageIntegracion, ImageMercadoPago
                                 End If
                              Else
                                 ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery, ImageIntegracion
                              End If
                      End If
                   Else
                      If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                         ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                         NFactura = sCorrela
                         lblFactura.Caption = NFactura
                      ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                         If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                            ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                         End If
                      Else
                         ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                      End If
                      
                   End If
                   
                End If

            End If
            
            
       End If
       End If 'FFFFFFFFFFFF
       LimpiaRs
       
       If pais = "002" And lFEEcuador = False Then
          sXML = Trim(LeerIni(App.path + "\INFOREST.INI", "Configuracion", "RutaXML", "."))
          GeneraFacturaElectronica sXML, sDocumento
       End If
       
       If lFEEcuador Then
        If INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Now) = False Then
            MsgBox "No se pudo enviar el documento a Facturacion Electronica!!! Verificar con su area de sistemas!!!"
        End If
       End If
    
       If lPagoAntesImpresion = False Then
              'Cancelacion del Documento
              If RsTipoDocumento!TTipoEmision <> "00" Then
                   Cn.Execute "Update TMESA set tEstadoMesa = '05' where tCodigoMesa ='" & sMesa & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica = 1  where tCodigoPedido ='" & sPedido & "'"
                   xTipo = ""
              
                   If (CD2 = True Or WEBAP = True) And xTiporecepcion <> 0 Then
                                                              
                           If xTiporecepcion = 1 Then
                               Isql = "USP_CD_SELECCIONAR_DPREPAGO_X_PEDIDO '" & tCodigoPedidoCD & "',''"
                               Set RsRecord = Lib.OpenRecordset(Isql, CnCD)
                           ElseIf xTiporecepcion = 2 Or xTiporecepcion = 3 Then
                                Isql = " exec usp_WS_Inforest '" & sPedido & "' , 1 " '
                                Set RsRecord = Lib.OpenRecordset(Isql, Cn)
                           End If
                             
                            ' GCAA 26012019
                           If xTiporecepcion <> 0 Then  'PrePago 'sVar1 = "02"
                                   If xTiporecepcion = 1 Then
                                           Isql = "Select tCodigoPedidoCD From MPEDIDO WHERE tCodigoPedido = '" + sPedido + "'"
                                           Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
                                           If Not RsCodigoPCD.EOF Then
                                             If Not IsNull(RsCodigoPCD!tCodigoPedidoCD) Then
                                                 Call ModifcarEstadoDeliveryCabecera(RsCodigoPCD!tCodigoPedidoCD, "3", txtMotorizado.Caption)
                                             End If
                                           End If
        
                                        'GCAA 16112020
        
                                        
                                           Call frmPrePago.Form_Load
                                           Do While Not RsRecord.EOF
                                               If RsRecord!tMONEDA = "01" Then 'Soles
                                                   Select Case RsRecord!tTipoPago 'pp
                                                       Case "01"
                                                           If Trim(RsRecord!tTarjeta) = "Efectivo" Then
                                                               frmPrePago.PnEN = RsRecord!nMonto
                                                           End If
                                                       Case "02"
                                                           If Not Trim(RsRecord!tCodigoTarjeta) = "" Then
                                                           'GCAA 30072020 - 04092020
                                                            Isql = "select top 1 tDetallado , nboton  from TTARJETACREDITO WHERE tCodigoTarjeta = '" + Trim(RsRecord!tCodigoTarjeta) + "'"
                                                            Set RsCodigoTar = Lib.OpenRecordset(Isql, Cn)
                                                             
                                                             Dim valx As Integer
                                                              valx = CInt(RsCodigoTar!nBoton)
                                                           
                                                               If valx = 8 Or valx = 0 Then
                                                                    frmPrePago.nIndex = 8
                                                               Else
                                                                    frmPrePago.nIndex = valx
                                                               End If
                                                               
                                                               frmPrePago.sTipoTarjeta = RsRecord!tCodigoTarjeta
                                                               frmPrePago.txtNT.Caption = RsRecord!tNumero
                                                               frmPrePago.txtVT.Caption = RsRecord!tFechaVencimiento
                                                               frmPrePago.sTitulo = RsCodigoTar!tDetallado
                                                               frmPrePago.PnET = RsRecord!nMonto
                                                               frmPrePago.PnPT = RsRecord!nPropina
                                                               frmPrePago.cmdOpcion_Click (4)
                                                           End If
                                                       Case "04"
                                                           Call frmPrePago.CargarDocumento(RsRecord!tOtroTipoPago, RsRecord!tBanco, RsRecord!nMonto, RsRecord!tNumero)
                                                   End Select
                                               Else 'Dolares
                                                   If Trim(RsRecord!tTarjeta) = "Efectivo" Then
                                                       frmPrePago.PnEE = RsRecord!nMonto / RsRecord!nTipoCambio
                                                   End If
                                               End If
                                               RsRecord.MoveNext
                                           Loop
                                           frmPrePago.Show vbModal
        
                                           If wEnter Then
                                               cmdOpcion_Click (4)
                                               Cn.Execute "Delete from Dprepago where tdocumento='" & sPedido & "' or tcodigopedido='" & sPedido & "'"
                                           End If
                                           
                                        '18112020
                                        frmVenta.Caption = " Caja : " & sCaja
                                           
                                           
                                   ElseIf xTiporecepcion = 2 Or xTiporecepcion = 3 Then     'WEB AP---------------------
                                        Call frmPrePago.Form_Load
                                        Do While Not RsRecord.EOF
                                                Select Case RsRecord!tipoPago
                                                    Case "01"
                                                            frmPrePago.PnEN = RsRecord!Monto
                                                    Case "02"
                                                        If Not Trim(RsRecord!tarjeta) = "" Then
'                                                            frmPrePago.cmdTarjeta_Click (RsRecord!tarjeta)
'                                                            frmPrePago.txtNT.Caption = RsRecord!Numero
'                                                            frmPrePago.PnET = RsRecord!Monto
                                                            frmPrePago.nIndex = RsRecord!nBoton
                                                            frmPrePago.sTipoTarjeta = RsRecord!tarjeta
                                                            frmPrePago.sTitulo = RsRecord!DesTarjeta
                                                            frmPrePago.txtNT.Caption = RsRecord!Numero
                                                            frmPrePago.PnET = RsRecord!Monto
                                                            frmPrePago.cmdOpcion_Click (4)
                                                        End If
                                                  End Select
                        
                                            RsRecord.MoveNext
                                        Loop
                                        frmPrePago.Show vbModal
                        
                                        If wEnter Then
                                            cmdOpcion_Click (4)
                                            Cn.Execute "Delete from Dprepago where tdocumento='" & sPedido & "' or tcodigopedido='" & sPedido & "'"
                                        End If
                                   End If
                           
                           Else 'PAGO
                                   If lPagoRapidoPV = False Then
                                      sFormulario = ""
                                      sTipo = ""
                                      frmPago.Show vbModal
                                   Else
                                      lModuloPago = "PuntoVenta"
                                      frmPagoRapido.Show vbModal
                                   End If
                           End If
                   
                   Else 'Sin Central Delivery
                           If lCanalDelivery Then
                               frmPrePago.Show vbModal
                               If wEnter Then
                                   cmdOpcion_Click (4)
                               End If
                           Else
                               If lPagoRapidoPV = False Then
                                  sFormulario = ""
                                  sTipo = ""
                                  frmPago.Show vbModal
                                  '09102020 gcaa
                               Else
                                  lModuloPago = "PuntoVenta"
                                  frmPagoRapido.Show vbModal
                               End If
                           End If
                   End If
               Else
                   'Libera la Mesa por cuenta Interna
                   Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
                   Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica = 1 where tCodigoPedido ='" & sPedido & "'"
                   'Juntar Mesa
                   Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
               End If
       Else
       
                'Libera la Mesa por cuenta Interna
                Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa ='" & sMesa & "'"
                Cn.Execute "Update MPEDIDO set tEstadoPedido = '02', lReplica = 1 where tCodigoPedido ='" & sPedido & "'"
                'Juntar Mesa
                Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
       End If

        '-----------------------
        If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
             If Not FacturarTCPIP(3, sDocumento, 0) Then
                MsgBox ("La confirmacion ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
             End If
        End If
        '------------------------
    End If

    If sCliente <> "" Then
       Cn.Execute "Update TDELIVERY set tCodigoCliente = '" & sClienteFactura & "' where tCodigoDelivery ='" & sTemp & "'"
    End If
    
    'INICIO GCAA 06072022 - CREDITO CORPORATIVO
    Dim xSaberCuentasxCobrar As String
             
    If (lActivaCreditoCoorporativo) Then
'       xSaberCuentasxCobrar = Calcular("Select tEstadoDocumento as codigo from MDOCUMENTO Where tdocumento='" & sDocumento & "'", Cn)
'
'       If xSaberCuentasxCobrar = "03" Then
'          If Not INSERTAFE_CREDITO_CORPO(sDocumento) Then '----CABECERA
'          End If
'       End If
          If Not INSERTAFE_CREDITO_CORPO(sDocumento) Then '----CABECERA
          End If
    End If
        'FIN GCAA 06072022 - CREDITO CORPORATIVO
    
    If lImpComandaBarra And Not lPrinter Then
        cmdOpcion_Click (3)
'    Else
'        If lIntPagoEfectivo Then
'            cmdOpcion_Click (3)
'        End If
    End If

    sTemp = ""
    RsCabecera.Requery
    RsDetalle.Requery
    RsCombo.Requery
    If RsCabecera.EOF Then
       LimpiaCabecera
    End If
    sVar1 = sUsuario ' FERM 29/10/2024
    Screen.MousePointer = vbDefault
    HabilitaTimerColor (True)
    Exit Sub
fin:
    sVar1 = sUsuario ' FERM 29/10/2024
    FrameFeSpring.Visible = False
    RsCabecera.Requery
    RsDetalle.Requery
    RsCombo.Requery
    If RsCabecera.EOF Then
       LimpiaCabecera
    End If
    Screen.MousePointer = vbDefault
    HabilitaTimerColor (True)
    Call Log_Inforest("PUNTO VENTA", "EMISION DE DOCUMENTO PUNTO DE VENTA", sPedido, "", sDocumento, Estado & " " & Error, "", "FALLA AL GENERAR DOCUMENTO EN PUNTO DE VENTA ", sUsuario)
    MsgBox "Error " & msgError & " : " & Error & vbNewLine & "- Proceso: Emision de Documentos Punto Venta", vbCritical, sMensaje
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Me.bocina.Settings.autoStart = True
    Me.bocina.uiMode = "none"

  ' origenes de venta
  Me.fraOrigenVentas.Visible = False
  '----------------------------------
  ValidaDesc = True
  If lBloqueaPrecuenta Then
    cmdOpcion(4).Enabled = False
  Else
    cmdOpcion(4).Enabled = True
  End If
      
  
  Label1(8).Caption = Mesero & ":"
   frmVenta.Caption = " Caja : " & sCaja
    'GCAA 16112020
   xSinPrepago = Calcular("select isnull(lPagoDirectoCD,0)AS codigo from tcaja Where tcaja='" & sCaja & "'", Cn)
   
   
   
   muestra = Me.Caption
   sComanda = ""
   sMozoD = ""
   sHabitacion = ""
   sReserva = ""
   sPuntoVenta = ""
   sCompania = ""
   sContacto = ""
   sPasajero = ""
   sUsuarioAutoriza = ""
   
   cmdCabecera(6).Caption = IIf(sBoton1 <> "", sBoton1, "n/d")
   cmdCabecera(7).Caption = IIf(sBoton2 <> "", sBoton2, "n/d")
   cmdCabecera(8).Caption = IIf(sBoton3 <> "", sBoton3, "n/d")
   cmdCabecera(17).Caption = IIf(sBoton4 <> "", sBoton4, "n/d")
   cmdCabecera(18).Caption = IIf(sBoton5 <> "", sBoton5, "n/d")

   cmdCabecera(6).Enabled = IIf(sBoton1 <> "", True, False)
   cmdCabecera(7).Enabled = IIf(sBoton2 <> "", True, False)
   cmdCabecera(8).Enabled = IIf(sBoton3 <> "", True, False)
   cmdCabecera(17).Enabled = IIf(sBoton4 <> "", True, False)
   cmdCabecera(18).Enabled = IIf(sBoton5 <> "", True, False)
   
   If lBloqueoCanal1 Then
    cmdCabecera(6).Enabled = False
   End If
    If lBloqueoCanal2 Then
    cmdCabecera(7).Enabled = False
   End If
    If lBloqueoCanal3 Then
    cmdCabecera(8).Enabled = False
   End If
    If lBloqueoCanal4 Then
    cmdCabecera(17).Enabled = False
   End If
   If lBloqueoCanal5 Then
    cmdCabecera(18).Enabled = False
   End If
   
   nOperadorPropiedad = 0
  
   'Cabecera
   'JL
   If lMCPV Then
      Isql = "SELECT  *, MPEDIDO.tTiporecepcion  from vPedidoCabecera inner join MPEDIDO on vPedidoCabecera.Codigo  = MPEDIDO.tCodigoPedido  " & _
             "where vPedidoCabecera.tEstadoPedido = '01' and vPedidoCabecera.tUsuario = '" & sUsuario & "' order by Codigo"
   Else
      Isql = "SELECT  *, MPEDIDO.tTiporecepcion  from vPedidoCabecera inner join MPEDIDO on vPedidoCabecera.Codigo  = MPEDIDO.tCodigoPedido  " & _
             "where vPedidoCabecera.tEstadoPedido = '01' and vPedidoCabecera.tCaja = '" & sCaja & "' order by Codigo"
   End If
   Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
   
   'Timer
   If lEnvioAutomatico Then
      Isql = "SELECT  *, MPEDIDO.tTiporecepcion  from vPedidoCabecera inner join MPEDIDO on vPedidoCabecera.Codigo  = MPEDIDO.tCodigoPedido  " & _
          "where vPedidoCabecera.tEstadoPedido = '01' and vPedidoCabecera.tCaja = '" & sCaja & "' and FecProg is not null and isnull(vPedidoCabecera.nMinutosAntesEnvio,0) <> 0 order by Codigo"
      Set RsTimer = Lib.OpenRecordset(Isql, Cn)
   End If

   'Detalle
   'JL
   If lMCPV Then
      Isql = "select * from vPedidoDetalle " & _
             "where tEstadoItem ='N' and isnull(tfacturado,'')='' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tusuario='" & sUsuario & "') " & _
             "ORDER BY tCodigoPedido, tItem"
   Else
      Isql = "select * from vPedidoDetalle " & _
             "where tEstadoItem ='N' and isnull(tfacturado,'')='' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tCaja = '" & sCaja & "') " & _
             "ORDER BY tCodigoPedido, tItem"
   End If
   Set RsDetalle = Lib.OpenRecordset(Isql, Cn)
      
   'Grupo
   Isql = "select * from vGrupo where lActivo = 1 order by nBoton"
   Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
   If RsGrupo.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un Grupo creado", vbCritical, sMensaje
      End
   End If
   
   'Sub Grupo
   Isql = "select * from vSubGrupo where lActivo = 1 Order by nBoton"
   Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
   If RsSubgrupo.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un SubGrupo creado", vbCritical, sMensaje
      End
   End If
      
   'Producto
  ' Isql = "select * from vProducto where lActivo = 1 Order by nBoton"
   Set RsProducto = Lib.OpenRecordset("usp_Inforest_ObtieneProductos", Cn)
   
   If RsProducto.RecordCount = 0 Then
      MsgBox "Error: Se necesita al menos un producto creado", vbCritical, sMensaje
      End
   End If
                                                      
   'Mozos
    Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by nBoton"
    Set RsMozo = Lib.OpenRecordset(Isql, Cn)
   'OO
   
   If lMCPV Then
      txtMozo.Caption = sVar1
      fraMozo.Visible = False
   End If
   
   'Motorizado
   Isql = "select * from vMotorizado where lActivo = 1 Order by nBoton"
   Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
   
   'Origen de ventas
   Isql = "select * from vOrigenVenta where Activo = 1 and Visible = 1 and CodOrigenVenta not in (select CodigoOrigenVenta from TCAJAORIGEN_BLOQUEO  where tcaja='" & sCaja & "' ) Order by Boton"
   Set RsOrigenVentas = Lib.OpenRecordset(Isql, Cn)
   
   Isql = "select * from vTipoPedido where Codigo = '02'"
   Set RscanalOrigenVentas = Lib.OpenRecordset(Isql, Cn)
   '-----------------------------------------------------
   'Motivo de Eliminacion
   Isql = "select * from vMotivoEliminacion where lActivo = 1 order by Codigo"
   Set RsMotivoEliminacion = Lib.OpenRecordset(Isql, Cn)
   AsignaComando 38, RsMotivoEliminacion, cmdEliminacion()
            
   'Operador
   ListarOperadoresConFiltro (sProducto) 'OO
        
   'Propiedades
   Dim xSql As String
   If lAlmacen Then
      Dim RsOp As Recordset
      Set RsOp = Lib.OpenRecordset("select Codigo, Descripcion from vOperador where lStockMenos=1", Cn)
      If RsOp.RecordCount > 0 Then
        'CASO CORP.FERNANDEZ 'JSJR UNION - UNION ALL
            If BDLink = "1" Then
                xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                       "nInsumo, nGasto, nManoObra, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union all " & _
                       "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0,0 " & _
                       "FROM " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.DRECETAVENTA t1x INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                       "Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "'"
            Else
                xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TOPERADOR.tOperador as tOperador, TOPERADOR.tDetallado as Operador, nPrecio, tEnlace, " & _
                       "nInsumo, nGasto, nManoObra, ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad  " & _
                       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 union all " & _
                       "select '9999' as Codigo, tDetallado as Descripcion, tCodigoPlato as tProducto, '" & RsOp!codigo & "' as tOperador, '" & RsOp!descripcion & "' as Operador, 0, " & "t1x.tCodigoProducto as tEnlace, nCantidad * nPrecio as nInsumo, 0, 0,0 " & _
                       "FROM " & sAlmacenMDB & ".dbo.DRECETAVENTA t1x INNER JOIN " & sAlmacenMDB & ".dbo.MRECETAVENTA t2x ON " & "t1x.tLocal = " & "t2x.tLocal AND " & "t1x.tRecetaVenta = " & "t2x.tRecetaVenta INNER JOIN " & sAlmacenMDB & ".dbo.TPRODUCTO t3x ON " & "t1x.tCodigoProducto = " & "t3x.tCodigoProducto " & _
                       "Where lNoDescargo = 1 and " & "t1x.tLocal='" & sLocal & "'"
            End If
      Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, tOperador.tDetallado AS Operador , ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
      End If
   Else
         xSql = "select tCodigoPropiedad as Codigo, TPROPIEDAD.tDetallado as Descripcion, tProducto, TPROPIEDAD.tOperador, nPrecio, tEnlace, " & _
                "nInsumo, nGasto, nManoObra, tOperador.tDetallado AS Operador , ISNULL(tpropiedad.lsolicitacantidad,0) lsolicitacantidad " & _
                "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
                "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1"
   End If
   'CASO CORP.FERNANDEZ
   If BDLink = "1" Then
        Set RsPropiedad = Lib.OpenRecordset("usp_Inforest_ObtienePropiedad '" & IIf(lAlmacen = True, 1, 0) & "','" & "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" & "','" & sLocal & "'", Cn)
    Else
        Set RsPropiedad = Lib.OpenRecordset("usp_Inforest_ObtienePropiedad '" & IIf(lAlmacen = True, 1, 0) & "','" & sAlmacenMDB & "','" & sLocal & "'", Cn)
    End If
      
    If lMCPV Then
      'Producto Propiedad
       Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador,isnull(tproductopropiedad.ncantidad,1) ncantidad  " & _
              " FROM    dbo.TOPERADOR INNER JOIN dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ")T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto AND dbo.TPRODUCTOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador AND dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
              " where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tUsuario = '" & sUsuario & "') ORDER BY dbo.TOPERADOR.nBoton "
             
      Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
         
      'Combo Propiedad
       Isql = "SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador, isnull(tcombopropiedad.ncantidad,1) ncantidad  " & _
              " FROM         dbo.TOPERADOR INNER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador RIGHT OUTER JOIN dbo.TCOMBOPROPIEDAD INNER JOIN  (" & xSql & ") T1 ON  dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND                       dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS AND dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TCOMBOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TCOMBOPROPIEDAD.tProducto  " & _
              "where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tUsuario = '" & sUsuario & "') order by toperador.nboton "
             
       Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
   Else
       'Producto Propiedad
       Isql = "SELECT dbo.TPRODUCTOPROPIEDAD.tCodigoPedido, dbo.TPRODUCTOPROPIEDAD.tItem, dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad, dbo.TPRODUCTOPROPIEDAD.tProducto, t1.Descripcion AS Descripcion, Operador,isnull(tproductopropiedad.ncantidad,1) ncantidad  " & _
              " FROM    dbo.TOPERADOR INNER JOIN dbo.TPRODUCTOPROPIEDAD INNER JOIN (" & xSql & ")T1 ON dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TPRODUCTOPROPIEDAD.tProducto = T1.tProducto AND dbo.TPRODUCTOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador AND dbo.TPRODUCTOPROPIEDAD.tCodigoPropiedad = dbo.TPROPIEDAD.tCodigoPropiedad AND dbo.TPRODUCTOPROPIEDAD.tProducto = dbo.TPROPIEDAD.tProducto " & _
              "where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' or (tEstadoPedido in ('02','04') AND fregistro>getdate()-1 )  and tCaja = '" & sCaja & "') ORDER BY dbo.TOPERADOR.nBoton "
       Set RsProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
         
       'Combo Propiedad
       Isql = " SELECT dbo.TCOMBOPROPIEDAD.tCodigoPedido, dbo.TCOMBOPROPIEDAD.tItem, dbo.TCOMBOPROPIEDAD.tItemCombo, T1.Descripcion, T1.Operador, isnull(tcombopropiedad.ncantidad,1) ncantidad  " & _
              " FROM         dbo.TOPERADOR INNER JOIN dbo.TPROPIEDAD ON dbo.TOPERADOR.tOperador = dbo.TPROPIEDAD.tOperador RIGHT OUTER JOIN dbo.TCOMBOPROPIEDAD INNER JOIN  (" & xSql & ") T1 ON  dbo.TCOMBOPROPIEDAD.tCodigoPropiedad = T1.Codigo AND dbo.TCOMBOPROPIEDAD.tProducto = T1.tProducto AND                       dbo.TCOMBOPROPIEDAD.tEnlace = T1.tEnlace ON dbo.TOPERADOR.tOperador = T1.tOperador COLLATE Modern_Spanish_CI_AS AND dbo.TPROPIEDAD.tCodigoPropiedad = dbo.TCOMBOPROPIEDAD.tCodigoPropiedad AND dbo.TPROPIEDAD.tProducto = dbo.TCOMBOPROPIEDAD.tProducto  " & _
              " where tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido ='01' or (tEstadoPedido in ('02','04') AND fregistro>getdate()-1 )  and tCaja = '" & sCaja & "')  order by toperador.nboton"
       Set RsComboPropiedad = Lib.OpenRecordset(Isql, Cn)
   End If
         
   'COMBOS2013
    Set RsProductoCombo = Lib.OpenRecordset("usp_Inforest_ObtieneCombos", Cn)
   
   'Areas
   Set RsArea = Lib.OpenRecordset("select * from vAreaImpresora where tCaja ='" & sCaja & "'", Cn)
                                                                                                           
   'Tipo de Documentos
   If pais = "002" Then 'Ecuador
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 and isnull(tNumeroAutorizacion,'')<>'' And lNotaCredito = 0 And lActivo = 1 UNION Select * From vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lFacturacionElectronica=1 and lActivo =1 order by tTipoEmision", Cn)
   Else
      Set RsTipoDocumento = Lib.OpenRecordset("select * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 order by tTipoEmision", Cn)
   End If

   'Canales de Venta
   Set RsCanalesVenta = Lib.OpenRecordset("select * from vTipoPedido", Cn)
   If Not RsCanalesVenta.RecordCount > 0 Then
      MsgBox "Debe tener configurado al menos un Canal de Venta", vbCritical + vbOKOnly, sMensaje
      Unload Me
 
   End If

   'GCCA 02082021 - REQ REST-N014-2021 - TORI
   If (lNoVisualizarItemCombos) Then
      Isql = "SELECT * from vPedidoCombox " & _
          "WHERE tEstadoPedido='01' and tCaja = '" & sCaja & "' " & _
          "ORDER BY tCodigoPedido, tItemCombo"
          Set RsCombo = Lib.OpenRecordset(Isql, Cn)
   Else
      Isql = "SELECT * from vPedidoCombo " & _
          "WHERE tEstadoPedido='01' and tCaja = '" & sCaja & "' " & _
          "ORDER BY tCodigoPedido, tItemCombo"
          Set RsCombo = Lib.OpenRecordset(Isql, Cn)
   End If

   
            
   'Configuración de Grillas
   Call ConfGrilla(3, grdCabecera, "Pax", 2, "nAdulto", 430, 2, 0, "", _
                                   "Mesa", 2, "Mesa", 1000, 0, 0, "", _
                                   "Observacion", 2, "tObservacion", 1200, 0, 0, "")
   
   If (lActivaTresDecimales) Then
   Call ConfGrilla(10, grdDetalle, "Or", 2, "nOrden", 300, 1, 0, "#0", _
                                  "-", 2, "lCorte", 250, 0, 4, "", _
                                  "Producto", 2, "Producto", 1820, 0, 0, "", _
                                  "Precio", 2, "nPrecioVenta", 700, 1, 0, "###,###,##0.00", _
                                  "Cant.", 2, "nCantidad", 590, 1, 0, "##0.000", _
                                  "SubTotal", 2, "nVenta", 790, 1, 0, "###,###,##0.00", _
                                  "F", 2, "tFacturado", 250, 0, 0, "", _
                                  "E", 2, "lImprime", 250, 0, 4, "", _
                                  "P", 2, "lPropiedad", 250, 0, 4, "", _
                                  "O", 2, "lObservacion", 250, 0, 4, "")
   Else
   Call ConfGrilla(10, grdDetalle, "Or", 2, "nOrden", 300, 1, 0, "#0", _
                                  "-", 2, "lCorte", 250, 0, 4, "", _
                                  "Producto", 2, "Producto", 1820, 0, 0, "", _
                                  "Precio", 2, "nPrecioVenta", 700, 1, 0, "###,###,##0.00", _
                                  "Cant.", 2, "nCantidad", 590, 1, 0, "###0.00", _
                                  "SubTotal", 2, "nVenta", 790, 1, 0, "###,###,##0.00", _
                                  "F", 2, "tFacturado", 250, 0, 0, "", _
                                  "E", 2, "lImprime", 250, 0, 4, "", _
                                  "P", 2, "lPropiedad", 250, 0, 4, "", _
                                  "O", 2, "lObservacion", 250, 0, 4, "")
    End If
                                        
   Call ConfGrilla(7, grdCombo, "Od", 2, "nOrden", 400, 1, 0, "#0", _
                                "-", 2, "lCorte", 250, 0, 4, "", _
                                "Producto", 2, "Producto", 1950, 0, 0, "", _
                                "Cant.", 2, "nCantidad", 650, 1, 0, "#,##0.00", _
                                "E", 2, "lImprime", 250, 0, 4, "", _
                                "P", 2, "lPropiedad", 250, 0, 4, "", _
                                "O", 2, "lObservacion", 250, 0, 4, "")

   Set grdCabecera.DataSource = RsCabecera
   Set grdDetalle.DataSource = RsDetalle
   
   'GCAA 20072021
   
   'RsCombo.Filter = "Automatico = 0"
   Set grdCombo.DataSource = RsCombo
   
      
   'Asignar Grupo
   
   AsignaBoton 5, RsGrupo, cmdGrupo()
   RsGrupo.MoveFirst
   RsGrupo.Find "Codigo='" & sGrupoDefault & "'"
   If RsGrupo.EOF Then
      MsgBox "Error: Se necesita configurar en caja, un Grupo por defecto", vbCritical, sMensaje
      End
   End If
   If Not RsGrupo.EOF And RsGrupo!nBoton > 0 Then
      cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
   Else
      RsGrupo.MoveFirst
      RsGrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vGrupo where nBoton>0 and lActivo=1", Cn) & "'"
      cmdGrupo(RsGrupo!nBoton).backColor = vbBlue
   End If
   xGrupo = RsGrupo!codigo
   RsSubgrupo.Filter = "tGrupo = '" & xGrupo & "'"
         
   'Asignar SubGrupo
   AsignaBoton 7, RsSubgrupo, cmdSubGrupo()
   RsSubgrupo.MoveFirst
   RsSubgrupo.Find "nBoton='" & Calcular("select min(nBoton) as Codigo FROM vSubGrupo where nBoton>0 and tGrupo = '" & xGrupo & "' and lActivo=1", Cn) & "'"
   If RsSubgrupo.EOF Or RsSubgrupo!nBoton = 0 Then
      MsgBox "Error: Se necesita configurar un SubGrupo con Boton", vbCritical, sMensaje
      End
   End If
   cmdSubGrupo(RsSubgrupo!nBoton).backColor = vbRed
   xSubGrupo = RsSubgrupo!codigo
   RsProducto.Requery
   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
      
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedidoPD, sUnidadNegocio
   
   'Asigna Operador
   If RsOperador.RecordCount > 0 Then
      AsignaBoton 13, RsOperador, cmdOperador()
      RsOperador.MoveFirst
      If Not IsNull(RsOperador!nBoton) And RsOperador!nBoton > 0 Then
         'Oscar Ortega----------------------------------------------
         'cmdOperador_Click (RsOperador!nBoton)
         HabilitaTimerColor (False)
         Dim ii As Integer
         Screen.MousePointer = vbHourglass
         For ii = 1 To 13
             cmdOperador(ii).backColor = vbButtonFace
         Next ii
         RsOperador.MoveFirst
         RsOperador.Find "nboton = " & Trim(str(RsOperador!nBoton))
         nOperadorPropiedad = RsOperador!nControl
         If RsOperador!nBoton > 0 Then
            cmdOperador(RsOperador!nBoton).backColor = vbRed
         End If
         If wAgregaCombo Then
            AsignaComboPropiedad
         Else
            AsignaPropiedad
         End If
         Screen.MousePointer = vbDefault
         HabilitaTimerColor (True)
         'Fin Oscar Ortega------------------------------------------
      End If
   End If
         
   If Not lInfhotel Then
      AsignaComando 4, RsTipoDocumento, cmdTipoDocumento()
      AsignaBoton 19, RsMozo, cmdMozo()
      AsignaBoton 19, RsMotorizado, cmdMotorizado()
      'origen de ventas
      AsignaBotonOrigenVentas 19, RsOrigenVentas, Me.cmdOrigen()
      
      fraPuntoVenta.Visible = False
      txtPuntoVenta.Visible = False
      cmdCabecera(14).Visible = False
      cmdCabecera(16).Visible = False
      cmdOpcion(14).Visible = False
   Else
      'Moneda Base
      sMonedaBase = Calcular("select tMoneda as Codigo from TPARAMETRO", CnInfhotel)
      
      'Punto de Venta
      Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
             " From tPuntoVenta " & _
             " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1"
      Set rsPuntoVenta = Lib.OpenRecordset(Isql, CnInfhotel)
   
      If rsPuntoVenta.RecordCount = 0 Then
         MsgBox "Error: Se necesita configurar los puntos de venta en Infhotel", vbCritical, sMensaje
         End
      End If
   
      Dim i As Integer
      For i = 10 To 19
          cmdMozo(i).Visible = False
          cmdMotorizado(i).Visible = False
      Next i
      
      txtFechaProg.Visible = False
      txtPuntoVenta.Visible = True
      Label1(25).Caption = "Pto. Venta :"
      
      cmdBuscar(3).Top = 1050
      fraMozo.Top = 3060 '3300 eraaaaaaaaaa
      fraMozo.Height = 1850 'eraaaaaaaaaaaa 1875
      AsignaBoton 9, RsMozo, cmdMozo()
            
      cmdBuscar(4).Top = 1050
      fraMotorizado.Top = 3060 ' 3300 eraaaaaaaaa
      fraMotorizado.Height = 1850 'eraaaaaaaaaaaa 1875
      AsignaComando 4, RsTipoDocumento, cmdTipoDocumento()
      AsignaBoton 9, RsMotorizado, cmdMotorizado()
      AsignaComando 9, rsPuntoVenta, cmdPunto()
      
      cmdCabecera(14).Visible = True
      cmdCabecera(16).Visible = True
      cmdOpcion(14).Visible = True
   End If
   
   'Obligatoriedad Mozo
   If lObligaMozo Then
      cmdMozo(1).Enabled = False
   End If
   
   'Tipo Documentos
   lTipoDocumento1 = IIf(cmdTipoDocumento(1).Caption = "", False, True)
   lTipoDocumento2 = IIf(cmdTipoDocumento(2).Caption = "", False, True)
   lTipoDocumento3 = IIf(cmdTipoDocumento(3).Caption = "", False, True)
   lTipoDocumento4 = IIf(cmdTipoDocumento(4).Caption = "", False, True)
   
   ActivaFrame fraPlato
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
   Else
      RsCabecera.MoveFirst
      sPedido = RsCabecera!codigo
   End If
   '-- cover ecuador
   'If lcover Then
        cmdCabecera(21).Visible = True
    'Else
        'cmdCabecera(21).Visible = False
   ''End If
   
   wDetalle = True
   wAgrega = False
   wAgregaCombo = False

   LblFecha.Caption = Format(FechaServidor(), "dddd, dd MMMM yyyy")
   lblFactura.Caption = NFactura
   lPropiedad = False
   lComboActivo = False
   LblMoneda.Caption = sMonN
   If sLey1 <> "" Then
    lblLey1.Visible = True
   End If
    If CD = True Or WEBAP = True Then
        ActivarDesactivarCD (True)
    Else
        ActivarDesactivarCD (False)
    End If
    
    If lActivaAnticipo Then
        Me.cmdCabecera(20).Visible = True
    Else
        Me.cmdCabecera(20).Visible = False
     End If

   nTC = Calcular("SELECT nVenta as Codigo From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)

   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaPedido()
    
    HabilitaTimerColor (False)

   Screen.MousePointer = vbHourglass
   
   'Llena los valores
   If RsCabecera.RecordCount = 0 Then
      sTipoPedido = sTipoPedidoPD
      sPedido = ""
      LimpiaCabecera
   Else
      
      txtEnvioAntes.Text = IIf(IsNull(RsCabecera!nMinutosAntesEnvio), 0, RsCabecera!nMinutosAntesEnvio)
      
      'invitado2013
      txtObservacion.Caption = ""
      'invitado2013
            'pariente2013
      txtCliente.Caption = ""
      'pariente2013
      sPedido = IIf(IsNull(RsCabecera!codigo), "", RsCabecera!codigo)
      sMesa = IIf(IsNull(RsCabecera!tMesa), "", RsCabecera!tMesa)
      
      sMozo = IIf(IsNull(RsCabecera!tMozo), "0000", RsCabecera!tMozo)
      sMotorizado = IIf(IsNull(RsCabecera!tMotorizado), "0000", RsCabecera!tMotorizado)
      sCliente = IIf(IsNull(RsCabecera!tClienteDelivery), "", RsCabecera!tClienteDelivery)
      sTienda = IIf(IsNull(RsCabecera!tTienda), "", RsCabecera!tTienda)
      sTipoAtencion = IIf(IsNull(RsCabecera!tTipoAtencion), "01", RsCabecera!tTipoAtencion)
      sTipoPedido = IIf(IsNull(RsCabecera!tTipoPedido), "01", RsCabecera!tTipoPedido)
      
      txtAdulto.Text = IIf(IsNull(RsCabecera!nAdulto), 0, RsCabecera!nAdulto)
      txtNino.Text = IIf(IsNull(RsCabecera!nNino), 0, RsCabecera!nNino)
      txtJuntar.Text = IIf(IsNull(RsCabecera!nMesa), 0, RsCabecera!nMesa)
   
      fraPedido.Caption = " Pedido " & IIf(RsCabecera.RecordCount = 0, "", IIf(IsNull(RsCabecera!codigo), "", RsCabecera!codigo)) & " "
      txtFecha.Caption = Format(IIf(IsNull(RsCabecera!fFecha), "", RsCabecera!fFecha), "dd/MM/yyyy HH:nn")
      txtTipoPedido.ForeColor = IIf(RsCabecera!lPrioridad, vbRed, &H800000)
     
      txtFechaProg.Caption = Format(IIf(IsNull(RsCabecera!fProgramacion), "", RsCabecera!fProgramacion), "dd/MM/yyyy HH:nn")
      txtMozo.Caption = IIf(IsNull(RsCabecera!Mozo), "", RsCabecera!Mozo)
      txtMotorizado.Caption = IIf(IsNull(RsCabecera!Motorizado), "", RsCabecera!Motorizado)
      txtMontoMax.Caption = IIf(IsNull(RsCabecera!nMontoMaximo), 0, RsCabecera!nMontoMaximo)
            'pariente'
      sCodigoParienteSeleccionado = IIf(IsNull(RsCabecera!tcodigopariente), "", RsCabecera!tcodigopariente)
      
      If sCodigoParienteSeleccionado <> "" Then

        txtCliente.Caption = Calcular("select pariente + ' ' +case when lconyugue=1 then ' Conyugue de ' else 'Hijo de ' end as codigo from vPariente where tCodigoPariente='" & sCodigoParienteSeleccionado & "' and tCodigoDelivery='" & sCliente & "'", Cn)

      End If
            txtCliente.Caption = txtCliente.Caption & " " & IIf(IsNull(RsCabecera!Cliente), "", RsCabecera!Cliente) & IIf(IsNull(RsCabecera!Tienda), "", " - " & RsCabecera!Tienda)

      'pariente2013
      Tienda = IIf(IsNull(RsCabecera!Tienda), "", RsCabecera!Tienda)
      
      txtDireccion.Caption = IIf(IsNull(RsCabecera!direccion), "", RsCabecera!direccion)
      If lClub Then
      txtTelefono.Caption = sCliente
      Else
      txtTelefono.Caption = IIf(IsNull(RsCabecera!Telefono), "", RsCabecera!Telefono)
        End If
      txtObservacion.Caption = IIf(IsNull(RsCabecera!tObservacion), "", RsCabecera!tObservacion)
      'entregarA
      txtEntregar.Caption = IIf(IsNull(RsCabecera!tentregara), "", RsCabecera!tentregara)
      
      xDescuento = IIf(IsNull(RsCabecera!nDescuento), 0, RsCabecera!nDescuento)
      sDescripcionDescuento = IIf(IsNull(RsCabecera!tObservacionDescuento), 0, RsCabecera!tObservacionDescuento)
      sCodigoDescuento = IIf(IsNull(RsCabecera!tDescuento), 0, RsCabecera!tDescuento)
      xMontoMaximo = IIf(IsNull(RsCabecera!nMontoMaximo), 0, RsCabecera!nMontoMaximo)
      TxtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
            
      'MODIFICAR DESCUENTO JULIO24
      sAutoriza = IIf(IsNull(RsCabecera!tUsuarioDescuento), 0, RsCabecera!tUsuarioDescuento)
      tAutorizaDescuento = IIf(IsNull(RsCabecera!tUsuarioDescuento), 0, RsCabecera!tUsuarioDescuento)
            'invitado2013
      sCodigoInvitado = IIf(IsNull(RsCabecera!tcodigoinvitado), "", RsCabecera!tcodigoinvitado)
'      If Len(Trim(RsCabecera!invitado)) > 0 Then
'      txtObservacion.Caption = txtObservacion.Caption & " Inv.: " & RsCabecera!invitado
'      End If
      'invitado2013
      vOrigenVentas = Calcular("select isnull(codigoorigenventa,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
      Me.txtOrdenCompra = IIf(IsNull(RsCabecera!tOrdenCompra), "", RsCabecera!tOrdenCompra)

      ' ELDCQ 07/07/2021
      Dim DescripcionOrigen As String
      DescripcionOrigen = Calcular("select isnull(Descripcion,'') as codigo from vOrigenVenta where CodOrigenVenta='" & vOrigenVentas & "'", Cn)
      txtTipoPedido.Caption = IIf(IsNull(RsCabecera!TipoPedido), "", RsCabecera!TipoPedido) & IIf(DescripcionOrigen = "0" Or DescripcionOrigen = "", "", " / " & DescripcionOrigen)
      
      If lcover Then
         grdDetalle.Caption = "Pedidos de " & IIf(sMesa = "000" Or Len(Trim(sMesa)) = 0, Mid(grdCabecera.Columns(2).Text, 1, 50), grdCabecera.Columns(1).Text) & " / Cover : " & val(RsCabecera!nAdulto) * sMontoMinCover
      Else
         grdDetalle.Caption = "Pedidos de " & IIf(sMesa = "000" Or Len(Trim(sMesa)) = 0, Mid(grdCabecera.Columns(2).Text, 1, 50), grdCabecera.Columns(1).Text)
      End If

      If lInfhotel Then
         sPuntoVenta = IIf(IsNull(RsCabecera!tPuntoVenta), "", RsCabecera!tPuntoVenta)
         sHabitacion = IIf(IsNull(RsCabecera!tHabitacion), "", RsCabecera!tHabitacion)
         sReserva = IIf(IsNull(RsCabecera!TRESERVA), "", RsCabecera!TRESERVA)
         sPasajero = IIf(IsNull(RsCabecera!tPasajero), "", RsCabecera!tPasajero)
         sCompania = IIf(IsNull(RsCabecera!tCompania), "", RsCabecera!tCompania)
         sContacto = IIf(IsNull(RsCabecera!tContacto), "", RsCabecera!tContacto)
         
         rsPuntoVenta.MoveFirst
         rsPuntoVenta.Find "Codigo='" & sPuntoVenta & "'"
         
         If Not rsPuntoVenta.EOF Then
            txtPuntoVenta.Caption = rsPuntoVenta!descripcion
         Else
            txtPuntoVenta.Caption = ""
         End If
      End If
      
    ' GCAA 12112020
    If RsCabecera!TipoPedido = "Delivery" Then
        Isql = "select isnull(tCodigoPedidoCD,'')tCodigoPedidoCD from mpedido where tCodigoPedido = '" & sPedido & "'"
         Set RsverCD = Lib.OpenRecordset(Isql, Cn)
            ' MsgBox (RsverCD!tCodigoPedidoCD)
            If RsverCD!tCodigoPedidoCD <> "" Then
                 frmVenta.Caption = " Caja : " & sCaja & " - PedidoCD: " & RsverCD!tCodigoPedidoCD
                 muestra = Me.Caption
            Else
                 frmVenta.Caption = " Caja : " & sCaja
                 muestra = Me.Caption
            End If
    Else
        frmVenta.Caption = " Caja : " & sCaja
    End If
      
      
      
      cmdCabecera(9).FontBold = IIf(txtTipoPedido.ForeColor = True, True, False)
      cmdCabecera(6).FontBold = IIf(sTipoPedido = "01", True, False)
      cmdCabecera(7).FontBold = IIf(sTipoPedido = "02", True, False)
      cmdCabecera(8).FontBold = IIf(sTipoPedido = "03", True, False)
      cmdCabecera(17).FontBold = IIf(sTipoPedido = "04", True, False)
      cmdCabecera(18).FontBold = IIf(sTipoPedido = "05", True, False)
      
      If sMesa <> "" Then
         If Calcular("select nValor as Codigo from vSalon where Codigo in(select tSalon from TMESA where tCodigoMesa='" & sMesa & "')", Cn) > 0 Then
            lPrecuenta = True
            sMozoD = ""
         Else
            lPrecuenta = False
            sMozoD = sMozo
         End If
      Else
         lPrecuenta = False
         sMozoD = sMozo
      End If
   End If
      
   RsDetalle.Filter = "tCodigoPedido = '" & sPedido & "'"
   
   'GCCA 02082021 - REQ REST-N014-2021 - TORI
   If (lNoVisualizarItemCombos) Then
        RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "' and automatico='" & 0 & "'"
   Else
        RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "'"
   End If
   
   If RsDetalle.RecordCount <> 0 Then
      RsDetalle.MoveFirst
      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sItem = RsDetalle!tItem
   Else
      sProducto = ""
      wCombo = False
      nCombo = 0
      sItem = ""
      
      ActivaCabecera True
      ActivaOpcion True
      ActivaFrame fraPlato
   End If
   RsProducto.Requery
   RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
   AsignaBotonProducto 20, RsProducto, cmdProducto(), sTipoPedido, sUnidadNegocio
   'Imprime el Total
   'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and isnull(tfacturado,'')='' and  tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
   SumPrecioVenta (sPedido)
   Screen.MousePointer = vbDefault
      
      HabilitaTimerColor (True)
      
       If lvisor Then
        Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
         Cn.Execute "insert into infovisor(id,tcaja,Pedido,estado) values(1,'" & sCaja & "','" & sPedido & "',1)"
                   
      End If
On Error Resume Next
   
   Dim CodigoReservaAnticipo As String
   CodigoReservaAnticipo = Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
    If CodigoReservaAnticipo = "" Then
        Me.cmdCabecera(20).FontBold = False
    Else
        Me.cmdCabecera(20).FontBold = True
    End If
      
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If lvisor Then
        Cn.Execute "delete from infovisor where tcaja='" & sCaja & "'"
    End If
    
       
'    If lBal Then
'       frmMsComm.MSCommBalanza.PortOpen = False
'    End If

    Set RsT3 = Nothing
    Set RsCabecera = Nothing
    Set RsDetalle = Nothing
    Set RsCombo = Nothing
    Set RsImpresion = Nothing
    Set RsGrupo = Nothing
    Set RsSubgrupo = Nothing
    Set RsProducto = Nothing
    Set RsMozo = Nothing
    Set RsMotorizado = Nothing
    Set RsArea = Nothing
    Set RsPropiedad = Nothing
    Set RsTipoDocumento = Nothing
    Set RsTimer = Nothing
    Set RsProductoCombo = Nothing
    Set frmVenta = Nothing
End Sub



Private Sub grdCabecera_DblClick()

  

On Error GoTo fin
         
    If MESA247 Then
      If RsCabecera.RecordCount > 0 Then
        sPedido = RsCabecera!codigo
      Else
        Exit Sub
      End If
        If Calcular("select count(*) as codigo from mpedido where isnull(idpedidomesa,'') <> '' and tcodigopedido='" & sPedido & "'", Cn) > 0 Then 'Calcular("SELECT CASE WHEN COUNT(*)>0 THEN 1 ELSE CASE WHEN (SELECT COUNT(*)  FROM MPEDIDO M WHERE M.TCODIGOPEDIDO='" & sPedido & "' AND isnull(M.idPedidoMesa,'')='' )>0 THEN 1 ELSE 0 END END  AS CODIGO FROM MPEDIDO MP INNER JOIN DPEDIDO DP ON MP.tCodigoPedido=DP.tCodigoPedido WHERE  MP.tcodigopedido='" & sPedido & "'", Cn) = 0 Then
           If MsgBox("Retornar Checking MESA24/7", vbYesNo) = vbYes Then
            If lPasswordC Then
               If Supervisor("01") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
            End If
            Isql = " exec usp_WS_MESA_Inforest '" & sPedido & "','',4"
'            Isql = "update mpedido set tTurno = '' , tCaja = '' ,tMozo=''  where tCodigoPedido ='" & sPedido & "'"
'            Cn.Execute Isql
'            Isql = "UPDATE TMESA SET TESTADOMESA='01' WHERE TCODIGOMESA= (SELECT TMESA FROM MPEDIDO WHERE TCODIGOPEDIDO ='" & sPedido & "')"
            Cn.Execute Isql
            RsCabecera.Requery
           End If
           
        End If
        
    End If
Exit Sub
fin:
    MsgBox ("error En proceso :" & Error)
End Sub

Private Sub grdCabecera_HeadClick(ByVal ColIndex As Integer)

    HabilitaTimerColor (False)

    Screen.MousePointer = vbHourglass
    RsCabecera.Sort = grdCabecera.Columns(ColIndex).DataField & " ASC"
    RsCabecera.Requery
    Screen.MousePointer = vbDefault

    ActivarDesactivarCD (True)
    HabilitaTimerColor (True)
End Sub

Public Sub ActivaCabecera(Activa As Boolean)
   cmdCabecera(0).Visible = Activa
   cmdCabecera(1).Visible = Activa
   cmdCabecera(2).Visible = Activa
   cmdCabecera(3).Visible = Activa
   cmdCabecera(4).Visible = Activa
   cmdCabecera(5).Visible = Activa
      
   cmdDetalle(0).Visible = Activa
   cmdDetalle(1).Visible = Activa
   cmdDetalle(2).Visible = Activa
   cmdDetalle(3).Visible = Activa
   cmdDetalle(4).Visible = Activa
   'cmdDetalle(5).Visible = Activa
   cmdDetalle(6).Visible = Activa
   cmdDetalle(7).Visible = Activa
   cmdDetalle(8).Visible = Activa
   
   cmdNavegar(0).Visible = Activa
   cmdNavegar(1).Visible = Activa
   cmdNavegar(2).Visible = Activa
   cmdNavegar(3).Visible = Activa
   cmdNavegar(4).Visible = Activa
   cmdNavegar(5).Visible = Activa
   cmdNavegar(6).Visible = Activa
   cmdNavegar(7).Visible = Activa
   cmdNavegar(8).Visible = Activa
   cmdNavegar(9).Visible = Activa
   cmdNavegar(10).Visible = Activa
   cmdNavegar(11).Visible = Activa
   
    'VSJ Valida si el bloqueo de modificación de los pedidos realizados desde integraciones 19.07.2024
    idPedidoMesa = Calcular("select idPedidoMesa as codigo from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
    If idPedidoMesa <> "0" Then
    sValidaIntegracion = True
    Else
    sValidaIntegracion = False
    End If
      
    If sValidaIntegracion = True And lbloqueoEdicionPedidoConsolaAgr = True Then
    cmdCabecera(2).Enabled = False
    cmdCabecera(3).Enabled = False
    cmdDetalle(0).Enabled = False
    cmdDetalle(2).Enabled = False
    cmdDetalle(3).Enabled = False
    cmdDetalle(4).Enabled = False
    cmdOpcion(11).Enabled = False
    
    Else
    cmdCabecera(2).Enabled = True
    cmdCabecera(3).Enabled = True
    cmdDetalle(0).Enabled = True
    cmdDetalle(2).Enabled = True
    cmdDetalle(3).Enabled = True
    cmdDetalle(4).Enabled = True
    cmdOpcion(11).Enabled = True
    End If
      
   If (fraProductoCombo.Visible = False And fraEliminacion.Visible = False And fraPropiedad.Visible = False) Then
    Me.FrmOrdenCompra.Visible = Not Activa
   End If
   
End Sub

Public Sub ActivaFrame(Frame As Frame)
    fraCabecera.Visible = False
    'fraDetalle.Visible = False
    fraPlato.Visible = False
    fraProductoCombo.Visible = False
    fraEliminacion.Visible = False
    fraPropiedad.Visible = False
    Frame.Visible = True
    Me.FrmOrdenCompra.Visible = False

     
End Sub

Public Sub ActivaOpcion(Activa As Boolean)
    cmdOpcion(0).Visible = Activa
    cmdOpcion(1).Enabled = Not Activa
    cmdOpcion(2).Enabled = Not Activa
    
    cmdOpcion(1).Visible = Not Activa
    cmdOpcion(2).Visible = Not Activa
    cmdOpcion(3).Visible = IIf(lPrinter, Activa, False)
    cmdOpcion(4).Visible = Activa
    cmdOpcion(5).Visible = Activa
    cmdOpcion(6).Visible = Activa
    cmdOpcion(7).Visible = IIf(lAdicion, Activa, False)
    cmdOpcion(9).Visible = Activa
    cmdOpcion(10).Visible = IIf(Not CD, Activa, False)
    cmdOpcion(11).Visible = Activa
    cmdOpcion(12).Visible = Activa
    cmdOpcion(14).Visible = IIf(lInfhotel, Activa, False)
    
    'origen de ventas
    Me.fraOrigenVentas.Visible = False
    '------------------------------------
    
    If lRapido = True Then
        cmdOpcion(15).Visible = IIf(lRapido, Activa, False)
    Else
        cmdOpcion(15).Visible = IIf(Not WEBAP, Activa, False)
    End If
        
    cmdOpcion(17).Visible = Activa
    cmdTipoDocumento(1).Visible = IIf(lTipoDocumento1, Activa, False)
    cmdTipoDocumento(2).Visible = IIf(lTipoDocumento2, Activa, False)
    cmdTipoDocumento(3).Visible = IIf(lTipoDocumento3, Activa, False)
    cmdTipoDocumento(4).Visible = IIf(lTipoDocumento4, Activa, False)
    'cmdDelivery.Visible = IIf(cmdDelivery.Visible, Activa, False)

End Sub

Public Sub InsertaProducto(xCantidad As Double)
    HabilitaTimerColor (False)

    Dim precioventa As Double
    Screen.MousePointer = vbHourglass
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim RsOrd As Recordset
    Dim nOrden As Integer
    Dim RsValidaDesc As Recordset
    
    'Restriccion de agregar productos con descuentos aplicados
    Set RsValidaDesc = Lib.OpenRecordset("select * from MPEDIDO where tcodigopedido='" & sPedido & "'", Cn)
    If RsValidaDesc!nDescuento > 0 Then
        MsgBox "Existe descuento aplicado, No se pueden agregar mas productos!!", vbExclamation, sMensaje
         Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'CPvalicacion central d pedido LG
   ' Dim codigoClienteF As String
    Dim lClienteExcluyeProducto As Boolean
    Dim lProductoPermiteDescuento As Boolean
    Dim lClienteControlaProducto As Boolean
  '  codigoClienteF = Calcular("select tclientedelivery as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn)
    
    If Calcular("select isnull(treservainf,'') as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn) <> "" Then
        MsgBox "Se ha aplicado Anticipo al Pedido!!!, no se puede Ingresar mas productos!!", vbInformation, sMensaje
         Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    If sCliente <> "" Then ' verificamos si tiene cliente frecuente
            If Calcular("select count(tcodigodelivery) as codigo from TCLIENTEPRODUCTO where tcodigoDelivery='" & sCliente & "'", Cn) > 0 Then
                lClienteExcluyeProducto = Calcular("select isnull(lexcluyeproductos,0) as codigo from tdelivery where tcodigodelivery='" & sCliente & "' ", Cn)
                If lClienteExcluyeProducto Then
                    If Calcular("select count(tcodigodelivery) as codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'  ", Cn) = 0 Then ' si el cliente tiene productos asociados
                                MsgBox "Según Configuración, este Producto no puede seleccionarse para el Cliente Frecuente Indicado"
                                Screen.MousePointer = vbDefault
                                Exit Sub
                    End If
                End If
                
                lClienteControlaProducto = False
                If Calcular("select count(tcodigoproducto) as codigo from tclienteproducto where tcodigodelivery='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn) > 0 Then
                    lClienteControlaProducto = True
                End If
                
            End If
    End If
      
    ' variables MULTIAREAPRODUCCION
    Dim lProductoMultiArea As Boolean
    Dim tsubalmacen As String
    Dim tAreaProduccion As String
        
   'INSUMOCRITICO23
    Dim rsInsumo As New ADODB.Recordset
    If Calcular("select isnull(lControlInsumoCritico,0) as codigo from tproducto where tcodigoproducto='" & sProducto & "'", Cn) = True Then
                    Set rsInsumo = Lib.OpenRecordset("select isnull(tcodigoinsumo,'') tcodigoinsumo , isnull(tinsumo.descripcion,'') ,isnull(nstock,0) from tproducto inner join tinsumo on tproducto.tcodigoinsumo =tinsumo.tcodigo where tproducto.tcodigoproducto='" & sProducto & "' and tinsumo.lactivo=1", Cn)
                    If Not (rsInsumo.EOF Or rsInsumo.BOF) Then
                        Me.Caption = muestra & "   Insumo Crítico --->   " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & " / Solicitado: " + str(xCantidad)
                    End If
    Else
        Me.Caption = muestra
    End If
    
    sItem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
    'INSUMOCRITICO
    If RsDetalle.RecordCount = 0 Then
       'sitem = "001"
       nOrden = 1
    Else
       'sitem = Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '" & sPedido & "'", Cn), 3)
       If lOrden Then
          Set RsOrd = Lib.OpenRecordset("select nOrden, lImprime from DPEDIDO where tCodigoPedido='" & sPedido & "' Order by nOrden DESC", Cn)
          If RsOrd.RecordCount > 0 Then
             If IIf(IsNull(RsOrd!lImprime), False, RsOrd!lImprime) Then
                nOrden = RsOrd!nOrden + 1
             Else
                nOrden = RsOrd!nOrden
             End If
          Else
             nOrden = 1
          End If
       Else
          nOrden = RsProducto!nOrden
       End If
    End If
    
    'Precios con Recargos / Descargos por Tipo de Pedido
    nRecargo = 0
    nDescuento = 0
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
    
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3
        
    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    ElseIf sTipoPedido = "04" Then
       If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    ElseIf sTipoPedido = "05" Then
       If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal5 = 0 Then
          nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nOficial = IIf(IsNull(RsProducto!nPrecioCanal5), 0, RsProducto!nPrecioCanal5)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto13, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto14, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto15, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto13
          lImp2 = RsProducto!lImpuesto14
          lImp3 = RsProducto!lImpuesto15
       End If
    Else
       nOficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    End If
    
    nOficial = IIf(RsProducto!tMONEDA = "02", nOficial * nTC, nOficial)
        
        
    'CPvalicacion LG
    If lClienteControlaProducto Then
       'gcaa 28/02/2020
       Dim vajaja As Integer
          Set ValProductoAsociado = Lib.OpenRecordset("  SP_INFOREST_VALPRODASOCIADOS '" & sProducto & "', '" & sTipoPedido & "', '" & sCliente & "','" & sPedido & "'", Cn)
             If Not (ValProductoAsociado.EOF Or ValProductoAsociado.BOF) Then
                vajaja = ValProductoAsociado.Fields(0)
            End If

        If vajaja = 1 Then
            nOficial = Calcular("SELECT ISNULL(NPRECIO,0) AS CODIGO FROM TCLIENTEPRODUCTO WHERE TCODIGODELIVERY='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn)
            nOficial = IIf(RsProducto!tMONEDA = "02", nOficial * nTC, nOficial)
            lProductoPermiteDescuento = Calcular("select isnull(lPermiteDescuentos,0) as codigo from tclienteproducto where TCODIGODELIVERY='" & sCliente & "' AND TCODIGOPRODUCTO='" & sProducto & "'", Cn)
        End If
    End If
    'CPvalicacion LG
    
    'Busca Oferta
    Dim sCriterio As String
    Dim nOferta As Double
    Dim tOferta As String
    Dim lAcumulable As Boolean
     'sCriterio = "tCodigoProducto ='" & sProducto & "' and lActivo=1 and lAutomatica=1"
    sCriterio = "tCodigoProducto ='" & sProducto & "' and lActivo=1  "
    sCriterio = sCriterio & " and (substring(tFrecuencia," & Weekday(FechaServidor(), vbMonday) & "+1,1) = '1' or (substring(tFrecuencia,1,1)='1') and MONTH(fFecha) = " & Month(FechaServidor()) & " AND DAY(fFecha)= " & Day(FechaServidor()) & ") and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "'"
    sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "')) "
    If sTipoPedido = "01" Then
       sCriterio = sCriterio & " and lLocal=1"
    ElseIf sTipoPedido = "02" Then
       sCriterio = sCriterio & " and lDelivery=1"
    ElseIf sTipoPedido = "03" Then
       sCriterio = sCriterio & " and lLlevar=1"
    ElseIf sTipoPedido = "04" Then
       sCriterio = sCriterio & " and lCanal4=1"
    Else
       sCriterio = sCriterio & " and lCanal5=1"
    End If
    tOferta = ""
    
    Isql = "select * from TOFERTA where tCodigoProducto ='" & sProducto & "' and lActivo=1  "
    Set RsExisteOferta = Lib.OpenRecordset(Isql, Cn)
    
    
    
    
    Isql = "select * from TOFERTA where " & sCriterio
    Set RsOferta = Lib.OpenRecordset(Isql, Cn)
    'GCAA 26022021 INFOREST 9015 BETA
    'Inserta descuento
    nOferta = 0
    lAcumulable = False
    If RsOferta.RecordCount > 0 And RsProducto!lDescuento Then
        If (RsOferta!lAutomatica) Then
            RsOferta.MoveFirst
            tOferta = RsOferta!tOferta
            If RsOferta!nPrecio > 0 Then
               nOferta = nOficial - IIf(IsNull(RsOferta!nPrecio), 0, RsOferta!nPrecio)
            ElseIf RsOferta!nMonto > 0 Then
               nOferta = RsOferta!nMonto
            Else
               nOferta = nOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
            End If
       End If
    Else
         'GCAA 05012021
             If RsExisteOferta.RecordCount > 0 Then
                 If Calcular("select lExcluyente as Codigo from TOFERTA where tCodigoProducto ='" & sProducto & "' and lActivo=1", Cn) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Este producto no puede ser cargado en esta franja horaria" & Chr(13) & "Consulte con el Manager", vbCritical, sMensaje
                    Exit Sub
'                  Else
'                     MsgBox "Fuera de rango" & Chr(13) & "Consulte con el Manager", vbCritical, sMensaje
'                    Exit Sub
                 End If
            End If
    End If
    
        'CPvalicacion LG
    If lClienteControlaProducto Then
          nPVenta = nOficial
                If lProductoPermiteDescuento Then
                            If xDescuento <> 0 And RsProducto!lDescuento Then
                               If RsOferta.RecordCount > 0 Then
                                  If RsOferta!lAcumulable Then
                                     nPVenta = (nPVenta - nOferta) - ((nPVenta - nOferta) * xDescuento / 100)
                                     nDescuento = nOficial - nPVenta
                                  Else
                                     nPVenta = nPVenta - nOferta
                                     nDescuento = nOficial - nPVenta
                                  End If
                               Else
                                  nPVenta = nPVenta - (nPVenta * xDescuento / 100)
                                  nDescuento = nOficial - nPVenta
                               End If
                               
                            Else
                               nPVenta = nPVenta - nOferta
                               nDescuento = nOficial - nPVenta
                            End If
                Else
'                    If nOficial > nPVenta Then
'                        nDescuento = nOficial - nPVenta
'                    Else
'                        nOficial = nPVenta
'                    End If
                    
                    nOferta = 0
                    
                End If
    Else
                        If xDescuento <> 0 And RsProducto!lDescuento Then
                               If RsOferta.RecordCount > 0 Then
                                  If RsOferta!lAcumulable Then
                                     nPVenta = (nOficial - nOferta) - ((nOficial - nOferta) * xDescuento / 100)
                                     nDescuento = nOficial - nPVenta
                                  Else
                                     nPVenta = nOficial - nOferta
                                     nDescuento = nOficial - nPVenta
                                  End If
                               Else
                                  nPVenta = nOficial - (nOficial * xDescuento / 100)
                                  nDescuento = nOficial - nPVenta
                               End If
                            Else
                               nPVenta = nOficial - nOferta
                               nDescuento = nOficial - nPVenta
                            End If
    
    End If
        'CPvalicacion LG
        
    Select Case pais 'ok
        Case "001" 'Bolivia
                nValor = (nValor / 100)
                nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
        'GCAA12012021
'                If RsProducto!tAplicaExoneracion = True Or RsProducto!tAplicaInafectacion Then
'                    nValor = 1 + (nValor / 100)
'                    nImpuesto1 = 0
'                    nImpuesto2 = 0
'                    nImpuesto3 = 0
'                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
'                Else
                    nValor = 1 + (nValor / 100)
                    nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                    nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                    nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                    nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
'                End If
                
    End Select
    
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMObra As Double
    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
    ElseIf sTipoPedido = "02" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
       nMObra = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
    ElseIf sTipoPedido = "03" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
       nMObra = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
    ElseIf sTipoPedido = "04" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
       nMObra = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
    Else
       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
       nMObra = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
    End If
    
    'Dim tAreaProduccion As String
      'multiarea produccion
    lProductoMultiArea = Calcular("select isnull(lmultiarea,0) as codigo from tproducto where tcodigoproducto='" & RsProducto.Fields("codigo") & "'", Cn)
    
    If lProductoMultiArea = False Then
        tsubalmacen = ""
    Else
        tsubalmacen = ""
        If lMultiAreaSubGrupo = True Then
            tAreaProduccion = Calcular("select isnull(tarea,'') codigo from TAREASUBGRUPO where tcaja='" & sCaja & "' and tSubGrupo='" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If
        If lMultiAreaCaja = True Then
            tAreaProduccion = Calcular("select isnull(tsubalmacen,'') as codigo from tcaja where tcaja='" & sCaja & "'", Cn)
            tsubalmacen = Calcular("select isnull(tvalor,'')  as codigo from varea where codigo='" & tAreaProduccion & "'", Cn)
        End If

        If tsubalmacen = "0" Then
            tsubalmacen = ""
        End If
    
    End If
    
    If lBloqInafecto And Calcular("select count(*) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn) > 0 And Round(nPVenta, 2) > 0 Then
        Dim IgvPED As Double
        IgvPED = CDbl(Calcular("select top 1 round(nprecioimpuesto1,2) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn))
        If IgvPED > 0 Then
            If Round(nImpuesto1, 2) <= 0 Then
                MsgBox ("No es posible ingresar el producto, no se puede ingresar Items inafecto y afecto en un mismo pedido!! "), vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            If Round(nImpuesto1, 2) > 0 Then
                MsgBox ("No es posible ingresar el producto, no se puede ingresar Items inafecto y afecto en un mismo pedido!! "), vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    
    'If lcover Then
        Dim maxped As Double
        maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
            If maxped > 0 Then
                   If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (nPVenta * xCantidad)) Then
                     MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                     
                     Screen.MousePointer = vbDefault
                    
                     RsDetalle.Requery
                     
                      'RsDetalle.MoveLast
                     'RsProducto.MoveFirst
                     Exit Sub

                   End If
            End If
    'End If
    
    
'MsgBox (RsProducto!tAplicaExoneracion)
    fxCombo "A", 1, sProducto
    'diacontable
    'GCAA - 2409202 SE AGREGA CAMPOS POR EL REQUERIMIENTO 0023-2020
    'NOMBRE DE CAMPOS : tAplicaExoneracion, tAplicaInafectacion
    'TABLA: DPEDIDO
    
    Isql = "insert into DPEDIDO " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, tEstadoItem, tComanda, fRegistro, tMozoD, tUsuarioD, nInsumo, nGasto, nManoObra, nOrden, tUnidadNegocio, tOferta, tsubalmacen,fdiacontable,tcajad ,tAplicaExoneracion, tAplicaInafectacion  ) " & _
           "Values(   '" & sPedido & "', " _
                   & "'" & sTipoPedido & "', '" & sItem & "', '" & sProducto & "', " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nPBase & ", " & nRecargo & ", " & nDescuento & ", " & nOficial & ", " _
                   & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                   & nPVenta & ", " & xCantidad & ", " & nPVenta * xCantidad & ", " _
                   & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                   & "0, '" & RsProducto!tArea & "', " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                   & IIf(RsProducto!lCombinacion, -1, 0) & "," _
                   & IIf(IsNull(RsProducto!nCombinacion), 0, RsProducto!nCombinacion) & "," _
                   & "'N', '" & sComanda & "', getDate(), '" & sMozoD & "', '" & sUsuario & "', " _
                   & nInsumo & ", " _
                   & nGasto & ", " _
                   & nMObra & ", " _
                   & nOrden & ", '" & sUnidadNegocio & "','" & tOferta & "','" & tsubalmacen & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', '" & sCaja & "', " _
                   & IIf(RsProducto!tAplicaExoneracion, -1, 0) & " , " _
                   & IIf(RsProducto!tAplicaInafectacion, -1, 0) & ")"
    Cn.Execute Isql
                    
    RsDetalle.Requery
    nCombo = IIf(IsNull(RsDetalle!nCombinacion), 1, RsDetalle!nCombinacion)
    'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and isnull(tfacturado,'')='' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
    SumPrecioVenta (sPedido)
    RsDetalle.MoveLast
    
    Screen.MousePointer = vbDefault
    
     If (RsProducto!lCombinacion) And (RsProducto!lComboVisible) Then
    lComboActivo = True
''       MsgBox ("es un combo !!!"), vbInformation
'        If RsCabecera.RecordCount = 0 Then
'           Exit Sub
'        End If
'
'        wAgregaCombo = True
'        ActivaCabecera False
'        ActivaOpcion False
'        If Not RsCombo.EOF Then
'           RsCombo.MoveFirst
'        End If
'        AsignaProductoCombo
'
'        CmdOpcion(1).Visible = False
'        CmdOpcion(2).Visible = False
'        txtBarra.SetFocus
''                 fraCombo.Visible = True
''                fraProductoCombo.Visible = True
'        ActivaFrame fraCombo
'        ActivaFrame fraProductoCombo
    End If
   
     HabilitaTimerColor (True)
End Sub

Public Sub AsignaProducto()
   If Not RsDetalle.EOF Then
      sProducto = IIf(IsNull(RsDetalle!tCodigoProducto), "", RsDetalle!tCodigoProducto)
      ListarOperadoresConFiltro (sProducto) 'Oscar Ortega
      sGrupo = IIf(IsNull(RsDetalle!tCodigoGrupo), "", RsDetalle!tCodigoGrupo)
      sSubGrupo = IIf(IsNull(RsDetalle!tCodigoSubGrupo), "", RsDetalle!tCodigoSubGrupo)
      sItem = IIf(IsNull(RsDetalle!tItem), "001", RsDetalle!tItem)
      nOrden = IIf(IsNull(RsDetalle!nOrden), 0, RsDetalle!nOrden)
      lblObservacion.Text = IIf(IsNull(RsDetalle!tObservacion), "", RsDetalle!tObservacion)
      nDescuento = IIf(IsNull(RsDetalle!nDescuento), 0, RsDetalle!nDescuento)
      nRecargo = IIf(IsNull(RsDetalle!nRecargo), 0, RsDetalle!nRecargo)
      
      nImpuesto1 = IIf(IsNull(RsDetalle!nprecioImpuesto1), 0, RsDetalle!nprecioImpuesto1)
      nImpuesto2 = IIf(IsNull(RsDetalle!nprecioImpuesto2), 0, RsDetalle!nprecioImpuesto2)
      nImpuesto3 = IIf(IsNull(RsDetalle!nprecioImpuesto3), 0, RsDetalle!nprecioImpuesto3)
      sCortesia = IIf(IsNull(RsDetalle!tCortesia), "", RsDetalle!tCortesia)
      nOficial = IIf(IsNull(RsDetalle!nPrecioOficial), 0, RsDetalle!nPrecioOficial)
      nPBase = IIf(IsNull(RsDetalle!nPrecioNeto), 0, RsDetalle!nPrecioNeto)
      nPVenta = IIf(IsNull(RsDetalle!nprecioVenta), 0, RsDetalle!nprecioVenta)
      nCantidad = IIf(IsNull(RsDetalle!nCantidad), 0, RsDetalle!nCantidad)
                    
        If IIf(IsNull(RsDetalle!lImprime), False, RsDetalle!lImprime) = False Then
                'luchoinsumos
                 verificatitulo
                 'luchoinsumos
        Else
                Me.Caption = muestra
        End If
                                       
      'Llena el Combo
      fraCombo.Caption = IIf(IsNull(RsDetalle!Producto), "", " " & RsDetalle!Producto & " ")
      wCombo = IIf(IsNull(RsDetalle!lCombinacion), False, RsDetalle!lCombinacion)
      nCombo = IIf(IsNull(RsDetalle!nCombinacion), 1, RsDetalle!nCombinacion)
      
      'GCCA 02082021 - REQ REST-N014-2021 - TORI
      If (lNoVisualizarItemCombos) Then
            RsCombo.Filter = "[tCodigoPedido]='" & sPedido & "' and [tItem]='" & sItem & "'and [automatico]='" & 0 & "'"
      Else
            RsCombo.Filter = "[tCodigoPedido]='" & sPedido & "' and [tItem]='" & sItem & "'"
      End If
      
      fraCombo.Visible = False
       If wCombo = True Then
        sProductoCombo = sProducto
      End If
      wAgregaCombo = False
      ActivaCabecera True
      ActivaOpcion True
      ActivaFrame fraPlato
      AsignaPropiedad
      
'      idPedidoMesa = Calcular("select idPedidoMesa as codigo from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
'      If idPedidoMesa <> "" Then
'      sValidaIntegracion = True
'      Else
'      sValidaIntegracion = False
'      End If
'
'      If sValidaIntegracion = True Then
''      cmdCabecera(14).Visible = IIf(lInfhotel, Activa, False)
''      cmdCabecera(14).Visible = IIf(lInfhotel, Activa, False)
''      cmdCabecera(14).Visible = IIf(lInfhotel, Activa, False)
'      End If
      
      
      TxtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
   End If
   
   HabilitaTimerColor (True)
End Sub

Public Sub AsignaCombo()
   If Not RsCombo.EOF Then
      sCombo = IIf(IsNull(RsCombo!tProductoCombo), "", RsCombo!tProductoCombo)
      sProducto = sCombo
      xItem = IIf(IsNull(RsCombo!tItemCombo), "001", RsCombo!tItemCombo)
      lblObservacion.Text = IIf(IsNull(RsCombo!tObservacion), "", RsCombo!tObservacion)
      ListarOperadoresConFiltro (sCombo)
      AsignaComboPropiedad
   End If
End Sub

Public Sub AsignaPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
        'Oscar Ortega--
        If RsOperador.EOF Then
            RsOperador.MoveFirst
        End If
        '--------------
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = IIf(IsNull(RsOperador!nControl), 0, RsOperador!nControl)
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sProducto & "'"
       nOperadorPropiedad = 0
    End If
    
    AsignaComando 20, RsPropiedad, cmdPropiedad()

   'Validaciòn de Visibilidad de Propiedades -BR
    With RsPropiedad
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 1 To IIf(.RecordCount > 20, 20, .RecordCount)
                If IsNull(!visibilidad) Or !visibilidad = False Then
                    cmdPropiedad(i).Visible = True
                Else
                    If (!visibilidad) Then
                         Select Case RsCabecera!tTipoPedido
                            Case "01"

                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal1) And RsPropiedad!canal1
                            Case "02"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal2) And RsPropiedad!canal2
                            Case "03"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal3) And RsPropiedad!canal3
                            Case "04"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal4) And RsPropiedad!canal4
                            Case "05"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal5) And RsPropiedad!canal5
                            Case Else
                                cmdPropiedad(i).Visible = False
                        End Select
                    Else
                         cmdPropiedad(i).Visible = False
                    End If
                End If
                cmdPropiedad(i).Caption = !descripcion
                .MoveNext
            Next i
        Else
               For i = 1 To 20
                   cmdPropiedad(i).Visible = False
               Next i
        End If
    End With


    For i = 1 To 20
        cmdPropiedad(i).FontBold = False
'        If RsOperador!nBoton = 0 And nPropiedadVisible = 0 Then
'            cmdPropiedad(i).Visible = False
'        End If
    Next i
    
    nPropiedadVisible = 1
    
    lblResumen.Text = ""
    
    
'    If Not (RsCabecera.BOF And RsCabecera.EOF) Then
'        ' --- El recordset tiene datos ---
'        Dim encontrado As Boolean
'        encontrado = False
'
'        RsCabecera.MoveFirst
'        Do Until RsCabecera.EOF
'            If Trim$(RsCabecera.Fields("tCodigoPedido").value & "") = sPedido And sPedido <> "" Then
'                sPedido = RsCabecera.Fields("tCodigoPedido").value
'                encontrado = True
'                Exit Do
'            End If
'            RsCabecera.MoveNext
'        Loop
'
'        ' Si no lo encontró, regreso al primero
'        If Not encontrado And sPedido <> "" Then
'            RsCabecera.MoveFirst
'            sPedido = RsCabecera.Fields("tCodigoPedido").value
'        End If
'    Else
'        ' --- El recordset está vacío ---
'        sPedido = ""
'    End If

'Dim existe As Boolean
'existe = False
'
'If Not (RsCabecera.BOF And RsCabecera.EOF) And sPedido <> "" Then
'    RsCabecera.Filter = "tCodigoPedido='" & Replace(sPedido, "'", "''") & "'"
'    If Not (RsCabecera.BOF And RsCabecera.EOF) Then existe = True
'    RsCabecera.Filter = adFilterNone   ' quitar filtro
'
'End If




    ' --- Uso en tu bucle ---
RsProductoPropiedad.Filter = "tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'"
If Not (RsProductoPropiedad.BOF And RsProductoPropiedad.EOF) Then
    RsProductoPropiedad.MoveFirst
    Do While Not RsProductoPropiedad.EOF
        For i = 1 To 20
            Dim propDesc As String
            Dim operadorDesc As String

            propDesc = FieldStr(RsProductoPropiedad, "descripcion")
            operadorDesc = FieldStr(RsProductoPropiedad, "Operador")

            ' Antes de comparar con RsOperador, asegúrate que RsOperador tiene datos
            If Not (RsOperador Is Nothing) And Not (RsOperador.BOF And RsOperador.EOF) Then
                Dim operDescOperador As String
                operDescOperador = FieldStr(RsOperador, "descripcion")
                
                ' compara de forma segura (sin Nulls y sin evaluar partes que pueden fallar)
                If propDesc <> "" And operDescOperador <> "" Then
                    If Trim$(cmdPropiedad(i).Caption & "") = propDesc And operDescOperador = operadorDesc Then
                        cmdPropiedad(i).FontBold = True
                        Exit For
                    End If
                End If
            End If
        Next i

        ' Construir el lblResumen (también usando propDesc / operadorDesc)
        If propDesc <> "" Then
            If RsProductoPropiedad!nCantidad = 1 Then
                lblResumen.Text = lblResumen.Text & LTrim(operadorDesc) & " " & LTrim(propDesc) & ", "
            Else
                lblResumen.Text = lblResumen.Text & LTrim(operadorDesc) & " " & LTrim(propDesc) & ": (" & RsProductoPropiedad!nCantidad & "), "
            End If
        End If

        RsProductoPropiedad.MoveNext
    Loop
End If
End Sub

Public Sub AsignaProductoCombo()
    Dim i As Integer
    RsProductoCombo.Filter = "tCombo = '" & sProducto & "'"
    'AsignaComando 48, RsProductoCombo, cmdProductoCombo()
    AsignaComandoColor 48, RsProductoCombo, cmdProductoCombo()
End Sub

Private Sub grdCabecera_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If RsCabecera.RecordCount > 0 And RsCabecera.AbsolutePosition > 0 Then
      
      AsignaPedido
   End If
    
   On Error Resume Next
   txtBarra.SetFocus
End Sub

Private Sub grdCombo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  AsignaCombo
  If fraPropiedad.Visible = True Then
     cmdOpcion_Click (8)
  End If
End Sub

Private Sub grdDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

   If RsCabecera.RecordCount > 0 Then
   nPropiedadVisible = 0
      If fraPropiedad.Visible = True Then
         nPos = RsDetalle.AbsolutePosition
         RsDetalle.Requery
         RsDetalle.AbsolutePosition = nPos
      End If
     
      AsignaProducto
      
      If lPropiedad Then
         lPropiedad = False
         cmdDetalle_Click (4)
      End If
      
       If lComboActivo Then
         lComboActivo = False
         cmdOpcion_Click (11)
      End If
      
      
   End If
   On Error Resume Next
   txtBarra.SetFocus
   'oo2
   ActivarDesactivarCD (True)
  '  frmVenta.Caption = " Caja : " & sCaja & "jajajajajaj"
End Sub




Private Sub Timer_Timer()
   'Pedido Timer
   RsTimer.Requery
   Dim i As Integer
   If RsTimer.RecordCount <> 0 Then
      RsTimer.MoveFirst
      While Not RsTimer.EOF()
            If Format(RsTimer!fProgramacion, "dd/mm/yyyy") = Format(FechaServidor(), "dd/mm/yyyy") And RsTimer!nMinutosAntesEnvio >= DateDiff("n", Format(Time, "hh:mm"), Format(RsTimer!fProgramacion, "hh:mm")) Then
            
               'Screen.MousePointer = vbHourglass
               Isql = "select * from vPedido Where lImprime = 0 And lImprimeArea = 1 and Codigo = '" & RsTimer!codigo & "' " & _
                      "ORDER BY tItem,tetiqueta,combo"
                
               Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
                       
               If RsImpresion.RecordCount = 0 Then
                  LimpiaRs
                  'MsgBox "No existen Productos Nuevos a Imprimir", vbExclamation, sMensaje
               Else
                  RsArea.MoveFirst
                  For i = 1 To RsArea.RecordCount
                      RsImpresion.Filter = "Codigo= '" & RsTimer!codigo & "' and tArea = '" & RsArea!tArea & "'"
                      
                      If RsImpresion.RecordCount <> 0 Then
                         RsImpresion.MoveFirst
                         ImprimePedido RsImpresion, "N", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                      End If
                      RsArea.MoveNext
                  Next i
                  Cn.Execute "Update DPEDIDO Set lImprime = 1 where tCodigoPedido = '" & RsTimer!codigo & "'"
                  RsDetalle.Requery
               End If
               LimpiaRs
            
               'Isql = "Update MPEDIDO Set tTipoAtencion = '01' where tCodigoPedido = '" & RsTimer!codigo & "'"
               'Cn.Execute Isql
               'RsCabecera.Requery
            End If
            RsTimer.MoveNext
      Wend
   End If
End Sub

Public Sub InsertaCombo(wProducto As String)
    Screen.MousePointer = vbHourglass
    Dim xItem As String
    Dim nValor As Double
    Dim nCNeto As Double
    Dim nCImp1 As Double
    Dim nCImp2 As Double
    Dim nCImp3 As Double
    Dim nCVenta As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim nInsumo As Double
    Dim nGasto As Double
    Dim nMano As Double
    
    'GCCA 02082021 - REQ REST-N014-2021 - TORI
     If (lNoVisualizarItemCombos) Then
       xItem = Lib.Correlativo(Calcular("select isnull(max(tItemCombo),'000') as codigo from CPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn), 3)
     Else
      If RsCombo.RecordCount = 0 Then
           xItem = "001"
        Else
           xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from CPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn), 3)
        End If
     End If
     
    nValor = 0
    nValor = nValor + IIf(RsProducto!lImpuesto1, nPorcentaje1, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto2, nPorcentaje2, 0)
    nValor = nValor + IIf(RsProducto!lImpuesto3, nPorcentaje3, 0)
      
    lImp1 = RsProducto!lImpuesto1
    lImp2 = RsProducto!lImpuesto2
    lImp3 = RsProducto!lImpuesto3
          
    If sTipoPedido = "02" Then
       If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioDelivery), 0, RsProducto!nPrecioDelivery)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto4, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto5, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto6, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto4
          lImp2 = RsProducto!lImpuesto5
          lImp3 = RsProducto!lImpuesto6
       End If
    ElseIf sTipoPedido = "03" Then
       If IsNull(RsProducto!nPreciollevar) Or RsProducto!nPreciollevar = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPreciollevar), 0, RsProducto!nPreciollevar)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto7, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto8, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto9, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto7
          lImp2 = RsProducto!lImpuesto8
          lImp3 = RsProducto!lImpuesto9
       End If
    ElseIf sTipoPedido = "04" Then
       If IsNull(RsProducto!nPrecioCanal4) Or RsProducto!nPrecioCanal4 = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto10, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto11, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto12, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto10
          lImp2 = RsProducto!lImpuesto11
          lImp3 = RsProducto!lImpuesto12
       End If
    ElseIf sTipoPedido = "05" Then
       If IsNull(RsProducto!nPrecioCanal5) Or RsProducto!nPrecioCanal4 = 0 Then
          nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) + (nLlevar * IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta) / 100)
       Else
          nCVenta = IIf(IsNull(RsProducto!nPrecioCanal4), 0, RsProducto!nPrecioCanal4)
          nValor = 0
          nValor = nValor + IIf(RsProducto!lImpuesto13, nPorcentaje1, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto14, nPorcentaje2, 0)
          nValor = nValor + IIf(RsProducto!lImpuesto15, nPorcentaje3, 0)
          lImp1 = RsProducto!lImpuesto13
          lImp2 = RsProducto!lImpuesto14
          lImp3 = RsProducto!lImpuesto15
       End If
    Else
       nCVenta = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
    End If
    nCVenta = IIf(RsProducto!tMONEDA = "02", nCVenta * nTC, nCVenta)
    
    Select Case pais 'ok
        Case "001" 'Bolivia
                nValor = (nValor / 100)
                nCImp1 = IIf(lImp1, nCVenta * nPorcentaje1 / 100, 0)
                nCImp2 = IIf(lImp2, nCVenta * nPorcentaje2 / 100, 0)
                nCImp3 = IIf(lImp3, nCVenta * nPorcentaje3 / 100, 0)
                nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
        
        Case Else 'Peru, Ecuador
                nValor = 1 + (nValor / 100)
                nCImp1 = IIf(lImp1, nCVenta / nValor * nPorcentaje1 / 100, 0)
                nCImp2 = IIf(lImp2, nCVenta / nValor * nPorcentaje2 / 100, 0)
                nCImp3 = IIf(lImp3, nCVenta / nValor * nPorcentaje3 / 100, 0)
                nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
    End Select
                   
                   
    If sTipoPedido = "01" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMano = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)
    ElseIf sTipoPedido = "02" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo2), 0, RsProducto!nInsumo2)
       nGasto = IIf(IsNull(RsProducto!nGasto2), 0, RsProducto!nGasto2)
       nMano = IIf(IsNull(RsProducto!nManoObra2), 0, RsProducto!nManoObra2)
    ElseIf sTipoPedido = "03" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo3), 0, RsProducto!nInsumo3)
       nGasto = IIf(IsNull(RsProducto!nGasto3), 0, RsProducto!nGasto3)
       nMano = IIf(IsNull(RsProducto!nManoObra3), 0, RsProducto!nManoObra3)
    ElseIf sTipoPedido = "04" Then
       nInsumo = IIf(IsNull(RsProducto!nInsumo4), 0, RsProducto!nInsumo4)
       nGasto = IIf(IsNull(RsProducto!nGasto4), 0, RsProducto!nGasto4)
       nMano = IIf(IsNull(RsProducto!nManoObra4), 0, RsProducto!nManoObra4)
    Else
       nInsumo = IIf(IsNull(RsProducto!nInsumo5), 0, RsProducto!nInsumo5)
       nGasto = IIf(IsNull(RsProducto!nGasto5), 0, RsProducto!nGasto5)
       nMano = IIf(IsNull(RsProducto!nManoObra5), 0, RsProducto!nManoObra5)
    End If
    
    'If lcover Then
        Isql = "select * from TCombo Where tCombo = '" & RsDetalle!tCodigoProducto & "' And tCodigoProducto = '" & wProducto & "'"
        Dim RstCombo2 As Recordset
        Set RstCombo2 = Lib.OpenRecordset(Isql, Cn)
        If RstCombo2.RecordCount > 0 Then
                
                 Dim maxped As Double
                 maxped = CDbl(Calcular("select isnull(nMontoMaximo,0) as codigo from mpedido where tcodigopedido='" & sPedido & "'", Cn))
                     If maxped > 0 Then
                            If maxped < (CDbl(Calcular("select sum(nVenta) as codigo from dpedido where tcodigopedido='" & sPedido & "'", Cn)) + (RstCombo2!nAumento)) Then
                              MsgBox ("No se puede ingresar el producto,Monto Maximo del Pedido: " & xMontoMaximo & " !!!"), vbInformation
                               RsCombo.Requery
                              Exit Sub
    
                            End If
                     End If
        End If
    'End If
    
    Dim xCantCombo As Double
    xCantCombo = Calcular("Select ISNULL(nCantidad,0) as Codigo From TCOMBO where tCombo = '" & RsDetalle!tCodigoProducto & "' and tCodigoProducto='" & wProducto & "'", Cn)
    
                          
    Isql = "insert into CPEDIDO " & _
           "(tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
           "Values(   '" & sPedido & "', " _
                   & "'" & RsDetalle!tCodigoProducto & "', " _
                   & "'" & sItem & "', " _
                   & "'" & xItem & "', " _
                   & "'" & wProducto & "',  " & xCantCombo & " , " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nCNeto & ", " _
                   & nCImp1 & ", " _
                   & nCImp2 & ", " _
                   & nCImp3 & ", " _
                   & nCVenta & ", " _
                   & nInsumo & ", " & nGasto & ", " & nMano & ", " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & ", 0, " _
                   & RsProducto!nOrden & ") "
    Cn.Execute Isql
    'Oscar Ortega------------------------------------------------------------
    Isql = "select * from TCombo Where tCombo = '" & RsDetalle!tCodigoProducto & "' And tCodigoProducto = '" & wProducto & "'"
    Dim RstCombo As Recordset
    Set RstCombo = Lib.OpenRecordset(Isql, Cn)
    If RstCombo.RecordCount > 0 Then
        If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
            txtSuma.Caption = CambiaPrecio(nPVenta + (RstCombo!nAumento * RstCombo!nCantidad) / nCantidad, sPedido)
        End If
    End If
    'Fin Oscar Ortega--------------------------------------------------------
    
    RsCombo.Requery
    RsCombo.MoveLast
    AsignaCombo
    Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaComboPropiedad()
    Dim i As Integer
    If RsOperador.RecordCount > 0 Then
       If RsOperador.EOF Then
        RsOperador.MoveFirst
       End If
       RsPropiedad.Filter = "tOperador = '" & RsOperador!codigo & "' and tProducto='" & sCombo & "'"
    Else
       RsPropiedad.Filter = "tOperador = '  ' and tProducto='" & sCombo & "'"
    End If
    AsignaComando 20, RsPropiedad, cmdPropiedad()
    'Validaciòn de Visibilidad de Propiedades -BR
    With RsPropiedad
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 1 To IIf(.RecordCount > 20, 20, .RecordCount)
                If IsNull(!visibilidad) Or !visibilidad = False Then
                    cmdPropiedad(i).Visible = True
                Else
                    If (!visibilidad) Then
                         Select Case RsCabecera!tTipoPedido
                            Case "01"

                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal1) And RsPropiedad!canal1
                            Case "02"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal2) And RsPropiedad!canal2
                            Case "03"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal3) And RsPropiedad!canal3
                            Case "04"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal4) And RsPropiedad!canal4
                            Case "05"
                                cmdPropiedad(i).Visible = Not IsNull(RsPropiedad!canal5) And RsPropiedad!canal5
                            Case Else
                                cmdPropiedad(i).Visible = False
                        End Select
                    Else
                         cmdPropiedad(i).Visible = False
                    End If
                End If
                cmdPropiedad(i).Caption = !descripcion
                .MoveNext
            Next i
        Else
               For i = 1 To 20
                   cmdPropiedad(i).Visible = False
               Next i
        End If
    End With
    
    For i = 1 To 20
        cmdPropiedad(i).FontBold = False
    Next i
    
    lblResumen.Text = ""
    RsComboPropiedad.Filter = "tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "' and tItemCombo='" & xItem & "'"
    If Not RsComboPropiedad.EOF Then
       RsComboPropiedad.MoveFirst
       Do While Not RsComboPropiedad.EOF
          For i = 1 To 20
              If cmdPropiedad(i).Caption = RsComboPropiedad!descripcion And RsOperador!descripcion = RsComboPropiedad!Operador Then
                 cmdPropiedad(i).FontBold = True
                 Exit For
              End If
          Next i
          lblResumen.Text = lblResumen.Text & LTrim(RsComboPropiedad!Operador) & " " & LTrim(RsComboPropiedad!descripcion) & ", "
          RsComboPropiedad.MoveNext
       Loop
    End If
End Sub
Public Sub LimpiaCabecera()
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
      txtFecha.Caption = ""
      txtTipoPedido.Caption = ""
      txtMozo.Caption = ""
      txtMotorizado.Caption = ""
      txtCliente.Caption = ""
      txtTelefono.Caption = ""
      TxtDescuento.Caption = "0.00"
      txtDireccion.Caption = ""
      txtObservacion.Caption = ""
      fraPlato.Visible = True
      txtSuma.Caption = "0.00"
      txtMontoMax.Caption = "0"
      vOrigenVentas = ""
       'entregarA
      txtEntregar.Caption = ""
      Me.txtOrdenCompra = ""
   End If
End Sub

Public Sub fxCombo(Funcion As String, Cantidad As Double, Combo As String)
   Dim i As Integer
   Dim xItem As String
   Dim RsTemp As Recordset
   Dim nValor As Double
   Dim nCNeto As Double
   Dim nCImp1 As Double
   Dim nCImp2 As Double
   Dim nCImp3 As Double
   Dim nCVenta As Double
   Dim lImp1 As Boolean
   Dim lImp2 As Boolean
   Dim lImp3 As Boolean
   Dim nInsumo As Double
   Dim nGasto As Double
   Dim nMano As Double
   
   If sTipoPedido = "01" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, isnull(dbo.TPRODUCTO.lImpuesto1,0) as lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo As nInsumo, dbo.TPRODUCTO.nGasto As nGasto, dbo.TPRODUCTO.nManoObra As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "02" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo2 As nInsumo, dbo.TPRODUCTO.nGasto2 As nGasto, dbo.TPRODUCTO.nManoObra2 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
   
   ElseIf sTipoPedido = "03" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo3 As nInsumo, dbo.TPRODUCTO.nGasto3 As nGasto, dbo.TPRODUCTO.nManoObra3 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "04" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo4 As nInsumo, dbo.TPRODUCTO.nGasto4 As nGasto, dbo.TPRODUCTO.nManoObra4 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   ElseIf sTipoPedido = "05" Then
        Isql = "SELECT dbo.TCOMBO.tCombo, dbo.TCOMBO.tCodigoProducto, dbo.TPRODUCTO.tGrupo, dbo.TPRODUCTO.tSubGrupo, dbo.TPRODUCTO.nPrecioVenta, dbo.TPRODUCTO.nPrecioLlevar, dbo.TPRODUCTO.nPrecioDelivery, dbo.TPRODUCTO.lImpuesto1, dbo.TPRODUCTO.lImpuesto2, dbo.TPRODUCTO.lImpuesto3, dbo.TPRODUCTO.lImpuesto4, dbo.TPRODUCTO.lImpuesto5, dbo.TPRODUCTO.lImpuesto6, dbo.TPRODUCTO.lImpuesto7, dbo.TPRODUCTO.lImpuesto8, dbo.TPRODUCTO.lImpuesto9, dbo.TPRODUCTO.tMoneda, dbo.TPRODUCTO.lImprimeArea, dbo.TSUBGRUPO.nOrden, dbo.TPRODUCTO.nInsumo5 As nInsumo, dbo.TPRODUCTO.nGasto5 As nGasto, dbo.TPRODUCTO.nManoObra5 As nManoObra, TCOMBO.nCantidad " & _
               "FROM dbo.TSUBGRUPO RIGHT OUTER JOIN dbo.TPRODUCTO ON dbo.TSUBGRUPO.tCodigoSubGrupo = dbo.TPRODUCTO.tSubGrupo RIGHT OUTER JOIN dbo.TCOMBO ON dbo.TPRODUCTO.tCodigoProducto = dbo.TCOMBO.tCodigoProducto Where dbo.TCOMBO.tCombo = '" & Combo & "' and dbo.TCOMBO.lFijo=1"
         
   End If
   
   Set RsTemp = Lib.OpenRecordset(Isql, Cn)
   If RsTemp.RecordCount = 0 Then
      Exit Sub
   End If
   
   'GCCA 02082021 - REQ REST-N014-2021 - TORI
    If (lNoVisualizarItemCombos) Then
        RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "' and automatico='" & 0 & "'"
    Else
        RsCombo.Filter = "tCodigoPedido = '" & sPedido & "' and tItem='" & sItem & "'"
    End If
   
   
   Select Case Funcion
          Case Is = "A"
               Do While Not RsTemp.EOF
               'GCCA 02082021 - REQ REST-N014-2021 - TORI
                If (lNoVisualizarItemCombos) Then
                     xItem = Lib.Correlativo(Calcular("select isnull(max(tItemCombo),'000') as codigo from CPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn), 3)
                Else
                    If RsCombo.RecordCount = 0 Then
                         xItem = "001"
                      Else
                         xItem = Lib.Correlativo(Calcular("select max(tItemCombo) as codigo from CPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'", Cn), 3)
                      End If
                End If
                
                  nValor = 0
                  nValor = nValor + IIf(RsTemp!lImpuesto1, nPorcentaje1, 0)
                  nValor = nValor + IIf(RsTemp!lImpuesto2, nPorcentaje2, 0)
                  nValor = nValor + IIf(RsTemp!lImpuesto3, nPorcentaje3, 0)
                   ' Exit Sub
                  lImp1 = RsTemp!lImpuesto1
                  lImp2 = RsTemp!lImpuesto2
                  lImp3 = RsTemp!lImpuesto3
                        
                  If sTipoPedido = "02" Then
                     If IsNull(RsProducto!nPrecioDelivery) Or RsProducto!nPrecioDelivery = 0 Then
                        nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) + (nDELIVERY * IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) / 100)
                     Else
                        nCVenta = IIf(IsNull(RsTemp!nPrecioDelivery), 0, RsTemp!nPrecioDelivery)
                        nValor = 0
                        nValor = nValor + IIf(RsTemp!lImpuesto4, nPorcentaje1, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto5, nPorcentaje2, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto6, nPorcentaje3, 0)
                        lImp1 = RsTemp!lImpuesto4
                        lImp2 = RsTemp!lImpuesto5
                        lImp3 = RsTemp!lImpuesto6
                     End If
                  ElseIf sTipoPedido = "03" Then
                     If IsNull(RsTemp!nPreciollevar) Or RsTemp!nPreciollevar = 0 Then
                        nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) + (nLlevar * IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta) / 100)
                     Else
                        nCVenta = IIf(IsNull(RsTemp!nPreciollevar), 0, RsTemp!nPreciollevar)
                        nValor = 0
                        nValor = nValor + IIf(RsTemp!lImpuesto7, nPorcentaje1, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto8, nPorcentaje2, 0)
                        nValor = nValor + IIf(RsTemp!lImpuesto9, nPorcentaje3, 0)
                        lImp1 = RsTemp!lImpuesto7
                        lImp2 = RsTemp!lImpuesto8
                        lImp3 = RsTemp!lImpuesto9
                     End If
                  Else
                     nCVenta = IIf(IsNull(RsTemp!nprecioVenta), 0, RsTemp!nprecioVenta)
                  End If
                  nInsumo = IIf(IsNull(RsTemp!nInsumo), 0, RsTemp!nInsumo)
                  nGasto = IIf(IsNull(RsTemp!nGasto), 0, RsTemp!nGasto)
                  nMano = IIf(IsNull(RsTemp!nManoObra), 0, RsTemp!nManoObra)
                    
                  nCVenta = IIf(RsTemp!tMONEDA = "02", nCVenta * nTC, nCVenta)
               
                  Select Case pais ' ok
                        Case "001" 'Bolivia
                                    nValor = (nValor / 100)
                                    nCImp1 = IIf(lImp1, nCVenta * nPorcentaje1 / 100, 0)
                                    nCImp2 = IIf(lImp2, nCVenta * nPorcentaje2 / 100, 0)
                                    nCImp3 = IIf(lImp3, nCVenta * nPorcentaje3 / 100, 0)
                                    nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
                                    
                        Case Else 'Peru, Ecuador
                                    nValor = 1 + (nValor / 100)
                                    nCImp1 = IIf(lImp1, nCVenta / nValor * nPorcentaje1 / 100, 0)
                                    nCImp2 = IIf(lImp2, nCVenta / nValor * nPorcentaje2 / 100, 0)
                                    nCImp3 = IIf(lImp3, nCVenta / nValor * nPorcentaje3 / 100, 0)
                                    nCNeto = nCVenta - nCImp1 - nCImp2 - nCImp3
                                    
                  End Select
                  
                  Isql = "insert into CPEDIDO " & _
                         "(tCodigoPedido, tProducto, tItem, tItemCombo, tProductoCombo, nCantidad, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, nInsumo, nGasto, nManoObra, lImprimeArea, lImprime, nOrden ) " & _
                         "Values(   '" & sPedido & "', " _
                                 & "'" & Combo & "', " _
                                 & "'" & sItem & "', " _
                                 & "'" & xItem & "', " _
                                 & "'" & RsTemp!tCodigoProducto & "', " & RsTemp!nCantidad & ", " _
                                 & "'" & IIf(IsNull(RsTemp!tgrupo), "", RsTemp!tgrupo) & "', " _
                                 & "'" & IIf(IsNull(RsTemp!tSubGrupo), "", RsTemp!tSubGrupo) & "', " _
                                 & nCNeto & ", " _
                                 & nCImp1 & ", " _
                                 & nCImp2 & ", " _
                                 & nCImp3 & ", " _
                                 & nCVenta & ", " _
                                 & nInsumo & ", " & nGasto & ", " & nMano & ", " _
                                 & IIf(RsTemp!lImprimeArea, -1, 0) & ", 0, " _
                                 & RsTemp!nOrden & ") "

                  Cn.Execute Isql
                  
                'Oscar Ortega------------------------------------------------------------
                Isql = "select * from TCombo Where tCombo = '" & Combo & "' And tCodigoProducto = '" & RsTemp!tCodigoProducto & "'"
                Dim RstCombo As Recordset
                Set RstCombo = Lib.OpenRecordset(Isql, Cn)
                If RstCombo.RecordCount > 0 Then
                    If IIf(IsNull(RstCombo!nAumento), 0, RstCombo!nAumento) > 0 Then
                        txtSuma.Caption = CambiaPrecio(nPVenta + RstCombo!nAumento, sPedido)
                    End If
                End If
                'Fin Oscar Ortega--------------------------------------------------------
                  
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          Case Is = "M"
               Do While Not RsTemp.EOF
                  Dim X As Double
                  X = Calcular("select nCantidad as Codigo FROM TCOMBO where tCombo='" & RsTemp!tCombo & "' and tCodigoproducto ='" & RsTemp!tCodigoProducto & "'", Cn)
                  Isql = "update CPEDIDO set nCantidad = " & X * Cantidad & " where tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "' and tProductocombo='" & RsTemp!tCodigoProducto & "'"
                  Cn.Execute Isql
                  RsCombo.Requery
                  RsTemp.MoveNext
               Loop
          
          Case Is = "D"
               For i = 1 To Cantidad
                   Isql = "DELETE from CPEDIDO " & _
                          "where tCodigoPedido ='" & sPedido & "' and tProducto='" & Combo & "' and tItem='" & sItem & "'"
                   Cn.Execute Isql
                   RsCombo.Requery
               Next i
   End Select
   Set RsTemp = Nothing
End Sub



Private Sub txtBarra_KeyPress(KeyAscii As Integer)
    ' Productos por Lectora de barras
    If KeyAscii = 13 And Trim(txtBarra.Text) <> "" Then
        
        If lRotulado = True Then
              'CESAR ROTULADO
              Dim rCodigoEtiqueta As String
              Dim rCodigoProducto As String
              Dim rLenBarra As String
              Dim rCantidad As Double
              Dim X As Integer
              
              rLenBarra = Len(Trim(txtBarra.Text))
              X = rLenBarra - 31
              
              If nLongitudBarra <> 0 Then
                rCodigoProducto = Mid(txtBarra.Text, nLongitudBarra + 1, 7)
                rCodigoEtiqueta = Mid(txtBarra.Text, 1, nLongitudBarra)
                
                If lCapturaPeso Then
                   rCantidad = val(Mid(txtBarra.Text, 31 + 1, X))
                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
                Else
                   rCantidad = 1
                   InsertaProductoRotulado rCodigoProducto, rCantidad, rCodigoEtiqueta
                End If
              Else
                  MsgBox "Error: Longitud de barra no registrada", vbCritical, sMensaje
              End If
              txtBarra.Text = ""
        
        Else
              If RsCabecera.RecordCount <> 0 Then
                 RsProducto.Filter = adFilterNone
                 RsProducto.MoveFirst
                 RsProducto.Find "tbarra = '" & Trim(txtBarra.Text) & "'"
                  
                  If Not RsProducto.EOF Then
                    'INSUMOCRITICO23
                            If validadIngresoProducto(RsProducto!codigo) = False Then
                                Exit Sub
                            End If
                   'INSUMOCRITICO23
                  End If
             
                  If Not RsProducto.EOF Then
                    sProducto = RsProducto!codigo
                    sComanda = IIf(sTemp = "", "", Mid("0000000000", 1, 10 - Len(Trim(sTemp))) & Trim(sTemp))
        
                    If lBal And RsProducto!lBalanza Then
                       Dim nResultado As Double
                       nResultado = Pesar(nBalanzaPuerto)
                       If (lActivaTresDecimales) Then
'                        nResultado = ""
                        nResultado = Format(nResultado, "#,##0.000")
                       Else
'                        nResultado = ""
                        nResultado = Format(nResultado, "#,##0.00")
                       End If
                       If nResultado > 0 Then
                          InsertaProducto nResultado
                       End If
                    Else
                    nCantidad = 1
                       InsertaProducto 1
                    End If
        
                    If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                       lPropiedad = True
                    End If
                  Else
                    If nLongitudBarra > 0 Then
                       RsProducto.MoveFirst
                       RsProducto.Find "tbarra = '" & Trim(Mid(txtBarra.Text, 1, nLongitudBarra)) & "'"
                       If Not RsProducto.EOF Then
                          sProducto = RsProducto!codigo
                          sComanda = IIf(sTemp = "", "", Mid("0000000000", 1, 10 - Len(Trim(sTemp))) & Trim(sTemp))
                          Dim nCantidadBarra As Double
        
                          If lCapturaPeso Then
                          
                            If EAN13 Then
                                nCantidadBarra = val(Mid(txtBarra.Text, nLongitudBarra + 1, 1) + "." + Mid(txtBarra.Text, nLongitudBarra + 2, 3))
                             Else
                                nCantidadBarra = val(Mid(txtBarra.Text, nLongitudBarra + 1, 2) + "." + Mid(txtBarra.Text, nLongitudBarra + 3, 4))
                             End If
                             
                             InsertaProducto nCantidadBarra
                          Else
                             nCantidadBarra = val(Mid(txtBarra.Text, nLongitudBarra + 1, 3) + "." + Mid(txtBarra.Text, nLongitudBarra + 4, 3))
                             InsertaProducto CalculaCantidad(nCantidadBarra)
                          End If
                          
                          If IIf(IsNull(RsProducto!lPropiedad), False, RsProducto!lPropiedad) Then
                             lPropiedad = True
                          End If
                       Else
                          MsgBox "Producto no encontrado", vbCritical, sMensaje
                       End If
                    Else
                          MsgBox "Producto no encontrado", vbCritical, sMensaje
                    End If
                 End If
                 RsProducto.Filter = "tSubGrupo = '" & xSubGrupo & "'"
              End If
              txtBarra.Text = ""
              sProducto = ""
        
        End If
    Else
        'txtBarra.Text = ""
    End If
     
End Sub


Public Function CalculaCantidad(nPrecio As Double) As Double
    Select Case sTipoPedido
           Case "01"
                CalculaCantidad = nPrecio / RsProducto!nprecioVenta
           Case "02"
                If RsProducto!nPrecioDelivery > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioDelivery
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
           
           Case "03"
                If RsProducto!nPreciollevar > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPreciollevar
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
                
           Case "04"
                If RsProducto!nPrecioCanal4 > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioCanal4
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
                   
           Case "05"
                If RsProducto!nPrecioCanal5 > 0 Then
                   CalculaCantidad = nPrecio / RsProducto!nPrecioCanal5
                Else
                   CalculaCantidad = nPrecio / RsProducto!nprecioVenta
                End If
    End Select
      
End Function

Public Sub EliminaCabecera()
    Call KDS_EliminarOrden(RsCabecera) 'KDS
    Dim i As Integer
    Dim sMotivo As String
    'cambio mango
    Screen.MousePointer = vbHourglass
    'Impresion del Pedidos Anulados
    sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
    If sCodigo = "000" Then
    sMotivo = sMotivo & " (" & sDescrip & ")"
    End If
   
   If lPrinter Then
      Isql = "select *, '" & sMotivo & "' as  MotivoEliminacion FROM dbo.vPedido LEFT OUTER JOIN dbo.vMotivoEliminacion ON dbo.vPedido.tMotivoEliminacion = dbo.vMotivoEliminacion.Codigo " & _
             "WHERE vPedido.Codigo='" & sPedido & "' AND lImprime=1 AND lImprimeArea=1 " & _
             "ORDER BY tItem,tetiqueta,combo"
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
              
       If Not RsImpresion.EOF Then
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea='" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                    'If (lCompatibilidadTVS) Then 'TVS
                    '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Anulacion", sPedido)
                    'End If
                 End If
              End If
              RsArea.MoveNext
          Next i
       End If
   End If
      
   If lInfhotel Then
      CnInfhotel.Execute "update MCOMANDA set TESTADO ='04', TOBSERVACIONANULA = 'Anulado por Inforest - " & sUsuarioAutoriza & " " & RsCabecera!codigo & " - " & Trim(sDescrip) & "' where tComanda ='" & RsCabecera!tComanda & "'"
   End If
   If lAlmacen Then
    Isql = "Update mRequerimiento Set lPedido = 0, tPedido = '' Where tPedido = '" & sPedido & "'"
    CnAlmacen.Execute Isql
   End If
   
   Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa ='" & sMesa & "'"
   
   'Juntar Mesa
   Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
   Cn.Execute "delete from TPEDIDOMESA where tCodigoPedido='" & sPedido & "'"
   
    'INSUMOCRITICO23
    Dim rstItems As New ADODB.Recordset
    Dim j As Integer
    Set rstItems = New ADODB.Recordset
    'Set rstItems = Lib.OpenRecordset("select tcodigoinsumo,ncantidad from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto where tcodigopedido='" & sPedido & "'  and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(dpedido.limprime,0)=1", Cn)
    
    Set rstItems = Lib.OpenRecordset("usp_Inforest_RevertirInsumosCriticosCabecera '" & sPedido & "'", Cn)
    If Not (rstItems.EOF Or rstItems.BOF) Then
        rstItems.MoveFirst
        For j = 0 To rstItems.RecordCount - 1
            modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
            rstItems.MoveNext
        Next j
    End If
   'INSUMOCRITICO23
   
   Cn.Execute "Update MPEDIDO set tEstadoPedido ='03', tMotivoAnulacion='" & sCodigo & "', tUsuarioAnulado='" & sUsuarioAutoriza & "', fRegAnulado= getdate(), tTurnoAnulado='" & sTurno & "', tObservacionAnulado='" & sDescrip & "'  where tCodigoPedido ='" & sPedido & "'"
   Cn.Execute "Update DPEDIDO Set tEstadoItem = 'A' where tCodigoPedido = '" & sPedido & "'"
   Cn.Execute "delete TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete CPEDIDO where tCodigoPedido='" & sPedido & "'"
   Cn.Execute "delete TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "'"
   
'   Dim CodReserva As String
'
'   CodReserva = Calcular("Select ISNULL(tReservaInf,'') as codigo from Mpedido where tCodigoPedido = '" & sPedido & "'", Cn)
'
'   If CodReserva <> "" Then
'        If Calcular("Select count(*) as codigo from TRESERVA where tReserva='" & CodReserva & "'", Cn) > 0 Then
'            Cn.Execute "Update TRESERVA Set tEstadoReserva='01' where tReserva = '" & CodReserva & "'"
'        End If
'   End If
   
   xTiporecepcion = Calcular("Select isnull(tTiporecepcion,0) as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
   xttipopedido = Calcular("Select ttipopedido as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
   
   'registro de log eliminacion del pedido
    Call Log_Inforest("PUNTO VENTA", "ELIMINA PEDIDO", sPedido, Error, "", "Eliminación del pedido " & sPedido, "Eliminación del pedido " & sPedido, "Se ha eliminado el pedido " & sPedido, sUsuario)
   
   
   'ELIMINA CENTRAL DELIVERY Y WEB APP
    If (CD = True Or WEBAP = True) And xTiporecepcion <> 0 Then
         
         
         If xTiporecepcion = 1 Then
                 Isql = "Select tCodigoPedidoCD From MPEDIDO where tCodigoPedido = '" & sPedido & "'"
                 Set RsPedido = Lib.OpenRecordset(Isql, Cn)
                 If RsPedido.RecordCount > 0 Then
                     Do While Not RsPedido.EOF
                         If Not IsNull(RsPedido!tCodigoPedidoCD) Then
                             EliminarPedidoCD (RsPedido!tCodigoPedidoCD)
                         End If
                         RsPedido.MoveNext
                     Loop
                 End If
         'ElseIf xTiporecepcion = 2 Then
'                 Isql = "Select tCodigoPedidoWebAp From MPEDIDO where tCodigoPedido = '" & sPedido & "'"
'                 Set RsPedido = Lib.OpenRecordset(Isql, Cn)
'                 If RsPedido.RecordCount > 0 Then
'                     Do While Not RsPedido.EOF
'                         If Not IsNull(RsPedido!tCodigoPedidoWebAp) Then
'                             Call ModificarEstadoPedidoWebAp(RsPedido!tCodigoPedidoWebAp, "3")
'                         End If
'                         RsPedido.MoveNext
'                     Loop
'                 End If
         End If
    End If
              

   RsCabecera.Requery
   RsDetalle.Requery
   RsProductoPropiedad.Requery
   RsCombo.Requery
   RsComboPropiedad.Requery
   AsignaPedido
   ActivaOpcion True
   
   If RsCabecera.RecordCount = 0 Then
      sPedido = ""
   Else
      RsCabecera.MoveFirst
      sPedido = RsCabecera!codigo
      ActivaFrame fraPlato
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Function CalculaDescuento() As Boolean
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
  
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM dbo.DPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoPedido='" & sPedido & "' and lDescuento=1", Cn)
        
'    If Not (xDescuento > nSuma) Then
        If RsDetalle.RecordCount <> 0 Then

    RsDetalle.MoveFirst
   
    Do While Not RsDetalle.EOF
      'Busca Oferta
      nPVenta = 0
      
      sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
      sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(FechaServidor(), vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(FechaServidor(), "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
      sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "'))"
        
      Isql = "select * from TOFERTA where " & sCriterio
      
      Set RsOferta = Lib.OpenRecordset(Isql, Cn)
      lAcumulable = True
      
      nOferta = 0
      Acumulado = 0
      
      If RsOferta.RecordCount > 0 Then
         RsOferta.MoveFirst
         lAcumulable = RsOferta!lAcumulable
         nOferta = RsDetalle!nPrecioOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
      End If
      
      If RsDetalle!lDescuento And lAcumulable = True Then
         If Calcular("select lRatio as Codigo FROM vMotivoDescuento where Codigo='" & sCodigoDescuento & "'", Cn) Then
            'nPVenta = Format((RsDetalle!nPrecioOficial - nOferta) - ((RsDetalle!nPrecioOficial - nOferta) * xDescuento / 100), "###,###,###,##0.000")
            nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((RsDetalle!nPrecioOficial - nOferta) * xDescuento / 100)
         Else
            Dim xPorc As Double
            xPorc = (RsDetalle!nPrecioOficial - nOferta) * RsDetalle!nCantidad * 100 / nSuma
            'nPVenta = Format((RsDetalle!nPrecioOficial - nOferta) - ((xPorc * xDescuento / 100) / RsDetalle!nCantidad), "###,###,###,##0.000")
            nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((xPorc * xDescuento / 100) / RsDetalle!nCantidad)
         End If
         
          Select Case pais ' ok
            Case "001" 'Bolivia
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
          
            Case Else 'Peru, Ecuador
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = 1 + (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
          End Select
          
         Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                "nRecargo = " & nRecargo & ", " & _
                "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                "nPrecioVenta = " & nPVenta & ", " & _
                "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                "nCantidad = " & RsDetalle!nCantidad & ", " & _
                "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                "tCortesia = '" & sCortesia & "' " & _
                "where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido ='" & sPedido & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
'End If


    End If
        
End Function

Public Sub EliminaItem()
    Dim xMax As String
    Dim sMotivo As String
    
    If lPrinter Then
       'Impresion del Pedidos Anulados
       sMotivo = Calcular("select Descripcion as Codigo from vMotivoEliminacion where Codigo='" & sCodigo & "'", Cn)
       'cambio mangos
       If sCodigo = "000" Then
        sMotivo = sMotivo & " (" & sDescrip & ")"
       End If
       
       Isql = "select *, '" & sMotivo & "' as MotivoEliminacion FROM dbo.vPedido " & _
              "WHERE Codigo = '" & sPedido & "' and tItem = '" & sItem & "' and lImprime = 1 And lImprimeArea = 1 " & _
              "ORDER BY tItem,tetiqueta,combo"
                    
       Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
       Dim i As Integer
       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
       Else
          RsArea.MoveFirst
          For i = 1 To RsArea.RecordCount
              RsImpresion.Filter = "tArea = '" & RsArea!tArea & "'"
              If RsArea!tIcono = "" Or RsCabecera!tSalon = RsArea!tIcono Or ((IsNull(RsCabecera!tSalon) Or RsCabecera!tSalon = "") And RsArea!nValor = 1) Then
                 If RsImpresion.RecordCount <> 0 Then
                    RsImpresion.MoveFirst
                    ImprimePedido RsImpresion, "A", RsArea!timpresora, RsArea!Area, False, RsProductoPropiedad, RsComboPropiedad, "Inforest"
                    'If (lCompatibilidadTVS) Then 'TVS
                    '    Call TVS_EnviarTicket(RsImpresion, RsArea!Area, "Anulacion", sPedido)
                    'End If
                 End If
              End If
              RsArea.MoveNext
          Next i
          RsDetalle.Requery
       End If
    End If
    
    xMax = Calcular("select max(tItem) as Codigo from APEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
    xMax = Lib.Correlativo(xMax, 3)
    Isql = "insert into APEDIDO (tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, " & _
           "tComanda, lImprime, tUsuario, fRegistro, tUsuarioAnulado, fRegistroAnulado, " & _
           "tObservacion, tObservacionAnulado, tEstadoItem, lImprimeArea, tArea, tMotivoEliminacion, tTurnoAnulado,fDiaContable) " & _
           "select '" & sPedido & "' as tCodigoPedido, '" & xMax & "' as tItem, tCodigoProducto, tCodigoGRupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, nRecargo, " & _
           "nDescuento, nPrecioOficial, nCantidad, nImpuesto1, nImpuesto2, nImpuesto3, nVenta, tComanda, lImprime, " & _
           "'" & RsCabecera!tUsuario & "' as tUsuario, dpedido.fregistro as fRegistro, " & _
           "'" & sUsuarioAutoriza & "' as tUsuarioAnulado, getDate() as fRegistroAnulado, " & _
           "tObservacion, '" & sDescrip & "' as tObservacion, tEstadoItem, lImprimeArea, tArea, '" & sCodigo & "', '" & sTurno & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "' " & _
           "from DPEDIDO " & _
           "where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'"
    Cn.Execute Isql
        'DIACONTABLE
     'INSUMOCRITICO23
    Dim rstItems As New ADODB.Recordset
    Set rstItems = New ADODB.Recordset
    'Set rstItems = Lib.OpenRecordset("select tcodigoinsumo,ncantidad from dpedido inner join tproducto on dpedido.tcodigoproducto=tproducto.tcodigoproducto where tcodigopedido='" & sPedido & "' and titem='" & sitem & "' and tproducto.lControlInsumoCritico=1 and isnull(tproducto.tcodigoinsumo,'')<>''  and isnull(dpedido.limprime,0)=1 ", Cn)
    Set rstItems = Lib.OpenRecordset("  usp_Inforest_RevertirInsumosCriticos '" & sPedido & "','" & sItem & "' ", Cn)

    If Not (rstItems.EOF Or rstItems.BOF) Then
        modificaStockInsumo rstItems.Fields(0), rstItems.Fields(1), "I"
    End If
    
    Isql = Calcular("select  tCodigoProducto AS Codigo from DPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)
    Cn.Execute "delete from TPRODUCTOPROPIEDAD where tCodigoPedido='" & sPedido & "' and tItem='" & sItem & "'"
    Cn.Execute "delete from TCOMBOPROPIEDAD where tCodigoPedido='" & sPedido & "' and tItem ='" & sItem & "'"
    Cn.Execute "delete from DPEDIDO where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'"
    Cn.Execute "delete from CPEDIDO where tCodigoPedido ='" & sPedido & "' and tItem ='" & sItem & "'"
    
    'If RsParametro!lActivoKDS Then
    'Cn.Execute "delete from DPEDIDO_ORDEN where tCodigoPedido = '" & sPedido & "' and tItem = '" & sItem & "'"
    '--Colocar para eliminar del DPEDIDO_ORDEN y no afectar KDS
    'End If
    
    'Registra la eliminacion del producto
    Call Log_Inforest("PUNTO VENTA", "ELIMINA PRODUCTO", sPedido, Error, "", "item: " & sItem & " - Producto: " & Isql & " - Pedido: " & sPedido, "Eliminó el producto " & Isql & ". Pedido " & sPedido, "Se ha eliminado el producto " & Isql, sUsuario)
    
    RsProductoPropiedad.Requery
    RsComboPropiedad.Requery
    RsCombo.Requery
    RsDetalle.Requery
    If RsDetalle.RecordCount = 0 Then
       txtSuma.Caption = "0.00"
    Else
       'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
       SumPrecioVenta (sPedido)
    End If
                
    If Not RsDetalle.EOF Then
       RsDetalle.MoveLast
       sItem = RsDetalle!tItem
    Else
       sProducto = ""
       wCombo = False
       nCombo = 0
       sItem = ""
              
       ActivaFrame fraPlato
       wAgregaCombo = False
       ActivaCabecera True
       ActivaOpcion True
    End If
    
    
      'MESA247
            If MESA247 Then
                  If Calcular("select isnull(idPedidoMesa,'') as codigo from MPEDIDO where tCodigopedido='" & sPedido & "'", Cn) <> "" Then
                        Cn.Execute "update MPEDIDO set lEnvioMesa = 0 where tCodigoPedido='" & sPedido & "'"
                  End If
            End If
    
End Sub

Sub PUltimaComanda()
Dim RsPuntoVentaU   As ADODB.Recordset
    Isql = "Select tPuntoVenta as Codigo, tDescripcion as Descripcion, nUltimoComanda, tmoneda" & _
           " From tPuntoVenta " & _
           " where tHotel='" & sHotel & "' AND lActivo=1 and lInforest=1 and tPuntoVenta='" & sPuntoVenta & "'"
    Set RsPuntoVentaU = Lib.OpenRecordset(Isql, CnInfhotel)
    UltimaComanda = IIf(IsNull(RsPuntoVentaU!nUltimoComanda), "", RsPuntoVentaU!nUltimoComanda)
    Set RsPuntoVentaU = Nothing
End Sub

Public Sub GrabaProducto()
   Screen.MousePointer = vbHourglass
   Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nDescuento = " & nDescuento & ", " & _
           "nRecargo = " & nRecargo & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & sPedido & "'"
           Cn.Execute Isql
           RsDetalle.Requery
           RsDetalle.MoveFirst
           RsDetalle.Find "tItem = '" & sItem & "'"
   'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
   SumPrecioVenta (sPedido)
   Screen.MousePointer = vbDefault
End Sub

Public Function CambiaPrecio(Valor As Double, Pedido As String)
    nPVenta = val(Valor)
    nOficial = nPVenta
    Dim Acumulado As Double
    Select Case pais
        Case "001" 'Bolivia
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = (Acumulado / 100)
                
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta * nPorcentaje3 / 100, 0)
                
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = 1 + (Acumulado / 100)
                
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
    
    End Select
           
           
    Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nPrecioOficial = " & nOficial & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & Pedido & "'"
           Cn.Execute Isql

    CambiaPrecio = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and isnull(tfacturado,'')='' and tCodigoPedido='" & Pedido & "'", Cn), "#,###,##0.00")
End Function


Public Function CambiaPrecioCombo(Valor As Double, Pedido As String)
    nPVenta = val(Valor)
    'nOficial = nPVenta
    Dim Acumulado As Double
    Select Case pais
        Case "001" 'Bolivia
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = (Acumulado / 100)
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
        
        Case Else 'Peru, Ecuador
                Acumulado = 0
                Acumulado = IIf(nImpuesto1 > 0, Acumulado + nPorcentaje1, Acumulado)
                Acumulado = IIf(nImpuesto2 > 0, Acumulado + nPorcentaje2, Acumulado)
                Acumulado = IIf(nImpuesto3 > 0, Acumulado + nPorcentaje3, Acumulado)
                Acumulado = 1 + (Acumulado / 100)
                nImpuesto1 = IIf(nImpuesto1 > 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                nImpuesto2 = IIf(nImpuesto2 > 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                nImpuesto3 = IIf(nImpuesto3 > 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
    
    End Select
           '"nPrecioOficial = " & nOficial & ", "
           
    Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
           "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
           "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
           "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
           "nPrecioVenta = " & nPVenta & ", " & _
           "nventa = " & nPVenta * nCantidad & ", " & _
           "nCantidad = " & nCantidad & ", " & _
           "nImpuesto1 = " & nImpuesto1 * nCantidad & ", " & _
           "nImpuesto2 = " & nImpuesto2 * nCantidad & ", " & _
           "nImpuesto3 = " & nImpuesto3 * nCantidad & " " & _
           "where tItem = '" & sItem & "' and tCodigoPedido ='" & Pedido & "'"
           Cn.Execute Isql

    CambiaPrecioCombo = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and isnull(tfacturado,'')='' and tCodigoPedido='" & Pedido & "'", Cn), "#,###,##0.00")
End Function
'-------------- CENTRAL DELIVERY
Private Sub TimerDelivery_Timer()
 
    
    'Buscar los pedidos en la tabal local
    
    If CD Then
        'Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=1 "
        Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=1  and tCodigoPedidoCD not in ( SELECT tCodigoPedido  FROM PEDIDO_PAGOEFECTIVO where lEstado = 0 )"
        Set RsPedido = Lib.OpenRecordset(Isql, Cn)
        If RsPedido.RecordCount > 0 Then
            Me.cmdDelivery.Caption = "(" + CStr(RsPedido.RecordCount) + ") Delivery"
            Me.cmdDelivery.Visible = True
            Me.TimerColor.Enabled = True
            Me.cmdOpcion(10).Visible = False
        Else
            Me.cmdDelivery.Visible = False
            Me.TimerColor.Enabled = False
            Me.cmdOpcion(10).Visible = True
        End If
    End If
    
    If WEBAP Then
        ' GCAA 29102020
               
        Dim vcanal1 As String
        Dim vcanal2 As String
        Dim vcanal3 As String
        Dim vcanal4 As String
        Dim vcanal5 As String
        Dim val As Integer
        
        Isql = "select tCanalventa1, tCanalventa2, tCanalventa3, tCanalventa4, tCanalventa5 from TCAJACANALVENTA Where  tCaja = '" + sCaja + "' "
        Set RsListaAsignada = Lib.OpenRecordset(Isql, Cn)
        
        val = 0
        If Not (RsListaAsignada.EOF Or RsListaAsignada.BOF) Then
            If RsListaAsignada!tCanalventa1 = True Then
                vcanal1 = "01"
                val = val + 1
            Else
                vcanal1 = ""
                val = val + 0
            End If
            
            If RsListaAsignada!tCanalventa2 = True Then
                vcanal2 = "02"
                val = val + 1
            Else
                vcanal2 = ""
                val = val + 0
            End If
            
            If RsListaAsignada!tCanalventa3 = True Then
                vcanal3 = "03"
                val = val + 1
            Else
                vcanal3 = ""
                val = val + 0
            End If
            
            If RsListaAsignada!tCanalventa4 = True Then
                vcanal4 = "04"
                val = val + 1
            Else
                vcanal4 = ""
                val = val + 0
            End If
            
            If RsListaAsignada!tCanalventa5 = True Then
                vcanal5 = "05"
                val = val + 1
            Else
                vcanal5 = ""
                val = val + 0
            End If
        End If
        
        If val = 0 Then
            Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=2 and testadopedido<>'03'"
        Else
            Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=2 and testadopedido<>'03'and  (tTipoPedido = '" + vcanal1 + "' or tTipoPedido = '" + vcanal2 + "' or tTipoPedido = '" + vcanal3 + "'  or tTipoPedido = '" + vcanal4 + "'  or tTipoPedido = '" + vcanal5 + "')"
        End If
        
        
        Set RsPedido = Lib.OpenRecordset(Isql, Cn)
        If RsPedido.RecordCount > 0 Then
            Me.cmdWebapp.Caption = "(" + CStr(RsPedido.RecordCount) + ") WebApp"
            Me.cmdWebapp.Visible = True
            Me.TimerColorWAP.Enabled = True
            Me.cmdOpcion(15).Visible = False
            If lAlertaConsola Then
            Me.bocina.url = App.path & "\bmps\sounds\pedido.wav"
            End If
        Else
            Me.cmdWebapp.Visible = False
            Me.TimerColorWAP.Enabled = False
            Me.cmdOpcion(15).Visible = True
        End If
    End If
    
    MMControl1.Command = "stop"
'    MMControl1.FileName = App.Path & "\bmps\patio.wav"
'    MMControl1.Command = "open"
'    MMControl1.Command = "play"
    
    If lIntPatio Then
'        Me.bocina.url = App.Path & "\bmps\patio.mp3"
        Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=3 and testadopedido<>'03' and ttipopedido in( '02','01') "
        Set RsPedido = Lib.OpenRecordset(Isql, Cn)
        If RsPedido.RecordCount > 0 Then
            Me.cmdOpcion(10).Caption = "(" + CStr(RsPedido.RecordCount) + ") Patio de Comida"
            'If Me.cmdOpcion(10).backColor = &H8000000F Then
                Me.cmdOpcion(10).backColor = &H808000
                 MMControl1.FileName = App.path & "\bmps\patio.wav"
                 MMControl1.Command = "open"
                 MMControl1.Command = "play"
            'Else
                'Me.cmdOpcion(10).backColor = &H8000000F
            'End If
        Else
            Me.cmdOpcion(10).backColor = &H8000000F
            Me.cmdOpcion(10).Caption = "Refrescar"
            MMControl1.Command = "stop"
        End If
        
'        Call comandaAuto
        
    End If
    
    'Me.cmdOpcion(10).backColor = &H8000000F
    
    If (Calcular("select isnull(lmotoractivo,0) as codigo from tparametro", Cn)) Then
        Me.lblestadomotor.ForeColor = vbBlue
        Me.lblestadomotor.Caption = "Motor Activo"
    Else
        Me.lblestadomotor.ForeColor = vbRed
        Me.lblestadomotor.Caption = "Motor Inactivo"
    End If
    
    If lIntPagoEfectivo Then
        Dim PED  As String
        PED = Calcular("select top 1 tCodigoPedido as codigo from vPedidoDetalle where tDocumento in ( select dp.tDocumento from DPAGODOCUMENTO DP inner join MDOCUMENTO MD on DP.tDocumento=MD.tDocumento  where tCodigoIntegracion=7 and DP.fregistro>= dateadd(day,-2,GETDATE()) and MD.tCaja='" & sCaja & "') and isnull(nEnvio,0)=0   group by tCodigoPedido ", Cn)
        If (PED <> "" And PED <> "0") Then
'            Me.bocina.url = App.Path & "\bmps\PagoEfectivo.mp3"
            Me.frmPedidoPagado.Visible = True
            Me.lblPedidoPagado.Caption = PED
        Else
            Me.frmPedidoPagado.Visible = False
        End If
    End If
    
    If MESA247 And cmdOpcion(1).Visible = False Then
        Isql = "select tCodigoPedido from Mpedido Where tTurno = '' AND tCaja = '' AND isnull(idPedidoMesa,'')<>'' "
        Set RsPedido = Lib.OpenRecordset(Isql, Cn)
        If RsPedido.RecordCount > 0 Then
            Me.cmdMesa247.Caption = "(" + CStr(RsPedido.RecordCount) + ") Mesa247"
            Me.cmdMesa247.Visible = True
            Me.TimerColorWAP.Enabled = True
            Me.cmdOpcion(10).Visible = False
        Else
            Me.cmdMesa247.Visible = False
            Me.TimerColorWAP.Enabled = False
            Me.cmdOpcion(10).Visible = True
        End If
        
        '---- impresion ---- automatico de DOcumento --------
        Dim RsImpresionAut As Recordset
        Dim j As Integer
         Isql = "exec usp_WS_MESA_Inforest '','',3"
        Set RsImpresionAut = Lib.OpenRecordset(Isql, Cn)
        If RsImpresionAut.RecordCount > 0 Then
            If Not (RsImpresionAut.EOF Or RsImpresionAut.BOF) Then
                RsImpresionAut.MoveFirst
                For j = 0 To RsImpresionAut.RecordCount - 1
'                    ImprimeComprobantePagoMesa247 RsImpresionAut.Fields(0)
'                    ReimpresionAut RsImpresionAut.Fields(0)
'                    Isql = "update mdocumento set lImpresionAut = 1 where tdocumento='" & RsImpresionAut.Fields(0) & "'"
'                    Cn.Execute Isql
                RsImpresionAut.MoveNext
                Next j
            End If
          'Call cmdOpcion_Click(10)
        End If
    End If
    
End Sub

Private Sub TimerColor_Timer()
    If Me.cmdDelivery.backColor = &H8000000F Then
        Me.cmdDelivery.backColor = &HFF&
    Else
        Me.cmdDelivery.backColor = &H8000000F
    End If
End Sub


Private Sub ActivarDesactivarCD(ByVal Opcion As Boolean)
    Me.TimerDelivery.Enabled = Opcion
    
    If CD Then
        Me.TimerColor.Enabled = Opcion
        If Opcion = False Then
            Me.cmdDelivery.Visible = False
        End If
    End If
    
    If WEBAP Then
        Me.TimerColorWAP.Enabled = Opcion
        If Opcion = False Then
            Me.cmdWebapp.Visible = False
        End If
    End If
    
End Sub

Private Sub ModifcarEstadoDeliveryCabecera(CodigoPedido As String, tEstadoDelivery As String, tNombre As String)
On Error GoTo ErrorHandler
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
    CnCD.Execute "usp_CD_Modificar_EstadoDelivery_Cabecera '" & tEstadoDelivery & "', '" & CodigoPedido & "', '" & tNombre & "'"
    CnCD.Close
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub

'Oscar Ortega - Central Delivery''''''''''''''''''''''''''''''''''''
Private Sub cmdDelivery_Click()
    Dim delmax As Integer
    Dim cantPed As Integer
    cantPed = Trim(LeerIni(App.path + "\INFOREST.INI", "CentralDelivery", "cantPed", "3"))
    delmax = 0
    Isql = "select tCodigoPedido, tCodigoPedidoCD from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=1"
    Set RsPedido = Lib.OpenRecordset(Isql, Cn)
    If RsPedido.RecordCount > 0 Then
        Dim response As Integer
        Do
            If delmax < cantPed Then
                Isql = "Update Mpedido Set tTurno = '" + sTurno + "', TCaja = '" + sCaja + "', tTurnoAnterior = '" + sTurno + "', tCajaAnterior = '" + sCaja + "', tSalon = '01' Where tCodigoPedido = '" + RsPedido!tCodigoPedido + "'"
                Cn.Execute Isql
                If Not IsNull(RsPedido!tCodigoPedidoCD) Then
                    Call ModifcarEstadoDeliveryCabecera(RsPedido!tCodigoPedidoCD, "2", "")
                End If
            End If
            delmax = delmax + 1
            RsPedido.MoveNext
        Loop While Not RsPedido.EOF
    End If
    Me.cmdDelivery.Visible = False
    RsCabecera.Requery
    RsDetalle.Requery
    RsCombo.Requery
    RsComboPropiedad.Requery
    RsProductoPropiedad.Requery
    Me.cmdOpcion(10).Visible = True
End Sub

Public Sub EliminarPedidoCD(ByVal tCodigoPedido As String)
On Error GoTo ErrorHandler
    Dim CnCD As Connection
    Set CnCD = New Connection
    CnCD.Provider = "SQLOLEDB"
    CnCD.CursorLocation = adUseServer
    CnCD.ConnectionString = "User ID=" & sUserName & _
                         ";password=" & sUserPassword & _
                         ";Data Source=" & sRutaCD & _
                         ";Initial Catalog=" & sMDBCD
    CnCD.CommandTimeout = 250
    CnCD.Open
    CnCD.Execute "usp_CD_Modificar_EstadoDelivery_Cabecera 6, '" + tCodigoPedido + "', ''" 'pp
    CnCD.Close
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub

'----------------- WEB APP
Private Sub TimerColorWAP_Timer()
    If Me.cmdWebapp.backColor = &H8000000F Then
        Me.cmdWebapp.backColor = &HFF&
        Me.cmdMesa247.backColor = &HFF00&
    Else
        Me.cmdWebapp.backColor = &H8000000F
        Me.cmdMesa247.backColor = &H8000000F
    End If
End Sub



Private Sub ModificarEstadoPedidoWebAp(CodigoPedido As String, tEstadoPedido As String)
On Error GoTo ErrorHandler
'    Dim CnWAP As Connection
'    'Configuración
'    Set CnWAP = New Connection
'    CnWAP.Provider = "SQLOLEDB"
'    CnWAP.CursorLocation = adUseServer
'    CnWAP.ConnectionString = "User ID=" & sUserName & _
'                            ";password=" & sUserPassword & _
'                            ";Data Source=" & sRutaWAP & _
'                            ";Initial Catalog=" & sMDBWAP
'    CnWAP.CommandTimeout = 250
'    CnWAP.Open
'    CnWAP.Execute "APP_ACTUALIZAR_MPEDIDO_ESTADO '" & CodigoPedido & "','" & tEstadoPedido & "'"
'    CnWAP.Close
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub

Private Sub cmdWebapp_Click()

    MMControl1.Command = "stop"
    Dim delmax As Integer
    Dim cantPed As Integer
    cantPed = Trim(LeerIni(App.path + "\INFOREST.INI", "WebMobile", "cantPed", "3"))
    delmax = 0
    
    ' GCAA 29102020
    Dim vcanal1 As String
    Dim vcanal2 As String
    Dim vcanal3 As String
    Dim vcanal4 As String
    Dim vcanal5 As String
    Dim val As Integer
    
    Isql = "select tCanalventa1, tCanalventa2, tCanalventa3, tCanalventa4, tCanalventa5 from TCAJACANALVENTA Where  tCaja = '" + sCaja + "'"
    Set RsListaAsignada = Lib.OpenRecordset(Isql, Cn)
        
     val = 0
             If RsListaAsignada!tCanalventa1 = True Then
            vcanal1 = "01"
            val = val + 1
        Else
            vcanal1 = ""
            val = val + 0
        End If
        
        If RsListaAsignada!tCanalventa2 = True Then
            vcanal2 = "02"
            val = val + 1
        Else
            vcanal2 = ""
            val = val + 0
        End If
        
        If RsListaAsignada!tCanalventa3 = True Then
            vcanal3 = "03"
            val = val + 1
        Else
            vcanal3 = ""
            val = val + 0
        End If
        
        If RsListaAsignada!tCanalventa4 = True Then
            vcanal4 = "04"
            val = val + 1
        Else
            vcanal4 = ""
            val = val + 0
        End If
        
        If RsListaAsignada!tCanalventa5 = True Then
            vcanal5 = "05"
            val = val + 1
        Else
            vcanal5 = ""
            val = val + 0
        End If
    

    If val = 0 Then
        Isql = "select tCodigoPedido, tCodigoPedidoWebAp from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=2 and testadopedido<>'03'"
    Else
        Isql = "select tCodigoPedido, tCodigoPedidoWebAp from Mpedido Where tTurno = '' AND tCaja = '' AND tTiporecepcion=2 and testadopedido<>'03' AND  (tTipoPedido = '" + vcanal1 + "' or tTipoPedido = '" + vcanal2 + "' or tTipoPedido = '" + vcanal3 + "' or tTipoPedido = '" + vcanal4 + "'  or tTipoPedido = '" + vcanal5 + "')"
    End If
    
    Set RsPedido = Lib.OpenRecordset(Isql, Cn)
    If RsPedido.RecordCount > 0 Then
        Dim response As Integer
        Do
            If delmax < cantPed Then
                Isql = "Update Mpedido Set tusuario='" & sUsuario & "', tTurno = '" + sTurno + "', TCaja = '" + sCaja + "', tTurnoAnterior = '" + sTurno + "', tCajaAnterior = '" + sCaja + "', tSalon = '01' Where tCodigoPedido = '" + RsPedido!tCodigoPedido + "'"
                Cn.Execute Isql
                If Not IsNull(RsPedido!tCodigoPedidoWebAp) Then
                    Call ModificarEstadoPedidoWebAp(RsPedido!tCodigoPedidoWebAp, "6")
                End If
            End If
            delmax = delmax + 1
            RsPedido.MoveNext
        Loop While Not RsPedido.EOF
    End If
    
    
    Me.cmdWebapp.Visible = False
    RsCabecera.Requery
    RsDetalle.Requery
    RsCombo.Requery
    RsComboPropiedad.Requery
    RsProductoPropiedad.Requery
    Me.cmdOpcion(15).Visible = True
End Sub
'----------------------------------------------------






'Paolo Quintero - Habilita y Desabilita el Timer del Boton Delivery
Private Sub HabilitaTimerColor(ByVal enable As Boolean)
    'If CD Then
    '    TimerDelivery.Enabled = enable
    '    TimerColor.Enabled = enable
    'End If
End Sub


'Oscar Ortega---------------------------------
Private Function ObligaPropiedad(ByVal tProducto As String) As Boolean
    Dim j As Integer
    Dim i As Integer
    Dim RstProductoPropiedad As Recordset
    Dim flag As Boolean
    Dim oPos As Integer
    oPos = RsOperador.AbsolutePosition
    Dim mensajeOperador As String
    flag = True 'Si permite salir
    ObligaPropiedad = True
    If RsOperador.RecordCount > 0 Then
        If RsOperador.EOF Then
            RsOperador.MoveFirst
            For i = 1 To 13
                cmdOperador(i).backColor = vbButtonFace
            Next i
        End If
        RsOperador.MoveFirst
        While RsOperador.EOF = False
            If RsOperador!lObligaPropiedad = True Then
                If wAgregaCombo Then
                    Isql = "Select * From tComboPropiedad Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tCodigoPedido = '" & sPedido & "' And tItem = '" & sItem & "' And tItemCombo = '" & xItem & "'"
                Else
                    Isql = "Select * From tProductoPropiedad Where tCodigoPropiedad In (Select tCodigoPropiedad from TPropiedad Where tOperador = '" & RsOperador!codigo & "' And tProducto = '" & tProducto & "' ) And tCodigoPedido = '" & sPedido & "' And tItem = '" & sItem & "' "
                End If
                Set RstProductoPropiedad = Lib.OpenRecordset(Isql, Cn)
                If RstProductoPropiedad.RecordCount = 0 Then
                    flag = False 'Esta Obligado y no ha elegido Propiedad
                    mensajeOperador = mensajeOperador + "(" + RsOperador!descripcion + ")"
                End If
            End If
            RsOperador.MoveNext
        Wend
        If flag = False Then
            MsgBox "Propiedades obligadas " & mensajeOperador, vbExclamation, sMensaje
        End If
        ObligaPropiedad = flag
        RsOperador.AbsolutePosition = oPos
        RsOperador.Find "nboton = " & Trim(str(RsOperador!nBoton))
        nOperadorPropiedad = RsOperador!nControl
        For i = 1 To 13
            cmdOperador(i).backColor = vbButtonFace
        Next i
        If RsOperador!nBoton > 0 Then
            cmdOperador(RsOperador!nBoton).backColor = vbRed
        End If
        'AsignaPropiedad
    End If
End Function

Private Sub ListarOperadoresConFiltro(ByVal tProducto As String)
    'Isql = "select * from vOperador where lActivo = 1 " & _
           "AND ((select Count(tCodigoPropiedad) " & _
           "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
           "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
           "order by Codigo"

Isql = "usp_Inforest_ObtieneOperadoresXProducto '" & tProducto & "'"
 
Set RsOperador = Lib.OpenRecordset("usp_Inforest_ObtieneOperadoresXProducto '" & tProducto & "'", Cn)
    'If RsOperador.RecordCount > 0 Then
    AsignaBoton 13, RsOperador, cmdOperador()
    If RsOperador.RecordCount > 0 Then
        RsOperador.MoveFirst
        Dim i As Integer
        For i = 1 To RsOperador.RecordCount
            If RsOperador!nBoton <> 0 Then
                cmdOperador(RsOperador!nBoton).backColor = vbButtonFace
            End If
            RsOperador.MoveNext
        Next i
        RsOperador.MoveFirst
                xOperador = RsOperador!codigo

        If RsOperador!nBoton <> 0 Then
            cmdOperador(RsOperador!nBoton).backColor = vbRed
        End If
    End If
End Sub

Private Function Obtener_ProductosNoImpresosPorPedido(ByVal tCodigoPedido) As Recordset
    If wAgregaCombo Then
        Isql = "Select CPedido.*,TProducto.tDetallado from CPedido Left Join TProducto On CPedido.tProductoCombo = TProducto.tCodigoProducto " & _
               "Where CPedido.lImprime = '0' And tItem = '" & sItem & "' And CPedido.tCodigoPedido = '" & tCodigoPedido & "' "
    Else
        Isql = "Select DPedido.*,TProducto.tDetallado from DPedido Left Join TProducto On DPedido.tCodigoProducto = TProducto.tCodigoProducto Where DPedido.tCodigoPedido = '" & tCodigoPedido & "' And DPedido.lImprime = '0'"
    End If
    Set Obtener_ProductosNoImpresosPorPedido = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_OperadoresObligatoriosPorProducto(ByVal tProducto) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1" & _
       "AND ((select Count(tCodigoPropiedad) " & _
       "FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
       "Where TPROPIEDAD.lActivo = 1 And IsNull(TOPERADOR.lStockMenos, 0) <> 1 " & _
       "And TPROPIEDAD.tOperador = vOperador.Codigo and tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) " & _
       "order by Codigo"
    Set Obtener_OperadoresObligatoriosPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_PropiedadesSeleccionadasPorProducto(ByVal tCodigoPedido, ByVal tItem, ByVal tProducto, ByVal tOperador) As Recordset
    If wAgregaCombo Then
        Isql = "Select * From TComboPropiedad " & _
            "Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Else
        Isql = "Select * From TProductoPropiedad " & _
            "Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tCodigoPropiedad In ( " & _
            "Select tCodigoPropiedad from TPropiedad " & _
            "Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    End If
    Set Obtener_PropiedadesSeleccionadasPorProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_ProductoDeCombo(ByVal tCombo As String, ByVal tCodigoProducto As String) As Recordset
    Isql = "Select lUnico,tEtiqueta From TCOMBO Where tCombo = '" & tCombo & "' And tCodigoProducto = '" & tCodigoProducto & "'"
    Set Obtener_ProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function ObtenerSumaCantidadesEnElCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tEtiqueta As String) As Double
    Isql = "Select ISNULL(Sum(nCantidad),0) as nCantidad from CPEDIDO " & _
           "Where   tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tProductoCombo IN ( " & _
           "Select tCodigoProducto From TCOMBO Where tCombo = CPEDIDO.tProducto And tEtiqueta = '" & tEtiqueta & "' And lUnico = '1') "
    Dim Suma As Double
    Dim oRsResultado As Recordset
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    Suma = oRsResultado!nCantidad
    ObtenerSumaCantidadesEnElCombo = Suma
End Function

Private Function ObtenerSumaCantidadesEnElComboSinEsteProducto(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tItemCombo As String, ByVal tEtiqueta As String) As Double
    Isql = "Select ISNULL(Sum(nCantidad),0) as nCantidad from CPEDIDO " & _
           "Where   tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tItemCombo <> '" & tItemCombo & "' And tProductoCombo IN ( " & _
           "Select tCodigoProducto From TCOMBO Where tCombo = CPEDIDO.tProducto And tEtiqueta = '" & tEtiqueta & "' And lUnico = '1') "
    Dim Suma As Double
    Dim oRsResultado As Recordset
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    Suma = oRsResultado!nCantidad
    ObtenerSumaCantidadesEnElComboSinEsteProducto = Suma
End Function

Private Function ObtenerDetalleProducto(ByVal tCodigoPedido As String, ByVal tItem As String) As Recordset
    'oo
    Isql = " Select D.* , " & _
           " case when ( " & _
           " CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa " & _
           " when '02' then p.npreciodelivery " & _
           "            when '03' then p.npreciollevar " & _
           "            when '04' then p.npreciocanal4 " & _
           "            when '05' then p.npreciocanal5 " & _
           " END)=0 then p.nprecioventa else " & _
           " ( CASE (D.tTipoPedido)  WHEN '01' THEN P.nprecioventa " & _
           "            when '02' then p.npreciodelivery " & _
           "            when '03' then p.npreciollevar " & _
           "            when '04' then p.npreciocanal4 " & _
           "            when '05' then p.npreciocanal5 End) end as 'PrecioProducto',P.tMoneda as 'tMonedaProducto' from DPedido As D left Join TProducto As P On D.tCodigoProducto = P.tCodigoProducto where D.tCodigoPedido='" & tCodigoPedido & "' and D.tItem ='" & tItem & "'"
    Set ObtenerDetalleProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_CantidadMaximaDeUnicoEtiqueta(ByVal tCodigoPedido As String, ByVal tItem As String, cantidadActual As Double) As Double
    Isql = "Select Sum(P.nCantidad) as Cantidad " & _
           "from CPedido as P Left Join TCOMBO as C On P.tProductoCombo = C.tCodigoProducto " & _
           "where P.tCodigoPedido='" & tCodigoPedido & "' and P.tItem ='" & tItem & "' And C.lUnico = '1' " & _
           "Group By C.tEtiqueta"
     '06/02
    Isql = "Select Sum(P.nCantidad) as Cantidad " & _
           "from CPedido as P Left Join TCOMBO as C ON P.tProducto = C.tCombo AND P.tProductoCombo = C.tCodigoProducto " & _
           "where P.tCodigoPedido='" & tCodigoPedido & "' and P.tItem ='" & tItem & "' And C.lUnico = '1' " & _
           "Group By C.tEtiqueta,c.tcodigoproducto  order by 1 desc"
    Dim oRsResultado As Recordset
    Dim oi As Integer
    Dim CantMax As Double
    CantMax = 0
    CantMax = cantidadActual
    Set oRsResultado = Lib.OpenRecordset(Isql, Cn)
    If oRsResultado.RecordCount > 0 Then
        For oi = 1 To oRsResultado.RecordCount
            If oRsResultado!Cantidad <> CantMax Then
                CantMax = oRsResultado!Cantidad
            End If
            oRsResultado.MoveNext
        Next oi
        Obtener_CantidadMaximaDeUnicoEtiqueta = CantMax
    Else
        Obtener_CantidadMaximaDeUnicoEtiqueta = 0
    End If
End Function

Private Function Obtener_PropiedadesPendientesPorItems(ByVal tProducto As String, ByVal tItem As String) As Boolean
    Dim oj As Integer
    Dim oflag As Boolean
    oflag = True
    Dim oRsOperadoresObligados As Recordset ' Lista de Operadores obligados de un producto
    Set oRsOperadoresObligados = Obtener_OperadoresObligatoriosPorProducto(tProducto)
    For oj = 1 To oRsOperadoresObligados.RecordCount
        Dim oRsPropiedadesDeOperador As Recordset ' Lista de Propiedades por Operador
        Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProducto(sPedido, tItem, tProducto, oRsOperadoresObligados!codigo)
        If oRsPropiedadesDeOperador.RecordCount = 0 Then
            oflag = False
           oj = oRsOperadoresObligados.RecordCount
        End If
        oRsOperadoresObligados.MoveNext
    Next oj
    Obtener_PropiedadesPendientesPorItems = oflag
End Function

Private Function ExistenPropiedadesPendientesEnCombos(ByVal tCodigoPedido As String) As Boolean
    Dim oRsCombosPedido As Recordset
    Dim oi, oj, ok As Integer
    Dim oMensaje As String
    oMensaje = "Combo(s) con productos con propiedades obligatorias: "
    Set oRsCombosPedido = Obtener_TodosLosCombosDelPedido(tCodigoPedido)
    Dim oflag As Boolean
    oflag = True
    For oi = 1 To oRsCombosPedido.RecordCount
        Dim oRsProductoNoImpCombo As Recordset
        Set oRsProductoNoImpCombo = Obtener_LosProductosNoImpDelCombo(tCodigoPedido, oRsCombosPedido!tItem)
        For oj = 1 To oRsProductoNoImpCombo.RecordCount
            Dim oRsOperadores As Recordset
            Set oRsOperadores = Obtener_OperadoresObligatoriosDeUnProducto(oRsProductoNoImpCombo!tProductoCombo)
            For ok = 1 To oRsOperadores.RecordCount
                Dim oRsPropiedadesDeOperador As Recordset
                Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProductoDeCombo(tCodigoPedido, oRsCombosPedido!tItem, oRsProductoNoImpCombo!tItemCombo, oRsProductoNoImpCombo!tProductoCombo, oRsOperadores!codigo)
                If oRsPropiedadesDeOperador.RecordCount = 0 Then
                    oflag = False
                    oMensaje = oMensaje + "(" + oRsCombosPedido!tDetallado + ")"
                    oj = oRsProductoNoImpCombo.RecordCount
                    ok = oRsOperadores.RecordCount
                End If
                oRsOperadores.MoveNext
            Next ok
            oRsProductoNoImpCombo.MoveNext
        Next oj
        oRsCombosPedido.MoveNext
    Next oi
    
    If (oflag = False) Then
        MsgBox (oMensaje)
    End If
    ExistenPropiedadesPendientesEnCombos = oflag
End Function

Private Function Obtener_TodosLosCombosDelPedido(ByVal tCodigoPedido As String) As Recordset
    Isql = "Select  DPedido.*,TProducto.tDetallado From DPedido Left Join TProducto On DPedido.tCodigoProducto = TProducto.tCodigoProducto " & _
           "Where   DPedido.tCodigoPedido = '" & tCodigoPedido & "' And DPedido.lCombinacion = '1'"
    Set Obtener_TodosLosCombosDelPedido = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_LosProductosNoImpDelCombo(ByVal tCodigoPedido As String, ByVal tItem As String) As Recordset
    Isql = "Select * from CPedido Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And lImprime = '0'"
    Set Obtener_LosProductosNoImpDelCombo = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_OperadoresObligatoriosDeUnProducto(ByVal tProducto As String) As Recordset
    Isql = "select * from vOperador where lActivo = 1 AND lObligaPropiedad = 1 AND ( " & _
           "(select Count(tCodigoPropiedad) FROM dbo.TPROPIEDAD LEFT OUTER JOIN dbo.TOPERADOR ON dbo.TPROPIEDAD.tOperador = dbo.TOPERADOR.tOperador " & _
           "Where TPROPIEDAD.lActivo = 1 And IsNull(tOperador.lStockMenos, 0) <> 1 And TPROPIEDAD.tOperador = vOperador.Codigo " & _
           "And tProducto='" & tProducto & "') > 0 OR lStockMenos > 0 ) order by Codigo"
    Set Obtener_OperadoresObligatoriosDeUnProducto = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function Obtener_PropiedadesSeleccionadasPorProductoDeCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tItemCombo As String, ByVal tProducto As String, ByVal tOperador As String) As Recordset
    Isql = "Select * From TComboPropiedad " & _
           "Where tCodigoPedido = '" & tCodigoPedido & "' And tItem = '" & tItem & "' And tItemCombo = '" & tItemCombo & "' And tCodigoPropiedad  In ( " & _
           "Select tCodigoPropiedad from TPropiedad Where tProducto = '" & tProducto & "' And tOperador = '" & tOperador & "') "
    Set Obtener_PropiedadesSeleccionadasPorProductoDeCombo = Lib.OpenRecordset(Isql, Cn)
End Function

Private Function ExistenPropiedadesPendientesEnPedido(ByVal tCodigoPedido As String) As Boolean
    Dim oRsDPedidoNoImp As Recordset 'Lista de Productos no impresos
    Set oRsDPedidoNoImp = Obtener_ProductosNoImpresosPorPedido(sPedido)
    Dim oi As Integer
    Dim oj As Integer
    Dim oflag As Boolean
    Dim oMensaje As String
    oMensaje = "Item(s) con obligatoriedad de propiedad: "
    oflag = True
    'Para cada`producto de DPedido cual lImprime = '0'
    For oi = 1 To oRsDPedidoNoImp.RecordCount
    'Obtener Operadores Obligatorios Filtrados
        Dim oRsOperadoresObligados As Recordset ' Lista de Operadores obligados de un producto
        Set oRsOperadoresObligados = Obtener_OperadoresObligatoriosPorProducto(oRsDPedidoNoImp!tCodigoProducto)
        'Para cada operador Obtener la lista de propiedades
        For oj = 1 To oRsOperadoresObligados.RecordCount
            Dim oRsPropiedadesDeOperador As Recordset ' Lista de Propiedades por Operador
            Set oRsPropiedadesDeOperador = Obtener_PropiedadesSeleccionadasPorProducto(sPedido, oRsDPedidoNoImp!tItem, oRsDPedidoNoImp!tCodigoProducto, oRsOperadoresObligados!codigo)
            'Verificar si en TProductoPropiedad existe para el item y tCodigoPropiedad IN (lista de propiedades)
            If oRsPropiedadesDeOperador.RecordCount = 0 Then
                oflag = False
                oMensaje = oMensaje + "(" + oRsDPedidoNoImp!tDetallado + ")"
                oj = oRsOperadoresObligados.RecordCount
            End If
            oRsOperadoresObligados.MoveNext
        Next oj
        oRsDPedidoNoImp.MoveNext
    Next oi
    If oflag = False Then
        MsgBox (oMensaje)
    End If
    ExistenPropiedadesPendientesEnPedido = oflag
End Function

Private Function ObtenerCodigoMozo(ByVal tResumido As String) As String
    Isql = "select * from vMozo where substring(Codigo,1,1)<>'*' AND lActivo = 1 Order by nBoton"
    Set RsMozo = Lib.OpenRecordset(Isql, Cn)
    RsMozo.Filter = "tResumido = '" & tResumido & "'"
    If RsMozo.RecordCount = 0 Then
        txtMozo.Caption = "Sin" & Mesero
        ObtenerCodigoMozo = "0000"
    Else
        txtMozo.Caption = tResumido
        ObtenerCodigoMozo = RsMozo!codigo
    End If
End Function


Private Function ObtenerClienteFacturaCD(ByVal tCodigoDelivery As String, ByVal tCodigoPedido As String) As String
On Error GoTo ErrorHandler
    Dim RsCodigoCliente As Recordset
    Isql = "USP_RD_OBTENER_CLIENTEFACTURA '" & tCodigoDelivery & "', '" & tCodigoPedido & "'"
    Set RsCodigoCliente = Lib.OpenRecordset(Isql, Cn)
    ObtenerClienteFacturaCD = RsCodigoCliente!tCodigoCliente
    Exit Function
ErrorHandler:
    MsgBox (Err.Description & "Obtener Cliente Factura")
    ObtenerClienteFacturaCD = ""
End Function

Private Function ObtenerClienteFacturaWebAp(ByVal tCodigoPedido As String) As String
On Error GoTo ErrorHandler
    Dim RsCodigoCliente As Recordset
    'Isql = "USP_APP_OBTENER_CLIENTE_FACTURA " & tCodigoWebAp & ", '" & tCodigoPedido & "'"
    Isql = "exec usp_WS_Inforest '" + tCodigoPedido + "',4" '
    Set RsCodigoCliente = Lib.OpenRecordset(Isql, Cn)
    ObtenerClienteFacturaWebAp = RsCodigoCliente!tClienteCorp
    Exit Function
ErrorHandler:
    MsgBox (Err.Description & "Obtener Cliente Factura")
    ObtenerClienteFacturaWebAp = ""
End Function


Public Sub verificatitulo()
    Dim rsInsumo As New ADODB.Recordset
    If Calcular("select isnull(lControlInsumoCritico,0) as codigo from tproducto  INNER JOIN Dpedido on tproducto.tcodigoproducto=dpedido.tcodigoproducto where dpedido.tcodigopedido='" & sPedido & "' and titem='" & sItem & "'", Cn) = True Then
                    Set rsInsumo = Lib.OpenRecordset("select isnull(tcodigoinsumo,'') tcodigoinsumo , isnull(tinsumo.descripcion,'') ,isnull(nstock,0) , dpedido.ncantidad from tproducto inner join tinsumo on tproducto.tcodigoinsumo =tinsumo.tcodigo inner join dpedido on tproducto.tcodigoproducto=dpedido.tcodigoproducto  where    dpedido.tcodigopedido='" & sPedido & "' and titem='" & sItem & "' and tinsumo.lactivo=1", Cn)
                    If Not (rsInsumo.EOF Or rsInsumo.BOF) Then
                            Me.Caption = muestra & "   Insumo Crítico --->   " & rsInsumo.Fields(1) & " =  Stock: " & str(rsInsumo.Fields(2)) & " / Solicitado: " + str(rsInsumo.Fields(3))
                    End If
            Else
                    Me.Caption = muestra
    End If
End Sub

'diaContable
Public Function obtieneDiaContable() As Date
   Dim DiaContable As Date
   Dim oComando As New clsComando
  
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
                'd = oComando.GetParameterValue("@fDiaContable")
End Function
'diaContable

Private Sub CalculaAplicaTope(nTope As Double)
    Dim sCriterio As String
    Dim lAcumulable As Boolean
    Dim nOferta As Double
    Dim nSuma As Double
    
    nSuma = Calcular("SELECT sum(nPrecioOficial*nCantidad) as Codigo FROM dbo.DPEDIDO LEFT OUTER JOIN dbo.TPRODUCTO ON dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTO.tCodigoProducto where tCodigoPedido='" & sPedido & "' and lDescuento=1", Cn)

If RsDetalle.RecordCount <> 0 Then

   RsDetalle.MoveFirst
   
   Do While Not RsDetalle.EOF
      'Busca Oferta
      nPVenta = 0
      sCriterio = "tCodigoProducto ='" & RsDetalle!tCodigoProducto & "' and lActivo=1"
      sCriterio = sCriterio & " and (tFrecuencia='00' or tFrecuencia='0" & Weekday(FechaServidor(), vbMonday) & "' or (tFrecuencia='99' and fFecha='" & Format(FechaServidor(), "yyyy/MM/dd 00:00") & "') and tHoraInicial<='" & Format(Time, "HH:mm") & "' and tHoraFinal>='" & Format(Time, "HH:mm") & "')"
      sCriterio = sCriterio & " and (lPermanente=1 or (lPermanente=0 and fFechaInicial<='" & Format(FechaServidor(), "yyyy/mm/dd") & "' and fFechaFinal>='" & Format(FechaServidor(), "yyyy/mm/dd") & "'))"
        
      Isql = "select * from TOFERTA where " & sCriterio
      Set RsOferta = Lib.OpenRecordset(Isql, Cn)
      
      lAcumulable = True
      nOferta = 0
      Acumulado = 0
      
      If RsOferta.RecordCount > 0 Then
         RsOferta.MoveFirst
         lAcumulable = RsOferta!lAcumulable
         nOferta = RsDetalle!nPrecioOficial * IIf(IsNull(RsOferta!nRatio), 1, RsOferta!nRatio) / 100
      End If
      
      If RsDetalle!lDescuento And lAcumulable = True Then
         
         xDescuento = (RsDetalle!nPrecioOficial - nOferta) * (RsDetalle!nCantidad * 100 / nSuma)
         nPVenta = (RsDetalle!nPrecioOficial - nOferta) - ((nTope * xDescuento / 100) / RsDetalle!nCantidad)
         
          Select Case pais ' ok
            Case "001" 'Bolivia
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
                         
            Case Else 'Peru, Ecuador
                         Acumulado = IIf(RsDetalle!nprecioImpuesto1 <> 0, Acumulado + nPorcentaje1, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto2 <> 0, Acumulado + nPorcentaje2, Acumulado)
                         Acumulado = IIf(RsDetalle!nprecioImpuesto3 <> 0, Acumulado + nPorcentaje3, Acumulado)
                         Acumulado = 1 + (Acumulado / 100)
                        
                         nImpuesto1 = IIf(RsDetalle!nprecioImpuesto1 <> 0, nPVenta / Acumulado * nPorcentaje1 / 100, 0)
                         nImpuesto2 = IIf(RsDetalle!nprecioImpuesto2 <> 0, nPVenta / Acumulado * nPorcentaje2 / 100, 0)
                         nImpuesto3 = IIf(RsDetalle!nprecioImpuesto3 <> 0, nPVenta / Acumulado * nPorcentaje3 / 100, 0)
                         nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
          
          End Select
         
         Isql = "Update DPEDIDO Set nPrecioNeto = " & nPBase & ", " & _
                "nDescuento = " & RsDetalle!nPrecioOficial - nPVenta & ", " & _
                "nRecargo = " & nRecargo & ", " & _
                "nPrecioOficial = " & RsDetalle!nPrecioOficial & ", " & _
                "nprecioImpuesto1 = " & nImpuesto1 & ", " & _
                "nprecioImpuesto2 = " & nImpuesto2 & ", " & _
                "nprecioImpuesto3 = " & nImpuesto3 & ", " & _
                "nPrecioVenta = " & nPVenta & ", " & _
                "nventa = " & nPVenta * RsDetalle!nCantidad & ", " & _
                "nCantidad = " & RsDetalle!nCantidad & ", " & _
                "nImpuesto1 = " & nImpuesto1 * RsDetalle!nCantidad & ", " & _
                "nImpuesto2 = " & nImpuesto2 * RsDetalle!nCantidad & ", " & _
                "nImpuesto3 = " & nImpuesto3 * RsDetalle!nCantidad & ", " & _
                "tCortesia = '" & sCortesia & "' " & _
                "where tItem = '" & RsDetalle!tItem & "' and tCodigoPedido ='" & sPedido & "'"
                Cn.Execute Isql
      End If
   RsDetalle.MoveNext
   Loop
End If

        TxtDescuento.Caption = Format(Calcular("select sum(nDescuento*nCantidad) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
        'txtSuma.Caption = Format(Calcular("select sum(nVenta) as Codigo FROM DPEDIDO where tEstadoItem = 'N' and isnull(tfacturado,'')='' and tCodigoPedido='" & sPedido & "'", Cn), "#,###,##0.00")
        SumPrecioVenta (sPedido)
End Sub

Public Sub InicializaVariableCabecera()
    wDetalle = False
    wAgrega = True
    sPedido = ""
    sOrigenVenta = ""
    sHabitacion = ""
    sReserva = ""
    sPuntoVenta = ""
    sPasajero = ""
    sCompania = ""
    sContacto = ""
    sCodigoDescuento = ""
    tAutorizaDescuento = ""
    txtPuntoVenta.Caption = ""
    sMesa = ""
    nTiempoDelivery = 0
    xDescuento = 0
    
    sAutoriza = ""
    txtAdulto.Text = 0
    txtNino.Text = 0
    txtJuntar.Text = 0
    lPrecuenta = False
    
    txtEnvioAntes.Text = 0
    
    'pariente2013
    sCodigoInvitado = ""
    sCodigoParienteSeleccionado = ""
    'pariente2013
    
    Me.txtOrdenCompra.Text = ""
End Sub

Public Sub InicializaVariableCanales()
If wAgrega Then
    txtMozo.Caption = ""
    txtMotorizado.Caption = ""
    TxtDescuento.Caption = "0.00"
    txtCliente.Caption = ""
    txtDireccion.Caption = ""
    txtTelefono.Caption = ""
    txtFecha.Caption = FechaServidor()
    txtFechaProg.Caption = ""
    txtObservacion.Caption = ""
    txtMontoMax.Caption = "0"
    Tienda = ""
    
    sTipoAtencion = "01"
    sMozo = "0000"
    sMotorizado = "0000"
    sCliente = ""
    sTienda = ""
      'pariente2013
    sCodigoInvitado = ""
    sCodigoParienteSeleccionado = ""
      'entregarA
    txtEntregar.Caption = ""
  End If
    'pariente2013
    'Canales de Venta
    RsCanalesVenta.MoveFirst
        
    RsCanalesVenta.Filter = "Codigo = '" & sTipoPedido & "'"
    
    If RsCanalesVenta.RecordCount = 0 Then
        Set RsCanalesVenta = Lib.OpenRecordset("select * from vTipoPedido", Cn)
        MsgBox "No se puede mostrar la información, el canal de venta predeterminado para esta caja esta Inactivo", vbCritical, sMensaje
        cmdOpcion_Click (2)
        Exit Sub
    End If
    
    lActivaMozo = IIf(IsNull(RsCanalesVenta!lActivaMozo), False, RsCanalesVenta!lActivaMozo)
    lActivaMotorizado = IIf(IsNull(RsCanalesVenta!lActivaMotorizado), False, RsCanalesVenta!lActivaMotorizado)
    lCanalDelivery = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    lCanalCentralPedidos = IIf(IsNull(RsCanalesVenta!lCanalCentralPedidos), False, RsCanalesVenta!lCanalCentralPedidos)
    'entregarA
    lObligaEntregarA = IIf(IsNull(RsCanalesVenta!lObligaEntregarA), False, RsCanalesVenta!lObligaEntregarA)
    
    'origen de ventas
     lOrigenVentas = IIf(IsNull(RsCanalesVenta!lCanalDelivery), False, RsCanalesVenta!lCanalDelivery)
    
    'If lMCPV Then
        'lObligaMozo = False
        'lActivaMozo = False
    'Else
        lObligaMozo = IIf(IsNull(RsCanalesVenta!lObligaMozo), False, RsCanalesVenta!lObligaMozo)
    'End If
    
    lObligaMotorizado = IIf(IsNull(RsCanalesVenta!lObligaMotorizado), False, RsCanalesVenta!lObligaMotorizado)
    lObligaMesa = IIf(IsNull(RsCanalesVenta!lObligaMesa), False, RsCanalesVenta!lObligaMesa)
    lObligaPax = IIf(IsNull(RsCanalesVenta!lObligaPax), False, RsCanalesVenta!lObligaPax)
    lObligaFechaEntrega = IIf(IsNull(RsCanalesVenta!lObligaIngresoFechaEntrega), False, RsCanalesVenta!lObligaIngresoFechaEntrega)
    lObligaClienteFrecuente = IIf(IsNull(RsCanalesVenta!lObligaClienteFrecuente), False, RsCanalesVenta!lObligaClienteFrecuente)
    lObligaOrigenVenta = IIf(IsNull(RsCanalesVenta!lObligaOrigenVenta), False, RsCanalesVenta!lObligaOrigenVenta)
    
    
    If lActivaMozo Then
       fraMozo.Visible = True
       Me.fraMozo.Caption = Mesero
    Else
       fraMozo.Visible = False
    End If
    
   'If lActivaMotorizado Then
      ' fraMotorizado.Visible = True
   ' Else
       fraMotorizado.Visible = False
   ' End If
    Call ActivarOrigenVentas
  

End Sub

Function Validaciones(ByVal proceso As String) As Boolean
' Desarrollador: Erick de la cruz Quispe 16/08/2017
On Error GoTo fin

    Dim i As Double
    Dim X As Long
    
    If proceso = "Documento" Then
        'Chequea Consistencia Infhotel
        If lInfhotel Then
            If sPuntoVenta = "" Then
              MsgBox "Falta ingresar el punto de venta", vbExclamation, sMensaje
              Validaciones = False
              Exit Function
            End If
        End If
        
        If lMCPV Then ' hacer seguimiento de la caja aun para evaluar
            If Not ValidaExistenciaProducto() Then
                MsgBox "El Pedido ya fue importado", vbExclamation, sMensaje
                RsCabecera.Requery
                RsDetalle.Requery
                Validaciones = False
                Exit Function
            End If
        End If
        
        'obligatoriedad de enviar platos a producciom
        If lPrinter And lObligaPrinter Then
           i = Calcular("select count(tCodigoPedido) as codigo from DPEDIDO where lImprime=0 and tCodigoPedido ='" & sPedido & "'", Cn)
           If i > 0 Then
              MsgBox "Falta Enviar platos a Producción", vbExclamation, sMensaje
              Validaciones = False
              Exit Function
           End If
        End If
        
        ' obligatoriedad de imprimir precuenta
        If lObligaPrecuenta Then
            i = Calcular("select nPrecuenta as codigo from MPEDIDO where tCodigoPedido ='" & sPedido & "'", Cn)
            If i = 0 Then
              MsgBox "Falta Imprimir la Precuenta", vbExclamation, sMensaje
              Validaciones = False
              Exit Function
            End If
        End If
        
        ' Validacion descuento negativo
        i = Calcular("select sum(ndescuento) as codigo from dpedido where tCodigoPedido ='" & sPedido & "'", Cn)
        If i < 0 Then
          MsgBox "El descuento es Negativo favor de Revisar descuentos u Ofertas Asignados!!", vbExclamation, sMensaje
          Validaciones = False
          Exit Function
        End If
        
        i = Calcular("select sum(nventa) as codigo from dpedido where tCodigoPedido ='" & sPedido & "'", Cn)
        If i < 0 Then
          MsgBox "El total del pedido Es Negativo Favor de Revisar!!", vbExclamation, sMensaje
          Validaciones = False
          Exit Function
        End If
    
        
        'Validacion de obligatoriedad de Observacion
        If lObservacion And Trim(txtObservacion.Caption) = "" Then
            MsgBox "Debes ingresar la Observación", vbInformation, sMensaje
            cmdCabecera_Click (5)
            If Trim(txtObservacion.Caption) = "" Then
              Validaciones = False
              Exit Function
            End If
        End If

        'Correccion
        If lMCPV Then
            If MultiCajeroOk = False Then
              Validaciones = False
              Exit Function
            End If
        End If
 
    End If

    Validaciones = True
    Exit Function
fin:
    Validaciones = False
    MsgBox ("Error:" & Error & " | Proceso: Validaciones al Generar Documento - POS")
End Function

Function Revisa_Descuento() As Boolean
On Error GoTo fin:
     'Chequea Descuento
        If CDbl(TxtDescuento) > 0 Then
           Dim nTope As Double
           Dim nTotalMes As Double
           
           Dim nConsumo As Double
           Dim aplicaTope As Boolean
           
           lAplicablePedido = Calcular("select lAplicablePedido as Codigo FROM vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
           nTope = Calcular("select nTope as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn)
           
                If nTope > 0 Then
                     If Calcular("select lTopePedido as Codigo from vMotivoDescuento where lActivo=1 and Codigo='" & sCodigoDescuento & "'", Cn) Then
                        If CDbl(TxtDescuento) > nTope Then
                           If MsgBox("El Descuento a aplicar Supera El Tope Registrado por Pedido" & Chr(13) & "¿Desea aplicar el Tope de " & sMonN & " " & nTope & "?", vbQuestion + vbYesNo) = vbYes Then
                              CalculaAplicaTope (nTope)
                           Else
                                Revisa_Descuento = False
                                Exit Function
                           End If
                        End If
                     Else
                        nTotalMes = Calcular("select sum(DPEDIDO.nDescuento*nCantidad) as Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
                                             "WHERE month(MPEDIDO.fFecha) = month(getdate()) and year(MPEDIDO.fFecha)=year(getdate()) and mPedido.tDescuento='" & sCodigoDescuento & "' and tEstadoPedido<>'01' and tEstadoPedido<>'03'", Cn)
                        
                        If CDbl(TxtDescuento) + nTotalMes > nTope Then
                           If CDbl(TxtDescuento) < nTope Then
                              If MsgBox("El Descuento a aplicar Supera El Tope Registrado dentro de un mes" & Chr(13) & "¿Desea aplicar el saldo " & sMonN & " " & nTope - nTotalMes & "?", vbQuestion + vbYesNo) = vbYes Then
                                 CalculaAplicaTope (nTope - nTotalMes)
                              Else
                                Revisa_Descuento = False
                                Exit Function
                              End If
                           Else
                               MsgBox "El Descuento a aplicar Supera El Tope Registrado dentro de un mes", vbExclamation
                               Revisa_Descuento = False
                               Exit Function
                           End If
                        End If
                     End If
                End If
        End If
        sCodigoDescuento = IIf(lAplicablePedido, "", sCodigoDescuento)

    Revisa_Descuento = True
    Exit Function
fin:
    MsgBox ("Error:" & Error)
    Revisa_Descuento = False
End Function
' FACTURACION ELECTRONICA

Private Sub ReimpresionAut(tDocumento As String)
                Dim RsEmision As Recordset
                Dim nRespuesta As Integer
                Dim fso1 As Object
               lReimpresion = True
               lImprimeAlternativa = False
                              
               If RsCabecera.RecordCount = 0 Then
                  'Exit Sub
               End If
               
               sDocumento = tDocumento
               Set RsEmision = Lib.OpenRecordset("select tFormulario, tImpresora, tTipoEmision, Cliente, lResumen, Monto, lFacturacionElectronica,lImprimeImageCab from vTipoDocumentoImpresora where tTipoEmision='" & Calcular("select ttipodocumento as codigo from MDOCUMENTO where tDocumento='" & tDocumento & "'", Cn) & "' and tCaja='" & Calcular("select tcaja as codigo from MDOCUMENTO where tDocumento='" & tDocumento & "'", Cn) & "'", Cn)
               If RsEmision.RecordCount = 0 Then
                  MsgBox "Error: Configuración en Caja o El documento pertenece a otra caja", vbCritical, sMensaje
                  Exit Sub
               End If
               
'               If Supervisor("12") = False Then
'                  MsgBox "Clave no permitida", vbExclamation, sMensaje
'                  Exit Sub
'               End If
               
               Select Case pais
               Case "001" 'Bolivia
                     nRespuesta = MsgBox("Desea imprimir una copia del comprobante?", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
                     If nRespuesta = vbYes Then
                         estadoReimpresion = True
                     ElseIf nRespuesta = vbCancel Then
                         Exit Sub
                     End If

                Case Else 'Peru, Ecuador
                     estadoReimpresion = False
                    
                End Select
               
                
                'Imprime Documentos
                  Dim wConsumo As Boolean
                  wConsumo = False
                  
'                  nRespuesta = MsgBox("Por Consumo? ", vbQuestion + vbYesNoCancel + vbDefaultButton2, sMensaje)
'                  If nRespuesta = vbYes Then
'                     frmKeyBoard.txtResultado = tTextoConsumo
'                      frmKeyBoard.Show vbModal
'                      If sDescrip = "" Or Not wEnter Then
'                         MsgBox "Error: La descripcion no puede ser en blanco", vbCritical, sMensaje
'                         Exit Sub
'                      End If
'                      wConsumo = True
'                  ElseIf nRespuesta = vbCancel Then
'                      Exit Sub
'                  End If
                
                  If wConsumo = False And lDescripcionAlternativa = True Then
                     If validaImpresionAlternativa(sDocumento) = False Then
                         If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                            lImprimeAlternativa = True
                         End If
                     End If
                  End If
                  
                  TimpresionDolaresDelivery = False
                      '-------- impresion en dolares si esta activo el check en el cliente delivery.
                If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from TDELIVERY where tCodigoDelivery=(select tClienteDelivery from MPEDIDO where tCodigoPedido=(select top 1 tCodigoPedido from DPEDIDO where tDocumento='" & sDocumento & "'))", Cn) Then
                    If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        TimpresionDolaresDelivery = True
                    Else
                        TimpresionDolaresDelivery = False
                    End If
                Else
                    TimpresionDolaresDelivery = False
                End If
                  
                '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
                ' ELDCQ 22/11/2017
                If lImprimeAlternativa = False Then
                    If lDocumentoAgrupado Then
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                    End If
               Else
                    If lDocumentoAgrupado Then
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
                    Else
                        Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"

                    End If
                End If

                Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        
                'impresion imagen
                Set rstFuente = New ADODB.Recordset
                imageCab.Picture = Nothing
                imagepIE.Picture = Nothing
                Set rstFuente = Lib.OpenRecordset("select iImagenCabDoc AS foto, iImagenPieDoc as fotoPie  from tcaja where tcaja='" & sCaja & "'", Cn)
                imageCab.DataField = "foto"
                Set imageCab.DataSource = rstFuente
                imagepIE.DataField = "fotoPie"
                Set imagepIE.DataSource = rstFuente
                
                
                
                 'FACTURA ELECTRONICA
                  Dim xlFacturacionE As Boolean
                  Dim xlTipoDoc As String
                  
                  xlTipoDoc = Calcular("Select tTipoDocumento As Codigo From MDOCUMENTO Where tDocumento = '" & sDocumento & "'", Cn)
                  RsEmision.Find "tTipoEmision='" & xlTipoDoc & "'"
                  xlFacturacionE = IIf(RsEmision!lFacturacionElectronica = True, 1, 0)
                  xImpresionFE = Calcular(" SELECT tImpresionFE as codigo FROM vTipodocumento WHERE Codigo=(SELECT TTIPODOCUMENTO FROM MDOCUMENTO WHERE tDocumento='" & sDocumento & "')", Cn)
                  xImpresioDE = Mid(sDocumento, 1, 1)
                  
                  'FACTURACION_E_PERU
                  If pais = "000" Then
                     If lFacturacionE Then
                            If lFEOfisis Then
                                    Dim lDocumentoElectronicoOfisis As Boolean
                                    lDocumentoElectronicoOfisis = Calcular("select isnull(tdi.lDocumentoElectronicoOfisis,0) as codigo from TTIPODOCUMENTOIMPRESORA tdi inner join MDOCUMENTO m on tdi.tTipoEmision = m.tTipoDocumento and tdi.tCaja = m.tCaja  where m.tDocumento= '" & sDocumento & "'", Cn)
                                    If lDocumentoElectronicoOfisis Then
                                         fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + "-" + CStr(CLng(Mid(sDocumento, 8, 8)))
                                         If lImpresionCodigoBarras Then
                                             imageHash.DataField = "foto"
                                             Set RsCodigoHash = Lib.OpenRecordset("USP_FactObtenerCodigoBarraOfisis '" & fDocumento & "','" & Mid(sDocumento, 1, 1) & "','' ", Cn)
                                             Set imageHash.DataSource = RsCodigoHash
                                             
                                        ElseIf lQRFE Then
                                                Set imageHash.Picture = LoadPicture(ImagenQR_Ofisis(fDocumento, sDocumento))
                                         Else
                                             Dim RscadenaCodigoHash As Recordset
                                             Set RscadenaCodigoHash = Lib.OpenRecordset("USP_FactConsultaHash '" & fDocumento & "','0' ", Cn)
                                             If RscadenaCodigoHash.RecordCount > 0 Then
                                                 cadenaCodigoHash = RscadenaCodigoHash!codigo
                                             End If
                                             'cadenaCodigoHash = Calcular("select CO_HASH as codigo from TCFACT_ELEC where NU_DOCU='" & fDocumento & "' and (TI_DOCU='B' or TI_DOCU ='F')", CnFE)
                                         End If
                                    End If
                            
                            ElseIf lFECarbajal Then
                                If xlFacturacionE Then
                                    If tCodigoFE = "000" Then
                                         If lQRFE Then
                                             Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(3, sDocumento, 0))
                                         Else
                                             If lImpresionCodigoBarras Then
                                                 Set imageHash.Picture = LoadPicture(ImagenFeCarvajal(1, sDocumento, 0))
                                             Else
                                                 cadenaCodigoHash = ImagenFeCarvajal(2, sDocumento, 0)
                                             End If
                                         End If
                                     End If
                                  End If
                            ElseIf lFESpring Then
                                If xlFacturacionE Then
                                    If tCodigoFE = "000" Then
                                         If lQRFE Then
                                             Set imageHash.Picture = LoadPicture(ImagenFeSpring(lQRFE, sDocumento))
                                         Else
                                             If lImpresionCodigoBarras Then
                                                 'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                             Else
                                                 'cadenaCodigoHash = Calcular("select isnull(tRespFacturacion,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                                                 cadenaCodigoHash = ImagenFeSpring(lQRFE, sDocumento)
                                             End If
                                         End If
                                     End If
                                End If
                            
                            ElseIf lFEpape Then
                                If xlFacturacionE Then
                                    If tCodigoFE = "000" Then
                                         If lQRFE Then
                                             Set imageHash.Picture = LoadPicture(CrearImagenQR(TCPQR(Calcular("select isnull(tImprTermica,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn))))
                                         Else
                                             If lImpresionCodigoBarras Then
                                                 'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                             Else
                                                 cadenaCodigoHash = Calcular("select isnull(tRespFacturacion,'') as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn)
                                             End If
                                         End If
                                     End If
                                End If
                            Else
                                    'VALIDAR RESPUESTA DE CODIGO HASH
                                    If lQRFE Then
                                        Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                    Else
                                    
                                        If lImpresionCodigoBarras Then
                                            Set imageHash.Picture = LoadPicture(lValidaCodBarra(lImpresionCodigoBarras, sDocumento))
                                        Else
                                            cadenaCodigoHash = lValidaCodBarra(lImpresionCodigoBarras, sDocumento)
                                        End If
                                    End If
                                    xi = 1
                                    fDocumento = Mid(sDocumento, 1, 1) + Mid(sDocumento, 4, 3) + Mid(sDocumento, 8, 8)
                            End If

                      End If
                  End If
                  
                  '------------------------------------------
                    
               If RsImpresion.RecordCount = 0 Then
                  LimpiaRs
                  MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
               Else
                  'Configura la Impresora
                  Imprimir (RsEmision!timpresora)
                  sResumen = RsEmision!lResumen
                  Printer.FontName = sFont
                  Printer.FontBold = False
        
                  If wConsumo Then
                  
                     If RsEmision!tFormulario = "01" Then
                        If RsEmision!Cliente And RsEmision!Monto = 0 Then
                        
                           'FACTURACION ELECTRONICA
                           If xlFacturacionE And tCodigoFE <> "999" Then
                                ImprimeFacturaConsumoElectronico RsImpresion, sDescrip, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                           Else
                                ImprimeFacturaConsumoT RsImpresion, sDescrip, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                           End If
                           
                        Else
                        
                           'FACTURACION ELECTRONICA
                           If xlFacturacionE And tCodigoFE <> "999" Then
                                ImprimeBoletaConsumoElectronico RsImpresion, sDescrip, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                           Else
                                ImprimeBoletaConsumoT RsImpresion, sDescrip, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                           End If
                           
                        End If
                     Else
                        If lFacturacionE And xlFacturacionE And lFESpring = False And lFEOfisis = False And lFEpape = False And lFECarbajal = False Then
                        
                                If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                    ImprimeFormatoAConsumo
                                    Kill App.path & "\fact.bmp"
                                Else
                                    ImprimeFormatoAConsumo
                                End If
                                
                        ElseIf lFacturacionE And xlFacturacionE And lFECarbajal Then
                                ImprimeFormatoAConsumo
                                Set fso1 = CreateObject("Scripting.FileSystemObject")
                                If fso1.FileExists(App.path & "\fact.bmp") Then
                                    Kill App.path & "\fact.bmp"
                                End If
                        Else
                                If RsEmision!Cliente And RsEmision!Monto = 0 Then
                                   ImprimeFacturaConsumoN RsImpresion, sDescrip, sEmpresa
                                Else
                                   ImprimeBoletaConsumoN RsImpresion, sDescrip, sEmpresa
                                End If
                        End If
                     End If
                     
                  Else
                     If RsEmision!tFormulario = "01" Then
                            If RsEmision!Cliente And RsEmision!Monto = 0 Then
                            
                                'FACTURACION ELECTRONICA
                                If xlFacturacionE And tCodigoFE <> "999" Then
                                    ImprimeFacturaElectronica RsImpresion, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                Else
                                    ImprimeFacturaT RsImpresion, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                                End If
                            
                            ElseIf RsEmision!TTipoEmision = "00" Then
                                If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                    ImprimeCortesia RsImpresion, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE
                                End If
                            Else
                            
                                'FACTURACION ELECTRONICA
                                If xlFacturacionE And tCodigoFE <> "999" Then
                                    ImprimeBoletaElectronica RsImpresion, imageHash, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                                Else
                                    ImprimeBoletaT RsImpresion, Calcular("select ttipodocumento  as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE, TimpresionDolaresDelivery
                                End If
                            
                            End If
                        
                     Else
                            If lFacturacionE And xlFacturacionE And lFESpring = False And lFEOfisis = False And lFEpape = False And lFECarbajal = False Then
                                    If Generar_Imagen(CnFE, "select imagen from IMAGENCODIGOBARRA where nro_efact='" & fDocumento & "'", "imagen", "\fact.bmp") = True Then
                                        ImprimeFormatoA
                                        Kill App.path & "\fact.bmp"
                                    Else
                                        ImprimeFormatoA
                                    End If
                            ElseIf lFacturacionE And xlFacturacionE And lFECarbajal Then
                                    ImprimeFormatoA
                                    Set fso1 = CreateObject("Scripting.FileSystemObject")
                                    If fso1.FileExists(App.path & "\fact.bmp") Then
                                        Kill App.path & "\fact.bmp"
                                    End If
                            Else
                                    If RsEmision!Cliente And RsEmision!Monto = 0 Then
                                       ImprimeFacturaN RsImpresion, sEmpresa, xlTipoDoc
                                    ElseIf RsEmision!TTipoEmision = "00" Then
                                       If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                                          ImprimeCortesia RsImpresion, Calcular("select ttipodocumento as codigo from mdocumento where tdocumento='" & sDocumento & "'", Cn), imageCab, imagepIE
                                       End If
                                    Else
                                       ImprimeBoletaN RsImpresion, sEmpresa, xlTipoDoc
                                    End If
                            End If
                        
                     End If
                     
                  End If

               End If
               lReimpresion = False
               LimpiaRs
               Screen.MousePointer = vbDefault
               'If flagTipoCambio = True Then: nTC = 0
End Sub
Function PedidoPagoWEb(ByVal Pedido As String) As Boolean
    On Error GoTo fin
    Cn.Execute "exec usp_WS_Inforest '" & Pedido & "',3"
    PedidoPagoWEb = True
    Exit Function
fin:
    MsgBox "Error" & Error & " | Proceso de PedidoPagoWEb", vbCritical, sMensaje
    PedidoPagoWEb = False
End Function


Private Function SumPrecioVenta(ByVal CodPedido As String) As Double
On Error GoTo fin:
   Dim vley1 As Double
   vley1 = CDbl(sValorLey1) * (Calcular("select sum(nCantidad) as Codigo FROM DPEDIDO INNER JOIN VPRODUCTO ON DPEDIDO.TCODIGOPRODUCTO=VPRODUCTO.CODIGO where tEstadoItem = 'N' AND VPRODUCTO.lLEY1=1 and isnull(tfacturado,'')='' and tCodigoPedido='" & CodPedido & "'", Cn))
   txtSuma.Caption = Format(Calcular("select sum(nVenta)+(" & vley1 & ") as Codigo FROM DPEDIDO where tEstadoItem = 'N' and isnull(tfacturado,'')='' and tCodigoPedido='" & CodPedido & "'", Cn), "#,###,##0.00")
   lblLey1.Caption = sLey1 & " " & Format(vley1, "##0.00")
   SumPrecioVenta = CDbl(txtSuma.Caption)
   Exit Function
fin:
    SumPrecioVenta = 9999999
    Call Log_Inforest("PUNTO VENTA", "SumPrecioVenta", CodPedido, "", "", Error, "SUMA DE PRECIO DE VENTA INFOREST", "", sUsuario)
    Exit Function
End Function

Private Function VerificaComboMaximo() As Boolean
   'Miguel Mendoza - 10062019 XD XD XD
   Dim nPos As Integer
   VerificaComboMaximo = True
   With RsDetalle
      If .RecordCount > 0 Then
         .MoveFirst

         Do While Not .EOF
            iProductoVerificaComboMaximo = 0
            iProductoVerificaComboMinimo = 0
            iProductoVerificaComboUsado = 0
            iProductoVerificaComboMaximo = _
            Calcular("select isnull(ncombinacion,0) codigo FROM TPRODUCTO WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)

            iProductoVerificaComboMinimo = _
            Calcular("select isnull(ncombinacionminima,0) codigo FROM TPRODUCTO WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)
            
            If iProductoVerificaComboMaximo > 0 Then
               sProductoVerificaComboUsado = _
               Calcular("select isnull(tdetallado,'') codigo FROM tproducto WHERE tCodigoProducto='" & Trim(.Fields(3)) & "'", Cn)

               iProductoVerificaComboUsado = _
               Calcular("select isnull(sum(nCantidad),0) codigo FROM CPEDIDO WHERE tCodigoPedido='" & sPedido & "' " & _
               "and tproducto='" & Trim(.Fields(3)) & "' and titem='" & RsDetalle!tItem & "'", Cn)

               If iProductoVerificaComboUsado > iProductoVerificaComboMaximo * CInt(.Fields(15)) Then
                  MsgBox "La cantidad máxima para el Combo " & UCase(Trim(sProductoVerificaComboUsado)) & " debe ser de " & _
                  Trim(CStr(iProductoVerificaComboMaximo * CInt(.Fields(15)))), vbCritical + vbOKOnly, "Inforest"
                  VerificaComboMaximo = False
                  'RsDetalle.Requery
                    nPos = RsDetalle.AbsolutePosition
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                  Exit Function
               End If
               
               If iProductoVerificaComboUsado < iProductoVerificaComboMinimo * CInt(.Fields(15)) Then
                  MsgBox "La cantidad mínima para el Combo " & UCase(Trim(sProductoVerificaComboUsado)) & " debe ser de " & _
                  Trim(CStr(iProductoVerificaComboMinimo * CInt(.Fields(15)))), vbCritical + vbOKOnly, "Inforest"
                  VerificaComboMaximo = False
                  'RsDetalle.Requery
                    nPos = RsDetalle.AbsolutePosition
                    RsDetalle.Requery
                    RsDetalle.AbsolutePosition = nPos
                  Exit Function
               End If
           End If
            .MoveNext
         Loop
      End If
   End With
End Function

' Helper para leer campos como cadena segura (evita Null y RS vacío)
Private Function FieldStr(Rs As ADODB.Recordset, ByVal fld As String) As String
    On Error Resume Next
    FieldStr = ""
    If Rs Is Nothing Then Exit Function
    If Rs.BOF Or Rs.EOF Then Exit Function
    If IsNull(Rs.Fields(fld).value) Then Exit Function
    FieldStr = Trim$(CStr(Rs.Fields(fld).value))
End Function


