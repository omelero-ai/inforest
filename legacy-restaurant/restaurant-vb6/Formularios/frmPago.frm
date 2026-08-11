VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPago 
   BackColor       =   &H00E2DEDB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12045
   Icon            =   "frmPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer_CashDro 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1080
      Top             =   8880
   End
   Begin VB.Frame frmMensaje 
      BackColor       =   &H00F0F0F0&
      Height          =   8775
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   12015
      Begin VB.CommandButton CmdCancelaCashDro 
         BackColor       =   &H00F2FAED&
         Caption         =   "Cancelar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4680
         Picture         =   "frmPago.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   7080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton BtnCancelarFPAY 
         BackColor       =   &H00F2FAED&
         Caption         =   "Cancelar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4680
         Picture         =   "frmPago.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   6240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelarMercadopago 
         Caption         =   "Cancelar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4680
         Picture         =   "frmPago.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   6240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Image5 
         Height          =   3480
         Left            =   2040
         Picture         =   "frmPago.frx":0610
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   6960
      End
      Begin VB.Image Image2 
         Height          =   3360
         Left            =   2040
         Picture         =   "frmPago.frx":A5EE
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   6960
      End
      Begin VB.Image Image1 
         Height          =   2760
         Left            =   3480
         Picture         =   "frmPago.frx":F992
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   4200
      End
      Begin VB.Label lblMensajetexto 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1935
         Left            =   2040
         TabIndex        =   80
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.PictureBox cmdOpcionPicture 
      BackColor       =   &H00E2DEDB&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   10800
      Picture         =   "frmPago.frx":215737
      ScaleHeight     =   615
      ScaleWidth      =   1170
      TabIndex        =   159
      Top             =   8160
      Width           =   1170
   End
   Begin VB.Timer Timer_Fpay 
      Enabled         =   0   'False
      Left            =   600
      Top             =   8880
   End
   Begin VB.Timer Timer_MercadoPago 
      Enabled         =   0   'False
      Left            =   120
      Top             =   8880
   End
   Begin VB.CommandButton cmdMercadoPago 
      BackColor       =   &H00F2FAED&
      Caption         =   "QR Mercado Pago"
      Height          =   735
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraEfectivo 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Efectivo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   1395
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   3330
      Begin VB.CommandButton cmdPago 
         Caption         =   "Monto"
         Height          =   555
         Index           =   17
         Left            =   150
         TabIndex        =   75
         Top             =   2040
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Monto"
         Height          =   555
         Index           =   16
         Left            =   150
         TabIndex        =   74
         Top             =   1440
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00649D34&
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
         Height          =   555
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00649D34&
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
         Height          =   555
         Index           =   2
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label txtME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   3
         Left            =   1530
         TabIndex        =   77
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label txtME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   2
         Left            =   1530
         TabIndex        =   76
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label txtME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   4
         Top             =   345
         Width           =   1665
      End
      Begin VB.Label txtME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   1
         Left            =   1530
         TabIndex        =   3
         Top             =   900
         Width           =   1665
      End
   End
   Begin VB.Frame fraCheque 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Cheque / Depósito "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   2450
      Left            =   30
      TabIndex        =   5
      Top             =   1380
      Width           =   3330
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Monto"
         Height          =   555
         Index           =   3
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1905
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Banco"
         Height          =   555
         Index           =   4
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Número"
         Height          =   555
         Index           =   5
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   795
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Monto"
         Height          =   555
         Index           =   6
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label txtMCH 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   1
         Left            =   1530
         TabIndex        =   73
         Top             =   2010
         Width           =   1695
      End
      Begin VB.Label txtBCH 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1530
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label txtNCH 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1530
         TabIndex        =   10
         Top             =   915
         Width           =   1695
      End
      Begin VB.Label txtMCH 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   9
         Top             =   1470
         Width           =   1695
      End
   End
   Begin VB.Frame fraPunto 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Puntos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   885
      Left            =   45
      TabIndex        =   58
      Top             =   3820
      Width           =   3330
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Monto"
         Height          =   555
         Index           =   14
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label txtCliente 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1530
         TabIndex        =   61
         Top             =   510
         Width           =   1695
      End
      Begin VB.Label txtPunto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1530
         TabIndex        =   60
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame fraDocumento 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Otros Pagos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   765
      Left            =   45
      TabIndex        =   53
      Top             =   4680
      Width           =   6735
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Otros"
         Height          =   555
         Index           =   13
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   160
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Nota de Crédito"
         Height          =   555
         Index           =   11
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   160
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   3240
         TabIndex        =   64
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia :"
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   63
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label txtReferencia 
         BackColor       =   &H00E2DEDB&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   290
         Left            =   5040
         TabIndex        =   62
         Top             =   450
         Width           =   1650
      End
      Begin VB.Label txtDocumento 
         BackColor       =   &H00E2DEDB&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   290
         Left            =   2760
         TabIndex        =   55
         Top             =   150
         Width           =   2115
      End
      Begin VB.Label txtDocMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2DEDB&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   290
         Left            =   5040
         TabIndex        =   54
         Top             =   150
         Width           =   1650
      End
   End
   Begin VB.Frame fraOpciones 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009E9188&
      Height          =   2640
      Left            =   10320
      TabIndex        =   48
      Top             =   5440
      Width           =   1665
      Begin VB.PictureBox cmdOpcionPicture 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   2
         Left            =   120
         Picture         =   "frmPago.frx":21B745
         ScaleHeight     =   510
         ScaleWidth      =   1410
         TabIndex        =   158
         Top             =   2040
         Width           =   1410
      End
      Begin VB.PictureBox cmdOpcionPicture 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   1
         Left            =   120
         Picture         =   "frmPago.frx":22167C
         ScaleHeight     =   510
         ScaleWidth      =   1410
         TabIndex        =   157
         Top             =   280
         Width           =   1410
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Cuentas por Cobrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1420
         Width           =   1410
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
         Left            =   2280
         Picture         =   "frmPago.frx":227765
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2520
         Width           =   1550
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
         Left            =   3120
         Picture         =   "frmPago.frx":227857
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   1550
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
         Height          =   555
         Index           =   2
         Left            =   2280
         Picture         =   "frmPago.frx":227959
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1800
         Width           =   1550
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Transferencia Gratuita"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   840
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Tipos de Tarjeta de Crédito o Pago "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   4695
      Left            =   3360
      TabIndex        =   21
      Top             =   0
      Width           =   4830
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   1
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   2
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   780
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   3
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   4
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1860
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   5
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   6
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2940
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   7
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00F2FAED&
         Height          =   545
         Index           =   8
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4020
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Index           =   9
         Left            =   2025
         TabIndex        =   47
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Propina"
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
         Index           =   10
         Left            =   3540
         TabIndex        =   46
         Top             =   180
         Width           =   660
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   1
         Left            =   1575
         TabIndex        =   45
         Top             =   470
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   1575
         TabIndex        =   44
         Top             =   1010
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   3
         Left            =   1575
         TabIndex        =   43
         Top             =   1550
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   4
         Left            =   1575
         TabIndex        =   42
         Top             =   2090
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   5
         Left            =   1575
         TabIndex        =   41
         Top             =   2630
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   6
         Left            =   1575
         TabIndex        =   40
         Top             =   3170
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   7
         Left            =   1575
         TabIndex        =   39
         Top             =   3710
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   8
         Left            =   1575
         TabIndex        =   38
         Top             =   4250
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   1
         Left            =   3180
         TabIndex        =   37
         Top             =   470
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   2
         Left            =   3180
         TabIndex        =   36
         Top             =   1010
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   3
         Left            =   3180
         TabIndex        =   35
         Top             =   1550
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   4
         Left            =   3180
         TabIndex        =   34
         Top             =   2090
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   5
         Left            =   3180
         TabIndex        =   33
         Top             =   2630
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   6
         Left            =   3180
         TabIndex        =   32
         Top             =   3170
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   7
         Left            =   3180
         TabIndex        =   31
         Top             =   3710
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Index           =   8
         Left            =   3180
         TabIndex        =   30
         Top             =   4250
         Width           =   1515
      End
   End
   Begin VB.Frame fraGrilla 
      BackColor       =   &H00F0F0F0&
      Caption         =   " Tarjeta de Crédito "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   3255
      Left            =   30
      TabIndex        =   12
      Top             =   5440
      Width           =   10215
      Begin TabDlg.SSTab SSTab1 
         Height          =   3015
         Left            =   3240
         TabIndex        =   71
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5318
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   15790320
         ForeColor       =   4210752
         TabCaption(0)   =   "Tarjeta de Credito"
         TabPicture(0)   =   "frmPago.frx":227E8B
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Recibo de Ingreso"
         TabPicture(1)   =   "frmPago.frx":227EA7
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "PagoEfectivo"
         TabPicture(2)   =   "frmPago.frx":227EC3
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frmPagoEfectivo"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "FPAY"
         TabPicture(3)   =   "frmPago.frx":227EDF
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "frmFPAY"
         Tab(3).ControlCount=   1
         Begin VB.Frame Frame4 
            BackColor       =   &H00F0F0F0&
            Height          =   2655
            Left            =   -75000
            TabIndex        =   101
            Top             =   320
            Width           =   6975
            Begin VB.CommandButton cmdOpcion 
               BackColor       =   &H00F2FAED&
               Caption         =   "Cancelar"
               Height          =   560
               Index           =   9
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   109
               Top             =   2050
               Visible         =   0   'False
               Width           =   1400
            End
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   4
               Left            =   6165
               Picture         =   "frmPago.frx":227EFB
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   120
               Width           =   600
            End
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   5
               Left            =   6165
               Picture         =   "frmPago.frx":2287C5
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   600
               Width           =   600
            End
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   6
               Left            =   6165
               Picture         =   "frmPago.frx":22908F
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   1080
               Width           =   600
            End
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   7
               Left            =   6165
               Picture         =   "frmPago.frx":229959
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   1560
               Width           =   600
            End
            Begin VB.CommandButton cmdPago 
               BackColor       =   &H00F2FAED&
               Caption         =   "Buscar Recibo"
               Height          =   560
               Index           =   15
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   104
               Top             =   2050
               Width           =   1400
            End
            Begin VB.CommandButton cmdOpcion 
               BackColor       =   &H00F2FAED&
               Caption         =   "Eliminar"
               Height          =   560
               Index           =   11
               Left            =   1520
               Style           =   1  'Graphical
               TabIndex        =   103
               Top             =   2050
               Width           =   1400
            End
            Begin VB.CommandButton cmdPago 
               BackColor       =   &H00F2FAED&
               Caption         =   "Agregar Recibos Enlazados "
               Height          =   560
               Index           =   12
               Left            =   2910
               Style           =   1  'Graphical
               TabIndex        =   102
               Top             =   2050
               Width           =   1400
            End
            Begin TrueOleDBGrid80.TDBGrid grdGrillaRecibo 
               Height          =   1695
               Left            =   0
               TabIndex        =   111
               Top             =   120
               Width           =   6195
               _ExtentX        =   10927
               _ExtentY        =   2990
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
               PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
               PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   7
               Caption         =   "Cancelación con recibos"
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
               _StyleDefs(7)   =   ":id=1,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Small Fonts"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
               _StyleDefs(10)  =   ":id=4,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00F0F0F0&
            Height          =   2655
            Left            =   0
            TabIndex        =   93
            Top             =   320
            Width           =   6975
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   0
               Left            =   6165
               Picture         =   "frmPago.frx":22A223
               Style           =   1  'Graphical
               TabIndex        =   100
               Top             =   120
               Width           =   600
            End
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   1
               Left            =   6165
               Picture         =   "frmPago.frx":22AAED
               Style           =   1  'Graphical
               TabIndex        =   99
               Top             =   600
               Width           =   600
            End
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   2
               Left            =   6165
               Picture         =   "frmPago.frx":22B3B7
               Style           =   1  'Graphical
               TabIndex        =   98
               Top             =   1080
               Width           =   600
            End
            Begin VB.CommandButton cmdNavegar 
               BackColor       =   &H00F2FAED&
               Height          =   500
               Index           =   3
               Left            =   6165
               Picture         =   "frmPago.frx":22BC81
               Style           =   1  'Graphical
               TabIndex        =   97
               Top             =   1560
               Width           =   600
            End
            Begin VB.CommandButton cmdOpcion 
               BackColor       =   &H00F2FAED&
               Caption         =   "Cancelar"
               Height          =   555
               Index           =   5
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   96
               Top             =   2050
               Width           =   1300
            End
            Begin VB.CommandButton cmdOpcion 
               BackColor       =   &H00F2FAED&
               Caption         =   "Agregar"
               Height          =   555
               Index           =   4
               Left            =   4830
               Style           =   1  'Graphical
               TabIndex        =   95
               Top             =   2050
               Width           =   1300
            End
            Begin VB.CommandButton cmdOpcion 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Eliminar"
               Height          =   555
               Index           =   6
               Left            =   2130
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   2050
               Width           =   1300
            End
            Begin TrueOleDBGrid80.TDBGrid grdGrilla 
               Height          =   1935
               Left            =   0
               TabIndex        =   110
               Top             =   120
               Width           =   6195
               _ExtentX        =   10927
               _ExtentY        =   3413
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
               PrintInfos(0).PageHeaderFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
               PrintInfos(0).PageFooterFont=   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Small Fonts"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   7
               Caption         =   "Cancelación con Tarjetas"
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
               _StyleDefs(7)   =   ":id=1,.fontsize=675,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Small Fonts"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H4D7C2A&,.fgcolor=&HFFFFFF&"
               _StyleDefs(10)  =   ":id=4,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         End
         Begin VB.Frame frmFPAY 
            BackColor       =   &H00F0F0F0&
            Enabled         =   0   'False
            Height          =   2655
            Left            =   -75000
            TabIndex        =   91
            Top             =   320
            Width           =   6975
            Begin VB.CommandButton BtnFpay 
               BackColor       =   &H00F2FAED&
               Caption         =   "Generar Intención de Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   92
               Top             =   720
               Width           =   2175
            End
            Begin VB.Image Image4 
               Height          =   1095
               Left            =   480
               Picture         =   "frmPago.frx":22C54B
               Stretch         =   -1  'True
               Top             =   600
               Width           =   2760
            End
         End
         Begin VB.Frame frmPagoEfectivo 
            BackColor       =   &H00F0F0F0&
            Enabled         =   0   'False
            Height          =   2655
            Left            =   -75000
            TabIndex        =   83
            Top             =   320
            Width           =   6975
            Begin VB.CommandButton cmdPagoEfectivo 
               BackColor       =   &H0080FF80&
               Caption         =   "Generar CIP PagoEfectivo"
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
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   85
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton cmdCancelarPagoEfectivo 
               BackColor       =   &H008080FF&
               Caption         =   "Cancelar Proceso"
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
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   1680
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Image Image3 
               Height          =   765
               Left            =   120
               Picture         =   "frmPago.frx":236529
               Stretch         =   -1  'True
               Top             =   120
               Width           =   2415
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   $"frmPago.frx":2391E6
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   2760
               TabIndex        =   89
               Top             =   120
               Width           =   3975
            End
            Begin VB.Line Line1 
               X1              =   2040
               X2              =   2040
               Y1              =   960
               Y2              =   2280
            End
            Begin VB.Label lblCIPpagoEfectivo 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "CIP:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   88
               Top             =   1440
               Visible         =   0   'False
               Width           =   4215
            End
            Begin VB.Label lblInformacion 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Espere.. hasta que el documento sea cancelado, o puede cancelar el proceso para efecturar otro medio de pago."
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   2160
               TabIndex        =   87
               Top             =   1920
               Visible         =   0   'False
               Width           =   4695
            End
            Begin VB.Label lblclick 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Click sobre CIP para abrir en el explorador de internet"
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   2160
               TabIndex        =   86
               Top             =   1200
               Visible         =   0   'False
               Width           =   4095
            End
         End
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00649D34&
         Caption         =   "Número de Autorización"
         Height          =   555
         Index           =   1
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   2500
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00649D34&
         Caption         =   "Propina"
         Height          =   555
         Index           =   10
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1930
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00649D34&
         Caption         =   "Monto"
         Height          =   555
         Index           =   9
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1360
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Fecha Vencimiento"
         Height          =   555
         Index           =   8
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   800
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         BackColor       =   &H00F2FAED&
         Caption         =   "Número"
         Height          =   555
         Index           =   7
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label txtNAT 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1395
         TabIndex        =   69
         Top             =   2600
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta de Credito :"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   67
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Label lblTarjetaC 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4800
         TabIndex        =   66
         Top             =   120
         Width           =   45
      End
      Begin VB.Label txtNT 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1395
         TabIndex        =   20
         Top             =   390
         Width           =   1785
      End
      Begin VB.Label txtVT 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   1395
         TabIndex        =   19
         Top             =   945
         Width           =   1785
      End
      Begin VB.Label txtMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1395
         TabIndex        =   18
         Top             =   1500
         Width           =   1785
      End
      Begin VB.Label txtPT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   1395
         TabIndex        =   17
         Top             =   2055
         Width           =   1785
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6855
      Left            =   8280
      TabIndex        =   65
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   15790320
      ForeColor       =   4210752
      TabCaption(0)   =   "Resultados"
      TabPicture(0)   =   "frmPago.frx":2392AD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraExtranjero"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Desglose"
      TabPicture(1)   =   "frmPago.frx":2392C9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraExtranjero 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   1680
         TabIndex        =   78
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00F0F0F0&
         Height          =   5415
         Left            =   -75000
         TabIndex        =   112
         Top             =   320
         Width           =   3735
         Begin VB.CommandButton cmdDivision 
            BackColor       =   &H00F2FAED&
            Caption         =   "División"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   1
            Left            =   2250
            Picture         =   "frmPago.frx":2392E5
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   510
            Width           =   1065
         End
         Begin VB.CommandButton cmdDivision 
            BackColor       =   &H00F2FAED&
            Caption         =   "División"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   2
            Left            =   2250
            Picture         =   "frmPago.frx":2393E7
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   1230
            Width           =   1065
         End
         Begin VB.CommandButton cmdCantidad 
            BackColor       =   &H00F2FAED&
            Caption         =   "Cantidad"
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
            Left            =   2250
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   1950
            Width           =   1065
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00F0F0F0&
            Caption         =   " Dividir en "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   735
            Left            =   2250
            TabIndex        =   113
            Top             =   2670
            Width           =   1110
            Begin VB.Label txtDivision 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
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
               Height          =   300
               Left            =   330
               TabIndex        =   114
               Top             =   330
               Width           =   450
            End
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total :"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   129
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento :"
            Height          =   195
            Index           =   19
            Left            =   150
            TabIndex        =   128
            Top             =   960
            Width           =   870
         End
         Begin VB.Label lblImpuesto1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto1 :"
            Height          =   195
            Left            =   150
            TabIndex        =   127
            Top             =   2400
            Width           =   825
         End
         Begin VB.Label lblImpuesto2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto2 :"
            Height          =   195
            Left            =   150
            TabIndex        =   126
            Top             =   3120
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            Height          =   195
            Left            =   150
            TabIndex        =   125
            Top             =   4080
            Width           =   450
         End
         Begin VB.Label txtSubTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Left            =   120
            TabIndex        =   124
            Top             =   510
            Width           =   1500
         End
         Begin VB.Label txtDescuento 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Left            =   150
            TabIndex        =   123
            Top             =   1230
            Width           =   1500
         End
         Begin VB.Label txtImpuesto1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Left            =   150
            TabIndex        =   122
            Top             =   2670
            Width           =   1500
         End
         Begin VB.Label txtImpuesto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Left            =   150
            TabIndex        =   121
            Top             =   3390
            Width           =   1500
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recargo :"
            Height          =   195
            Index           =   14
            Left            =   150
            TabIndex        =   120
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label txtRecargo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Left            =   150
            TabIndex        =   119
            Top             =   1950
            Width           =   1500
         End
         Begin VB.Label txtTotalNeto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Left            =   150
            TabIndex        =   118
            Top             =   4350
            Width           =   1500
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00F0F0F0&
         Height          =   6485
         Left            =   0
         TabIndex        =   130
         Top             =   320
         Width           =   3735
         Begin VB.CommandButton cmdOpcion 
            BackColor       =   &H00F2FAED&
            Caption         =   "Hora de Entrega"
            Height          =   345
            Index           =   7
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   4680
            Width           =   1500
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abonos :"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   156
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo :"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   155
            Top             =   1935
            Width           =   495
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Propinas :"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   154
            Top             =   2430
            Width           =   705
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vuelto :"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   153
            Top             =   2910
            Width           =   540
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Recibido :"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   152
            Top             =   3405
            Width           =   1125
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cortesia :"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   151
            Top             =   4020
            Width           =   660
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cargos :"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   150
            Top             =   915
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda Nacional"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   149
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda Extranjera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   148
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio :"
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
            Left            =   120
            TabIndex        =   147
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblTc 
            BackStyle       =   0  'Transparent
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
            Left            =   1800
            TabIndex        =   146
            Top             =   240
            Width           =   705
         End
         Begin VB.Label txtCargo 
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
            Index           =   1
            Left            =   1920
            TabIndex        =   145
            Top             =   1140
            Width           =   1500
         End
         Begin VB.Label txtAbono 
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
            Index           =   1
            Left            =   1920
            TabIndex        =   144
            Top             =   1665
            Width           =   1500
         End
         Begin VB.Label txtSaldo 
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
            Index           =   1
            Left            =   1920
            TabIndex        =   143
            Top             =   2160
            Width           =   1500
         End
         Begin VB.Label txtPropina 
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
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   142
            Top             =   2655
            Width           =   1500
         End
         Begin VB.Label txtVuelto 
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
            Index           =   1
            Left            =   1920
            TabIndex        =   141
            Top             =   3135
            Visible         =   0   'False
            Width           =   1500
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
            Index           =   1
            Left            =   1920
            TabIndex        =   140
            Top             =   3630
            Width           =   1500
         End
         Begin VB.Label txtFechaDelivery 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Left            =   1920
            TabIndex        =   139
            Top             =   4680
            Width           =   1500
         End
         Begin VB.Label txtCargo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Index           =   0
            Left            =   120
            TabIndex        =   138
            Top             =   1140
            Width           =   1500
         End
         Begin VB.Label txtAbono 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Index           =   0
            Left            =   120
            TabIndex        =   137
            Top             =   1665
            Width           =   1500
         End
         Begin VB.Label txtSaldo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Index           =   0
            Left            =   120
            TabIndex        =   136
            Top             =   2160
            Width           =   1500
         End
         Begin VB.Label txtPropina 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            ForeColor       =   &H00808000&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   135
            Top             =   2655
            Width           =   1500
         End
         Begin VB.Label txtVuelto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Index           =   0
            Left            =   120
            TabIndex        =   134
            Top             =   3135
            Width           =   1500
         End
         Begin VB.Label txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            Index           =   0
            Left            =   120
            TabIndex        =   133
            Top             =   3630
            Width           =   1500
         End
         Begin VB.Label txtCortesia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0F0F0&
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
            ForeColor       =   &H00004080&
            Height          =   285
            Left            =   120
            TabIndex        =   132
            Top             =   4245
            Width           =   3300
         End
      End
   End
   Begin VB.TextBox txtQRRecibo 
      Height          =   375
      Left            =   10920
      TabIndex        =   161
      Top             =   8280
      Width           =   975
   End
   Begin VB.Image imageHash 
      Height          =   855
      Left            =   12360
      Top             =   7800
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim RsDetalle As Recordset
Dim sTarjeta As String
Dim RsTarjeta As Recordset
Dim RsGrilla As Recordset
Dim rst1 As Recordset
Dim RsT2 As Recordset
Dim RsPrepago As Recordset

'GCAA
Dim RsSaberTipoPedido As Recordset

Dim tUsuActua As String
Dim sUsuarioAutoriza As String
Public nEN As Double
Dim nEE As Double
Dim nEE2 As Double
Dim nEE3 As Double
Dim nCH As Double
Dim nCE As Double
Dim nET As Double
Dim nPT As Double
Dim nDocumento As Double

Public nAbonoN As Double
Dim nAbonoE As Double
Dim nAbonoE2 As Double
Dim nAbonoE3 As Double
Dim nTotalN As Double
Dim nTotalE As Double

Dim nVueltoN As Double
Dim nVueltoE As Double
Dim nVueltoE2 As Double
Dim nVueltoE3 As Double
Dim nPropinaN As Double
Dim nPropinaE As Double

Dim nSaldo As Double
Dim mTarjeta(8, 4)
Dim sTipoTarjeta As String
Dim sCodigoIntegracion As String
Dim sTitulo As String
Dim nIndex As Integer
Dim sCortesia As String
Dim sTipoDocumento As String
Dim sOtroTipoCancelacion As String
Dim sMonDoc As String
Dim nTotalPuntos As Double
Dim sClientePuntos As String
Dim sOtroTipo As String

Dim nRet As Integer
Dim sOperacion As String
Dim sRetorno As String * 512
Dim sClave As String
Dim sMonto As String
Dim xError As String
Dim sRefer As String
Dim nCorrela As String
Dim lEmisor As Boolean
Dim lLoop As Boolean
Dim nContador As Integer

Dim xTabla As String
Dim RsGrillaRecibido As Recordset

'mastarjetas
Dim lTarjetaGrilla As Boolean
Dim lPinpadTarjeta As Boolean
'mastarjetas


'DESGLOSE
Dim dsTotalNeto As Double
Dim dsImpuesto2 As Double
Dim dsImpuesto1 As Double
Dim dsRecargo As Double
Dim dsDescuento As Double
Dim dsSubTotal As Double


Dim lTarjetaActiva As Boolean
Dim lBotonTarjeta As Boolean

Dim xTiporecepcion As Integer

Dim lGrabaPagoAntes As Boolean

Dim tCodigoIntegracion As Integer

Dim lActivaReciboSaldoQr As Boolean

'jsjr 07-07-2025 ---------
Dim RsOperationTerminal As Recordset
Dim RsOperationCaja As Recordset
'--------------------------------------

Public TerminalSeleccionado As String

Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long, ByVal wIDNewItem As Long, _
    ByVal lpString As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" _
    (ByVal hwnd As Long) As Long
'
Private Const MF_BYCOMMAND = &H0&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
'
Private Const SC_CLOSE = &HF060&

Dim RsCanalesVenta As Recordset
Dim RsCanalxDocumento As Recordset



Private Sub BtnCancelarFPAY_Click()
On Error GoTo fin
    DoEvents
    Screen.MousePointer = vbHourglass
    DoEvents
    If (EjecutaMotorFPAY(sDocumento, sUsuario, 1, 2) = True) Then

'        Me.frmMensaje.Visible = True
'        Me.lblMensajetexto.Caption = "Esperando confirmacion de pago por la Integraciòn con Mercado Pago" & vbNewLine & vbNewLine & "TOTAL:" & Me.txtSaldo(0).Caption
'        Me.Image1.Visible = False
'        Me.Image2.Visible = True
'        Me.cmdCancelarMercadopago.Visible = True
'
'        Frame1.Enabled = False
'        Me.fraEfectivo.Enabled = False
'
'
'        Me.Timer_MercadoPago.Enabled = True
'        Me.Timer_MercadoPago.Interval = 1000
'        Me.Timer_MercadoPago.Enabled = True

    Else

'        MsgBox "No se pudo realizar el envio del pago a Mercado Pago!!!", vbCritical
'        DoEvents
'        Me.frmMensaje.Visible = False
'        Frame1.Enabled = True
'        Me.fraEfectivo.Enabled = True

    End If
    
    Screen.MousePointer = vbDefault
    Me.frmMensaje.Visible = False
    Frame1.Enabled = True
    Me.fraEfectivo.Enabled = True
    Me.Timer_Fpay.Enabled = False
    Me.Timer_Fpay.Interval = 0
    Exit Sub
fin:
MsgBox "Error: " & Error
End Sub

Private Sub BtnFpay_Click()
On Error GoTo fin

    
If (Calcular("select count(*) as codigo from ttarjetacredito where tCodigoIntegracion = 11", Cn) = 0) Then
    MsgBox "No hay tarjeta configurada para la integraciòn, favor de verificar las configuraciones de tarjetas!!!", vbInformation
    Exit Sub
End If
    Screen.MousePointer = vbHourglass
    DoEvents
    
       If (EjecutaMotorFPAY(sDocumento, sUsuario, 1, 1) = True) Then
            Me.frmMensaje.Visible = True
            Me.lblMensajetexto.Caption = "Esperando confirmación de pago por la Integración con FPAY " & vbNewLine & vbNewLine & "TOTAL:" & Me.txtSaldo(0).Caption
            Me.Image1.Visible = False
            Me.Image2.Visible = False
            Me.Image5.Visible = True
            Me.BtnCancelarFPAY.Visible = True
        
            Frame1.Enabled = False
            Me.fraEfectivo.Enabled = False
        
        
            Me.Timer_Fpay.Enabled = True
            Me.Timer_Fpay.Interval = 1000
            Me.Timer_Fpay.Enabled = True
            If lPagoAntesImpresion Then
                lGrabaPagoAntes = True
            End If
         
 
       Else
            DoEvents
            Me.frmMensaje.Visible = False
            Frame1.Enabled = True
            Me.fraEfectivo.Enabled = True
            MsgBox "No se pudo generar la Intención de Pago!, Favor de contactar al administrador de su servicios", vbInformation
'        Call ActivaDesactivaPagoEfectivo(True)
       End If
    
    Screen.MousePointer = vbDefault
Exit Sub
fin:
    'Call ActivaDesactivaPagoEfectivo(True)
    MsgBox "Error: " & Error
End Sub

Private Sub CmdCancelaCashDro_Click()
    On Error GoTo fin
    If (MsgBox("¿Esta seguro de cancelar, debe haber cancelado la transaccion en la integración? No: Continuar...", vbYesNo) = vbYes) Then
        Me.Timer_CashDro.Enabled = False
        Me.frmMensaje.Visible = False
    End If
fin:
End Sub

Private Sub cmdCancelarMercadopago_Click()
On Error GoTo fin
    DoEvents
    Screen.MousePointer = vbHourglass
    DoEvents
    If (EjecutaMotorIntegraciones(sDocumento, txtSaldo(0).Caption, 5, 5) = True) Then

'        Me.frmMensaje.Visible = True
'        Me.lblMensajetexto.Caption = "Esperando confirmacion de pago por la Integraciòn con Mercado Pago" & vbNewLine & vbNewLine & "TOTAL:" & Me.txtSaldo(0).Caption
'        Me.Image1.Visible = False
'        Me.Image2.Visible = True
'        Me.cmdCancelarMercadopago.Visible = True
'
'        Frame1.Enabled = False
'        Me.fraEfectivo.Enabled = False
'
'
'        Me.Timer_MercadoPago.Enabled = True
'        Me.Timer_MercadoPago.Interval = 1000
'        Me.Timer_MercadoPago.Enabled = True

    Else

'        MsgBox "No se pudo realizar el envio del pago a Mercado Pago!!!", vbCritical
'        DoEvents
'        Me.frmMensaje.Visible = False
'        Frame1.Enabled = True
'        Me.fraEfectivo.Enabled = True

    End If
    
    Screen.MousePointer = vbDefault
    Me.frmMensaje.Visible = False
    Frame1.Enabled = True
    Me.fraEfectivo.Enabled = True
    Me.Timer_MercadoPago.Enabled = False
    Me.Timer_MercadoPago.Interval = 0
    Exit Sub
fin:
MsgBox "Error: " & Error
End Sub


Private Sub cmdCantidad_Click(Index As Integer)
On Error GoTo fin
   sTipo = "XX"
   
   frmNumPad.Show vbModal
   
   If wEnter And val(sDescrip) > 0 Then
   
            txtTotalNeto.Caption = Format(dsTotalNeto / sDescrip, "###,###,###,##0.00")
            txtImpuesto2.Caption = Format(dsImpuesto2 / sDescrip, "###,###,###,##0.00")
            txtImpuesto1.Caption = Format(dsImpuesto1 / sDescrip, "###,###,###,##0.00")
            TxtRecargo.Caption = Format(dsRecargo / sDescrip, "###,###,###,##0.00")
            TxtDescuento.Caption = Format(dsDescuento / sDescrip, "###,###,###,##0.00")
            txtSubTotal.Caption = Format(dsSubTotal / sDescrip, "###,###,###,##0.00")
            
            txtDivision.Caption = sDescrip
   
   End If
   Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "cmdCantidad_Click", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Private Sub cmdDivision_Click(Index As Integer)
On Error GoTo fin
    Select Case Index
        Case Is = 1
            txtDivision.Caption = Format(val(txtDivision.Caption) + 1, "###")
            txtTotalNeto.Caption = Format(dsTotalNeto / val(txtDivision.Caption), "###,###,###,##0.00")
            txtImpuesto2.Caption = Format(dsImpuesto2 / val(txtDivision.Caption), "###,###,###,##0.00")
            txtImpuesto1.Caption = Format(dsImpuesto1 / val(txtDivision.Caption), "###,###,###,##0.00")
            TxtRecargo.Caption = Format(dsRecargo / val(txtDivision.Caption), "###,###,###,##0.00")
            TxtDescuento.Caption = Format(dsDescuento / val(txtDivision.Caption), "###,###,###,##0.00")
            txtSubTotal.Caption = Format(dsSubTotal / val(txtDivision.Caption), "###,###,###,##0.00")
        
        Case Is = 2
            txtDivision.Caption = Format(IIf(val(txtDivision.Caption) = 1, "1", val(txtDivision.Caption) - 1), "###")
            txtTotalNeto.Caption = Format(dsTotalNeto / val(txtDivision.Caption), "###,###,###,##0.00")
            txtImpuesto2.Caption = Format(dsImpuesto2 / val(txtDivision.Caption), "###,###,###,##0.00")
            txtImpuesto1.Caption = Format(dsImpuesto1 / val(txtDivision.Caption), "###,###,###,##0.00")
            TxtRecargo.Caption = Format(dsRecargo / val(txtDivision.Caption), "###,###,###,##0.00")
            TxtDescuento.Caption = Format(dsDescuento / val(txtDivision.Caption), "###,###,###,##0.00")
            txtSubTotal.Caption = Format(dsSubTotal / val(txtDivision.Caption), "###,###,###,##0.00")
    End Select
    Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "cmddivision_Click", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Private Sub cmdMercadoPago_Click()
On Error GoTo fin

If (Calcular("select count(*) as codigo from ttarjetacredito where tCodigoIntegracion = 5", Cn) = 0) Then
    MsgBox "No hay tarjeta configurada para la integraciòn, favor de verificar las configuraciones de tarjetas!!!", vbInformation
    Exit Sub
End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    If (EjecutaMotorIntegraciones(sDocumento, txtSaldo(0).Caption, 5, 4) = True) Then
    
        Me.frmMensaje.Visible = True
        Me.lblMensajetexto.Caption = "Esperando confirmacion de pago por la Integraciòn con Mercado Pago" & vbNewLine & vbNewLine & "TOTAL:" & Me.txtSaldo(0).Caption
        Me.Image1.Visible = False
        Me.Image2.Visible = True
        Me.cmdCancelarMercadopago.Visible = True
    
        Frame1.Enabled = False
        Me.fraEfectivo.Enabled = False
    
    
        Me.Timer_MercadoPago.Enabled = True
        Me.Timer_MercadoPago.Interval = 1000
        Me.Timer_MercadoPago.Enabled = True
    Else
        MsgBox "No se pudo realizar el envio del pago a Mercado Pago!!!", vbCritical
        DoEvents
        Me.frmMensaje.Visible = False
        Frame1.Enabled = True
        Me.fraEfectivo.Enabled = True

    End If
    Screen.MousePointer = vbDefault
Exit Sub
fin:
    MsgBox "error: " & Error
    DoEvents
    Me.frmMensaje.Visible = False
    Frame1.Enabled = True
    Me.fraEfectivo.Enabled = True
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
On Error GoTo fin
   Select Case Index
          Case Is = 0  ' Primero
               MoverPuntero Primero, grdGrilla
               AsignarTotal
          
          Case Is = 1  ' Previo
               MoverPuntero previo, grdGrilla
               AsignarTotal
          
          Case Is = 2  ' Siguiente
               MoverPuntero siguiente, grdGrilla
               AsignarTotal
          
          Case Is = 3  ' Ultimo
               MoverPuntero Ultimo, grdGrilla
               AsignarTotal
               
               
               
               '--------- recibo de ingreso
          Case Is = 4  ' Primero
               MoverPuntero Primero, grdGrillaRecibo
               'AsignarTotal
          
          Case Is = 5  ' Previo
               MoverPuntero previo, grdGrillaRecibo
               'AsignarTotal
          
          Case Is = 6  ' Siguiente
               MoverPuntero siguiente, grdGrillaRecibo
               'AsignarTotal
          
          Case Is = 7  ' Ultimo
               MoverPuntero Ultimo, grdGrillaRecibo
               'AsignarTotal
   
   End Select
   Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "cmdNavegar_Click", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Private Sub cmdOpcion_Click(Index As Integer)
On Error GoTo fin
   Select Case Index
          Case Is = 0  ' Salir
               If (EstoyDivisionDoc = "1") Then
               
                Dim xSuma As String
                'xSuma = Calcular("select count(*) as Codigo FROM MDOCUMENTO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn)
                xSuma = Calcular("select count(*) as Codigo from MDOCUMENTO where tEstadoDocumento = '01' and tDocumento  in ( select distinct tDocumento from DDOCUMENTO where tCodigoPedido='" & sPedido & "')", Cn)
                  If xSuma > 0 Then
                      If MsgBox("Aun se tiene pendiente por pagar " + xSuma + " Documentos, esta seguro que desea salir del proceso?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                            Exit Sub
                        Else
                            EstoyDivisionDoc = "0"
                            wEnter = False
                            Unload Me
                      End If
                  End If
                Else
                     wEnter = False
                     Unload Me
               End If
          Case Is = 99  ' Salir
               wEnter = True
               Unload Me
          Case Is = 1  ' Cancelar
               Limpia
               txtNT.Caption = ""
               txtVT.Caption = ""
               txtMT.Caption = ""
               txtPT.Caption = ""
               sRefer = ""
               txtDocumento.Caption = ""
               txtReferencia.Caption = ""
               txtDocMonto.Caption = ""
               Cn.Execute "delete from " & sTarjeta & " where lPinpad=0"
               RsGrilla.Requery
               Bloquear False
               If RsGrilla.RecordCount > 0 Then
                  nAbonoN = nAbonoN + Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)
                  nPropinaN = nPropinaN + Calcular("select sum(nPropina) as Codigo from " & sTarjeta, Cn)
                  CalculaTotal
                  txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
                  txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
               End If
          
          Case Is = 2  ' Grabar
               'Tipo de Cambio
               Dim nCorrelativo As Integer
                                 
               If nTC = 0 Then
                  MsgBox "Tipo de Cambio no ingresado", vbCritical, sMensaje
                  Exit Sub
               End If
               
               If lFechaEntregaDelivery And txtFechaDelivery = "" Then
                  MsgBox "La hora de Entrega Delviery no fue Ingresada", vbCritical, sMensaje
                  Exit Sub
               End If
               
               If nCargo < nEN And (nCargo / nTC) < nEE Then
                  MsgBox "Ambos Efectivos no deben ser mayor al cargo", vbExclamation, sMensaje
                  Exit Sub
               End If
               
                '   If nCargo < nAbonoN Then
                '      MsgBox "El valor del Pago no puede ser mayor al Cargo ", vbExclamation, sMensaje
                '      Exit Sub
                '   End If
               
               If nSaldo > 0 And sCortesia = "" Then
                  MsgBox "Debe Cancelarse el total del documento", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               'If nCargo < Val(Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)) Then
               If Round(nCargo, 2) < Round(val(Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)), 2) Then
                  MsgBox "El abono(s) de las tarjeta(s) no puede ser mayor al total del documento", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If txtMCH(0).Caption > nCargo Then
                  MsgBox "El monto total en Cheque/Deposito no puede ser mayor al total del documento", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If IIf(txtDocMonto.Caption = "", 0, txtDocMonto.Caption) > nCargo Then
                  MsgBox "El monto cargado no puede ser mayor al total del documento", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If txtME(0).Caption = 0 And txtME(1).Caption = 0 Then
                 If txtAbono(0).Caption > txtCargo(0).Caption Then
                    MsgBox "El monto abonado no puede ser mayor al total del documento cuando no se cuenta con pago en efectivo", vbExclamation, sMensaje
                    Exit Sub
                 End If
               End If
               
               'jsjr 07-07-2025 -------------------------------------------
               If RsGrilla.RecordCount > 0 And sCortesia = "" Then
                  grdGrilla.MoveFirst
                  Dim X As Integer
                  For X = 1 To RsGrilla.RecordCount
                  
                        Set RsOperationTerminal = Lib.OpenRecordset("select tCodigoIntegracion from TTARJETACREDITO where tCodigoTarjeta = '" & RsGrilla!tTipo & "'", Cn)
                        
                        If RsOperationTerminal!tCodigoIntegracion = "19" Then
                        
                               TerminalSeleccionado = ""
                                
                               Set RsOperationCaja = Lib.OpenRecordset("select ISNULL(lActivaIntegracionNiubiz,0) as Valor from TCAJA where tCaja = '" & sCaja & "'", Cn)
                               
                               If Not RsOperationCaja.RecordCount = 0 Then
                                  If RsOperationCaja!Valor Then
                                  
                                     If Calcular("Select count(*) as Codigo from TTERMINAL where tIntegracion = '01'", Cn) > 0 Then
                                       frmPagoPinPad.Show vbModal
                                     End If
                                    
                                     If TerminalSeleccionado = "" Then
                                        MsgBox "No se seleccionó una terminal de envío. Verifique la configuración.", vbExclamation, sMensaje
                                        Exit Sub
                                     End If
                                  Else
                                     MsgBox "La caja actual no tiene habilitada la integración con la terminal de Niubiz.", vbExclamation, sMensaje
                                     Exit Sub
                                  End If
                               End If
                               
                               
                               If (IIf(IsNull(RsGrilla!nMonto), 0, RsGrilla!nMonto) + IIf(IsNull(RsGrilla!nPropina), 0, RsGrilla!nPropina)) < 1 Then
                                    MsgBox "El monto ingresado es menor al permitido por el terminal de Niubiz ( Minimo S/1.0 ).", vbExclamation, sMensaje
                                    Exit Sub
                               End If
                                       
                                       
                               Isql = "INSERT INTO DPAGOPOST (Integracion,IdTerminal,tDocumento,tTurno,tEstadoDocumento,nMonto,tUsuario,tCaja,FechaRegistro)" & _
                                      "VALUES ('NIUBIZ','" & TerminalSeleccionado & "','" & sDocumento & "','" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "','01','" & (IIf(IsNull(RsGrilla!nMonto), 0, RsGrilla!nMonto) + IIf(IsNull(RsGrilla!nPropina), 0, RsGrilla!nPropina)) & "','" & sUsuario & "','" & sCaja & "',GETDATE())"
                               Cn.Execute Isql
                               
                               Set RsOperationTerminal = Lib.OpenRecordset("execute SP_EJECUTA_MOTOR_INALAMBRICO '" & "COBRO," & sTurno & "," & sDocumento & "'", Cn)
                               If RsOperationTerminal!Valor = "1" Then
                                    MsgBox "Se ha enviado una solicitud de cobro al terminal (NIUBIZ). Por favor, verifique.", vbExclamation, sMensaje
                               Else
                                    Isql = "DELETE FROM DPAGOPOST WHERE Integracion ='NIUBIZ' AND tTurno='" & sTurno & "' AND tDocumento='" & sDocumento & "'"
                                    Cn.Execute Isql
                                    
                                    MsgBox "No se pudo enviar la solicitud de cobro al terminal (NIUBIZ). Por favor, verifique su configuracion.", vbCritical, sMensaje
                                    
                                    Exit Sub
                               End If
                        
                        
                        'Integracion Izipay
                        ElseIf RsOperationTerminal!tCodigoIntegracion = "20" Then
                        
                            Set RsOperationCaja = Lib.OpenRecordset("select ISNULL(lActivaIntegracionIzipay,0) as Valor from TCAJA where tCaja = '" & sCaja & "'", Cn)
                            
                            If Not RsOperationCaja.RecordCount = 0 Then
                                  If RsOperationCaja!Valor Then
                                  
                                     Set RsOperationCaja = Lib.OpenRecordset("select top 1 TCAJATERMINAL.* from TCAJATERMINAL inner join TTERMINAL on TCAJATERMINAL.tTerminal = TTERMINAL.Codigo where TTERMINAL.tIntegracion = '02' and TCAJATERMINAL.tCaja = '" & sCaja & "' and TCAJATERMINAL.lActivo = 1", Cn)
                                     
                                     If RsOperationCaja.RecordCount = 0 Then
                                        MsgBox "La caja actual no tiene un terminal asignado o habilitado. Verifique la configuración en el mantenimiento de terminales.", vbExclamation, sMensaje
                                        Exit Sub
                                     End If
                                     
                                     
                                  Else
                                     MsgBox "La caja actual no tiene habilitada la integración con la terminal de Izipay.", vbExclamation, sMensaje
                                     Exit Sub
                                  End If
                            End If
                            
                            If (IIf(IsNull(RsGrilla!nMonto), 0, RsGrilla!nMonto) + IIf(IsNull(RsGrilla!nPropina), 0, RsGrilla!nPropina)) < 1 Then
                                    MsgBox "El monto ingresado es menor al permitido por el terminal de Izipay ( Minimo S/1.0 ).", vbExclamation, sMensaje
                                    Exit Sub
                               End If
                                       
                                       
                               Isql = "INSERT INTO DPAGOPOST (Integracion,IdTerminal,tDocumento,tTurno,tEstadoDocumento,nMonto,tUsuario,tCaja,FechaRegistro)" & _
                                      "VALUES ('IZIPAY','" & RsOperationCaja!tTerminal & "','" & sDocumento & "','" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "','01','" & (IIf(IsNull(RsGrilla!nMonto), 0, RsGrilla!nMonto) + IIf(IsNull(RsGrilla!nPropina), 0, RsGrilla!nPropina)) & "','" & sUsuario & "','" & sCaja & "',GETDATE())"
                               Cn.Execute Isql
                               '------------------------------------------------------------------------
                               Dim sComando As String
                                Dim sResultado As String
                                Dim oShell As Object
                                Dim oExec As Object
                                Dim sLinea As String
                                Dim sPathExe As String
                               
                               sPathExe = "C:\INFOMATICA\MotorPagoInalambrico\ServicioPagoInalambrico.exe"
                                sComando = "COBRO," & sTurno & "," & sDocumento
                                
                                On Error Resume Next
                                Set oShell = CreateObject("WScript.Shell")
                                Set oExec = oShell.Exec("""" & sPathExe & """ " & sComando)
                                
                                ' Leer salida del EXE (si devuelve texto)
                                Do Until oExec.StdOut.AtEndOfStream
                                    sLinea = oExec.StdOut.ReadLine
                                    sResultado = sResultado & sLinea
                                Loop
                                
                                If Err.Number <> 0 Then
                                    MsgBox "Error al ejecutar el servicio del PINPAD: " & Err.Description, vbCritical, sMensaje
                                    Err.Clear
                                End If
                                
                                ' --- Validar resultado ---
                                If InStr(1, sResultado, "1") > 0 Then
                                    MsgBox "Pago Realizado en el terminal (Izipay).", vbInformation, sMensaje
                                    
                                    Dim Cmd As Object
                                    Set Cmd = CreateObject("ADODB.Command")
                                
                                    With Cmd
                                        .ActiveConnection = Cn
                                        .CommandType = adCmdStoredProc
                                        .CommandText = "SP_InsertPagoIntegracion"
                                
                                        .Parameters.Append .CreateParameter("@documento", adVarChar, adParamInput, 50, sDocumento)
                                        .Parameters.Append .CreateParameter("@estado", adBoolean, adParamInput, , False)
                                        .Parameters.Append .CreateParameter("@fecha", adDate, adParamInput, , Now)
                                        .Parameters.Append .CreateParameter("@tipo", adInteger, adParamInput, , 7)
                                
                                        .Execute
                                    End With
                                
                                    Set Cmd = Nothing
                                Else
                                    Isql = "DELETE FROM DPAGOPOST WHERE Integracion ='IZIPAY' AND tTurno='" & sTurno & "' AND tDocumento='" & sDocumento & "'"
                                    Cn.Execute Isql
                                    MsgBox "No se pudo enviar la solicitud de cobro al terminal (Izipay). Verifique configuración.", vbCritical, sMensaje
                                    Exit Sub
                                End If
                                
                                ' Limpieza
                                Set oExec = Nothing
                                Set oShell = Nothing
                               '----------------------------------------------------------
'                               Set RsOperationTerminal = Lib.OpenRecordset("execute SP_EJECUTA_MOTOR_INALAMBRICO '" & "COBRO," & sTurno & "," & sDocumento & "'", Cn)
'                               If RsOperationTerminal!Valor = "1" Then
'                                    MsgBox "Se ha enviado una solicitud de cobro al terminal (Izipay). Por favor, verifique.", vbExclamation, sMensaje
'                               Else
'                                    Isql = "DELETE FROM DPAGOPOST WHERE Integracion ='IZIPAY' AND tTurno='" & sTurno & "' AND tDocumento='" & sDocumento & "'"
'                                    Cn.Execute Isql
'
'                                    MsgBox "No se pudo enviar la solicitud de cobro al terminal (Izipay). Por favor, verifique su configuracion.", vbCritical, sMensaje
'                                    Exit Sub
'                               End If

                        End If
                      grdGrilla.MoveNext
                  Next X
               End If
                  
              
            '---------------------------------------------------------------
               
               
               
            If lActivarValesConsumo Then

               Dim sCOD_Vale_Temp As String
               sCOD_Vale_Temp = Calcular("SELECT TOP 1 tCOD_VALE AS Codigo FROM VALE_CONSUMO WHERE tCOD_VALE IN (SELECT tCOD_VALE FROM " & sTarjetaVC_Final & ") AND lESTADO_USO = 1 AND CONVERT(VARCHAR(20),fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) ", Cn)
               If sCOD_Vale_Temp <> "0" Then
                  MsgBox "El nro Vale " & sCOD_Vale_Temp & " ya se ha usado en otro comprobante.", vbInformation, sMensaje
                  Exit Sub
               End If

            End If

                ' Dim nCorrelativo2 As Integer
                ' nCorrelativo2 = 1
                ' If RsGrilla.RecordCount > 0 And sCortesia = "" Then
                '     grdGrilla.MoveFirst
                '     Dim X As Integer
                '     For X = 1 To RsGrilla.RecordCount
                '         Isql = "insert into " & "DPAGODOCUMENTO_VC" & _
                '                 "( tDocumento, tCorrelativo, tCod_Vale, nMonto_Vale) " & _
                '                 " SELECT '" & sDocumento & "','" & nCorrelativo2 & "'," & _
                '                          " tCOD_VALE," & _
                '                          " nMONTO_VALE " & _
                '                  " FROM " & "TMP_PRUEBA" & " WHERE tIndex = '" & RsGrilla!tIndex & "'" 'TMP_PRUEBA
                '         Cn.Execute Isql

                '         Isql = "UPDATE VC " _
                '                     & "SET VC.lESTADO_USO = 1 " _
                '                     & "FROM VALE_CONSUMO VC " _
                '                     & "INNER JOIN TMP_PRUEBA TMP ON VC.tCOD_VALE = TMP.tCOD_VALE " _
                '                     & "WHERE TMP.tIndex = '" & RsGrilla!tIndex & "'"
                '         Cn.Execute Isql

                '         nCorrelativo2 = nCorrelativo2 + 1
                '         grdGrilla.MoveNext
                '     Next X
                ' End If

                ' MsgBox "Prueba", vbExclamation, sMensaje
                ' Exit Sub
                             
               'JL Correccion
               If lMCPV Then
                    If MultiCajeroOk = False Then
                        Exit Sub
                    End If
               End If
                                            
               Screen.MousePointer = vbHourglass
               wEnter = True
               
               If xTipo = "Modificacion" Then
                 Dim fFechaPago As Date
                 Dim FDIAC As Date
                 Dim fActual As Date
                 Dim flag As String
                 'aqui se produce el error
                 flag = Calcular("select max(fRegistro) as Codigo from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'", Cn)
                 fActual = Calcular("select getdate() as Codigo    ", Cn)
                 fFechaPago = IIf(flag = "0", Format(fActual, "yyyy/mm/dd hh:nn"), flag)
                 FDIAC = Format(Calcular("select isnull(max(fDiaContable),0) as Codigo from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'", Cn), "yyyy/mm/dd hh:nn")
                    'XX AGOSTO 2011 LG ACTIVACION DEL INGRESO SI ES PAGADO POR RECIBO
                    Dim tTipoPago As String
                    Dim tnumerorecibo As String
                    tTipoPago = Calcular("SELECT ISNULL(TOTROTIPOPAGO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO='" & sDocumento & "' and totrotipopago='001'", Cn)
                    If tTipoPago = "001" Then
                            tnumerorecibo = Calcular("SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & sDocumento & "' and TOTROTIPOPAGO='001' ", Cn)
                            Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO in (SELECT ISNULL(TNUMERO,'') AS CODIGO FROM DPAGODOCUMENTO WHERE TDOCUMENTO ='" & sDocumento & "' and TOTROTIPOPAGO='001' )"
                            'Cn.Execute "UPDATE MINGRESO SET TESTADODOCUMENTO='01' WHERE TRECIBO='" & tnumerorecibo & "'"
                    End If
                    'FIN XX
                    '-- Recuperacion de la nota de credito cuando se Anula el pago del documento
                    Dim tnotacreditorecuperar As String
                    tnotacreditorecuperar = Calcular("select isnull(tNumero,'') as codigo from DPAGODOCUMENTO where  tDocumento in ('" & sDocumento & "') and totrotipopago='002'", Cn)
                    If tnotacreditorecuperar <> "0" Then
                        Cn.Execute "update MNOTACREDITO SET tEstadoDocumento='05' where tNotaCredito = '" & tnotacreditorecuperar & "' and  tEstadoDocumento<>'04'"
                    End If
                    '-- Fin Recuperacion de nota de credito

                    Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'"
                    
                    Cn.Execute "UPDATE VC " _
                                & "SET VC.lESTADO_USO = 0 " _
                                & "FROM DPAGODOCUMENTO_VC PVC " _
                                & "INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                & "WHERE PVC.tDocumento = '" & sDocumento & "' AND " _
                                & "CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "

                    Cn.Execute "DELETE PVC " _
                                & " FROM DPAGODOCUMENTO_VC PVC " _
                                & " INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                & " WHERE PVC.tDocumento = '" & sDocumento & "' AND " _
                                & " CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "
               End If
               '09102020
               If sCortesia = "" Then
                  Isql = "Update MDOCUMENTO set " & _
                         "tEstadoDocumento = '02', " & _
                         "nAbono = " & nAbonoN & ", " & _
                         IIf(xTipo = "Modificacion", "", "fPago = getdate(),") & _
                         "nVuelto = " & nVueltoN & " " & _
                         ",lreplica=1, TCORTESIA=''  where tDocumento ='" & sDocumento & "'"
                         lPagocortesiaAI = False
               Else
                  Isql = "update MDOCUMENTO set " & _
                         "tEstadoDocumento = '02', " & _
                         "tCortesia = '" & sCortesia & "', " & _
                         "nAbono = 0, " & _
                         IIf(xTipo = "Modificacion", "", "fPago = getdate(),") & _
                         "nVuelto = 0 " & _
                         ",lreplica=1 where tDocumento ='" & sDocumento & "'"
                         lPagocortesiaAI = True
               End If
               Cn.Execute Isql
               
               
               'aqui 09102020 GCAA
               
               Dim oComando As clsComando
               Set oComando = New clsComando

               Set oComando = New clsComando
                If Not oComando.CreateCmdSp("spUpdate_DDOCUMENTO_Ina", Cn) Then
                   Set oComando = Nothing
                   Exit Sub
                End If
                
                oComando.CreateParameter "@tdocumento", adVarChar, adParamInput, 20, sDocumento
                
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Sub
                End If
            
                If Not oComando.ExecSP Then
                   Set oComando = Nothing
                   Exit Sub
                 End If
               
               
               
               'fin 09102020 GCAA
               
               nCorrelativo = 1
               Dim nEfectivo As Double
               
               Cn.Execute "delete from DPREPAGO where tDocumento='" & sDocumento & "'"
               
               'Efectivo Moneda Nacional
               nEfectivo = 0
               If nEN > 0 Then
                  If nVueltoN > 0 Then
                     nEfectivo = nEN - nVueltoN
                     If nEfectivo < 0 Then
                        nEfectivo = nEN
                     End If
                  Else
                     nEfectivo = nEN
                  End If
                '09102020
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, fRegistro, tUsuario,fDiaContable, nDolar, tCodigoIntegracion,tReferencia ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'01'," _
                                  & "'01'," _
                                  & nTC & ", " _
                                  & nEfectivo & ", " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "', " & nEN & "," & tCodigoIntegracion & ",'" & IIf(tCodigoIntegracion = 4, "CASH", "") & "')"
                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               
               'Efectivo Moneda Extranjera--------------------------------------------------------------
               nEfectivo = 0
               If nEE > 0 Then
                  If nVueltoE > 0 And nEE > 0 And nEN < nVueltoN Then
                     nEfectivo = nEE - nVueltoE
                     If nEfectivo < 0 Then
                        nEfectivo = 0
                     End If
                  Else
                     nEfectivo = nEE
                  End If
                  
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, nDolar, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'01'," _
                                  & "'02'," _
                                  & nTC & ", " _
                                  & nEfectivo & ", " _
                                  & nEE & ", " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"

                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               '---------------------------------------------------------------------------------------------
                'Efectivo Moneda Extranjera 2 -------------------------------------------------------------
               nEfectivo = 0
               If nEE2 > 0 Then
                  If nVueltoE2 > 0 And nEE2 > 0 And nEN < nVueltoN Then
                     nEfectivo = nEE2 - nVueltoE2
                     If nEfectivo < 0 Then
                        nEfectivo = 0
                     End If
                  Else
                     nEfectivo = nEE2
                  End If
                  
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, nDolar, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'01'," _
                                  & "'03'," _
                                  & nTC2 & ", " _
                                  & nEfectivo & ", " _
                                  & nEE2 & ", " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"

                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               '---------------------------------------------------------------------------------------------
                'Efectivo Moneda Extranjera 3 -------------------------------------------------------------
               nEfectivo = 0
               If nEE3 > 0 Then
                  If nVueltoE3 > 0 And nEE3 > 0 And nEN < nVueltoN Then
                     nEfectivo = nEE3 - nVueltoE3
                     If nEfectivo < 0 Then
                        nEfectivo = 0
                     End If
                  Else
                     nEfectivo = nEE3
                  End If
                  
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, nDolar, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'01'," _
                                  & "'04'," _
                                  & nTC3 & ", " _
                                  & nEfectivo & ", " _
                                  & nEE3 & ", " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"

                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               '---------------------------------------------------------------------------------------------
               'Cheque
               If nCH <> 0 Then
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, tBanco, tNumero, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'03'," _
                                  & "'01'," _
                                  & nTC & ", " _
                                  & nCH & ", " _
                                  & "'" & txtBCH.Caption & "', " _
                                  & "'" & txtNCH.Caption & "', " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"

                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               
               'Cheque Dolares
               If nCE <> 0 Then
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, nDolar, tBanco, tNumero, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'03'," _
                                  & "'02'," _
                                  & nTC & ", " _
                                  & nCE & ", " _
                                  & nCE & ", " _
                                  & "'" & txtBCH.Caption & "', " _
                                  & "'" & txtNCH.Caption & "', " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"

                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               
               'Tarjeta
               If RsGrilla.RecordCount > 0 And sCortesia = "" Then
                  grdGrilla.MoveFirst
                  Dim i As Integer
                  For i = 1 To RsGrilla.RecordCount
                       If RsGrilla!lPinPad Then
                            If Not InsertaSiab(RsGrilla!tNumero, RsGrilla!nMonto) Then
                               Screen.MousePointer = vbDefault
                               MsgBox "Error Faltal: Conexión Siab perdida, consultar con Sistemas", vbCritical, sMensaje
                               Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'"
                                Cn.Execute "UPDATE VC " _
                                            & "SET VC.lESTADO_USO = 0 " _
                                            & "FROM DPAGODOCUMENTO_VC PVC " _
                                            & "INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                            & "WHERE PVC.tDocumento = '" & sDocumento & "' AND " _
                                            & "CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "
            
                                Cn.Execute "DELETE PVC " _
                                            & " FROM DPAGODOCUMENTO_VC PVC " _
                                            & " INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                            & " WHERE PVC.tDocumento = '" & sDocumento & "' AND " _
                                            & " CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "
                                            
                               Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '01', nAbono = 0, fPago = null, nVuelto = 0, lreplica=0 where tDocumento ='" & sDocumento & "'"
                               Exit Sub
                            End If
                       End If
                       
                       
                    'jsjr 07-07-2025 -----------------------------------------
                    Set RsOperationTerminal = Lib.OpenRecordset("select tCodigoIntegracion from TTARJETACREDITO where tCodigoTarjeta = '" & RsGrilla!tTipo & "'", Cn)
                    Set RsOperationCaja = Lib.OpenRecordset("select ISNULL(lActivaIntegracionNiubiz,0) as Valor from TCAJA where tCaja = '" & sCaja & "'", Cn)
                    If RsOperationCaja!Valor = 0 Then
                    Set RsOperationCaja = Lib.OpenRecordset("select ISNULL(lActivaIntegracionIzipay,0) as Valor from TCAJA where tCaja = '" & sCaja & "'", Cn)
                    End If
                       
                    If RsOperationTerminal!tCodigoIntegracion = "19" And RsOperationCaja!Valor Then
                         
                         Isql = "UPDATE MDOCUMENTO SET tEstadoDocumento='01' WHERE tDocumento = '" & sDocumento & "'"
                         Cn.Execute Isql
                         
                         Isql = "INSERT INTO DPREPAGO (tDocumento,tCorrelativo,tTurno,tTipoPago,tMoneda,nTipoCambio,nMonto,nPropina,nVuelto,tTarjeta,tReferencia,tUsuario,fDiaContable,fRegistro,tCodigoPedido) " & _
                                      "VALUES ('" & sDocumento & "','" & Trim(str(nCorrelativo)) & "','" & sTurno & "','02','01','" & nTC & "','" & IIf(IsNull(RsGrilla!nMonto), 0, RsGrilla!nMonto) & "','" & IIf(IsNull(RsGrilla!nPropina), 0, RsGrilla!nPropina) & "',0,'" & RsGrilla!tTipo & "','NBZ','" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "',GETDATE(),'" & sPedido & "')"
                         Cn.Execute Isql
                         
                    Else
                    
                         Isql = "insert into DPAGODOCUMENTO " & _
                             "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio, nMonto, npropina, tTarjeta, tNumero, tFechaVencimiento, fRegistro, tUsuario,fDiaContable,tNumeroATarjeta ) " & _
                             "Values(    '" & sDocumento & "'," _
                                      & "'" & Trim(str(nCorrelativo)) & "'," _
                                      & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                      & "'02'," _
                                      & "'01'," _
                                      & "'" & IIf(IsNull(RsGrilla!tReferencia), "", RsGrilla!tReferencia) & "', " _
                                      & nTC & ", " _
                                      & IIf(IsNull(RsGrilla!nMonto), 0, RsGrilla!nMonto) & ", " _
                                      & IIf(IsNull(RsGrilla!nPropina), 0, RsGrilla!nPropina) & ", " _
                                      & "'" & RsGrilla!tTipo & "', " _
                                      & "'" & IIf(IsNull(RsGrilla!tNumero), "", RsGrilla!tNumero) & "', " _
                                      & "'" & IIf(IsNull(RsGrilla!tFecha), "", RsGrilla!tFecha) & "', " _
                                      & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                      & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "', '" & IIf(IsNull(RsGrilla!tNumeroATarjeta), "", RsGrilla!tNumeroATarjeta) & "')"
                  
                      Cn.Execute Isql
                    End If
                    '--------------------------------------------------
                       
                       
                    If lActivarValesConsumo Then
                        Isql = "insert into " & "DPAGODOCUMENTO_VC" & _
                                "( tDocumento, tCorrelativo, tCod_Vale, nMonto_Vale) " & _
                                " SELECT '" & sDocumento & "','" & nCorrelativo & "'," & _
                                         " tCOD_VALE," & _
                                         " nMONTO_VALE " & _
                                 " FROM " & sTarjetaVC_Final & " WHERE tIndex = '" & RsGrilla!tIndex & "'"
                        Cn.Execute Isql

                        Isql = "UPDATE VC " _
                                    & "SET VC.lESTADO_USO = 1 " _
                                    & "FROM VALE_CONSUMO VC " _
                                    & "INNER JOIN " & sTarjetaVC_Final & " TMP ON VC.tCOD_VALE = TMP.tCOD_VALE " _
                                    & "WHERE TMP.tIndex = '" & RsGrilla!tIndex & "'"
                        Cn.Execute Isql
                    End If
                      nCorrelativo = nCorrelativo + 1
                      grdGrilla.MoveNext
                      
                  Next i
               End If

               'Recibo de Ingreso
               If RsGrillaRecibido.RecordCount > 0 Then
                  grdGrillaRecibo.MoveFirst
                  
                  Dim ii As Integer
                  Dim Moneda As String
                  Dim sOtroTipoCancelacionrecibo As String
                   sOtroTipoCancelacionrecibo = "001"
                  
                  For ii = 1 To RsGrillaRecibido.RecordCount
                                     
                    Moneda = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & RsGrillaRecibido!codigo & "'", Cn)
                       
                     Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'04'," _
                                  & "'" & sOtroTipoCancelacionrecibo & "'," _
                                  & "'" & Moneda & "'," _
                                  & nTC & ", " _
                                    & RsGrillaRecibido!nMonto & ", " _
                                  & "'" & RsGrillaRecibido!codigo & "', " _
                                  & "'" & IIf(sOtroTipoCancelacionrecibo = "000", Mid(txtDocumento.Caption, 1, 30), "") & "', " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"
                    Cn.Execute Isql
                    nCorrelativo = nCorrelativo + 1
                    Cn.Execute "update MINGRESO set tEstadoDocumento ='02' where tRecibo ='" & RsGrillaRecibido!codigo & "'"
                    
                    If lActivaReciboSaldoQr Then
                        Cn.Execute "insert into DINGRESO ( trecibo, nCorrelativo, monto, fRegistro, tUsuario, tCaja, tTurno, tDocumento) " & _
                                    " values ('" & RsGrillaRecibido!codigo & "', (select isnull(max(ncorrelativo),0)+1 from DINGRESO  where trecibo='" & RsGrillaRecibido!codigo & "'), " & _
                                    " " & RsGrillaRecibido!nMonto & ", getdate()," & _
                                    " '" & sUsuario & "','" & sCaja & "', '" & sTurno & "','" & sDocumento & "') "
                    End If
                    
                    grdGrillaRecibo.MoveNext
                  Next ii
               End If
               
               'Documento
               If nDocumento <> 0 Then
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'04'," _
                                  & "'" & sOtroTipoCancelacion & "'," _
                                  & "'" & sMonDoc & "'," _
                                  & nTC & ", " _
                                    & nDocumento & ", " _
                                  & "'" & sTipoDocumento & "', " _
                                  & "'" & IIf(sOtroTipoCancelacion = "000", Mid(txtDocumento.Caption, 1, 30), "") & "', " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"
                  Cn.Execute Isql
                  If sOtroTipoCancelacion = "001" Then
                     Cn.Execute "update MINGRESO set tEstadoDocumento ='02' where tRecibo ='" & sTipoDocumento & "'"
                  ElseIf sOtroTipoCancelacion = "002" Then
                     Cn.Execute "update MNOTACREDITO set tEstadoDocumento ='02',lreplica=1 where tNotaCredito ='" & sTipoDocumento & "'"
                  End If
                  nCorrelativo = nCorrelativo + 1
               End If
                        
               'Puntos
               If nTotalPuntos <> 0 Then
                  Isql = "insert into DPAGODOCUMENTO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, fRegistro, tUsuario,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & IIf(xTipo = "Modificacion", sTurnoModificacion, sTurno) & "'," _
                                  & "'05'," _
                                  & "''," _
                                  & "'" & sMonDoc & "'," _
                                  & nTC & ", " _
                                  & nTotalPuntos & ", " _
                                  & "'" & sClientePuntos & "', " _
                                  & "'', " _
                                  & IIf(xTipo = "Modificacion", " '" & Format(fFechaPago, "yyyy/mm/dd HH:nn") & "', ", "getdate(),") _
                                  & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "')"
                  Cn.Execute Isql
                  Cn.Execute "update TDELIVERY set nUtilizado=nUtilizado+" & nTotalPuntos & ", nDisponible = nDisponible - " & nTotalPuntos & " where tCodigoDelivery='" & sClientePuntos & "'"
                  nCorrelativo = nCorrelativo + 1
               Else
                  sClientePuntos = Calcular("select MAX(tClienteDelivery) as Codigo FROM dbo.DPEDIDO INNER JOIN dbo.MPEDIDO ON dbo.DPEDIDO.tCodigoPedido = dbo.MPEDIDO.tCodigoPedido Where tDocumento='" & sDocumento & "'", Cn)
                  If sCliente <> "0" Then
                     Cn.Execute "update TDELIVERY set nAcumulado=nAcumulado+" & nCargo / nPunto & ", nDisponible = nDisponible + " & nCargo / nPunto & " where tCodigoDelivery='" & sClientePuntos & "'"
                  End If
               End If
                              
               'Enlace Infhotel cambio estado comanda 03
               If lInfhotel Then
                    Dim linfhotelPedido As String
                    Dim linfhotelcomanda As String
                    linfhotelPedido = Calcular("select tCodigoPedido as Codigo from DDOCUMENTO where tDocumento='" & sDocumento & "'", Cn)
                    linfhotelcomanda = Calcular("select tComanda as Codigo from MPEDIDO where tCodigoPedido='" & linfhotelPedido & "'", Cn)
               
                    CnInfhotel.Execute "update MCOMANDA set tEstado='03' " & _
                                       "where tComanda ='" & linfhotelcomanda & "' and tPuntoVenta='" & sPuntoVentaInfhotel & "'"
               End If
                                                
               'Liberacion
               If xTipo = "" Then
                  Cn.Execute "Update dbo.DPEDIDO set tFacturado ='P', tCortesia='' where tDocumento ='" & sDocumento & "'"
                  Cn.Execute "UPDATE dbo.MPEDIDO set tEstadoPedido='02', lReplica=1 where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' )  AND TCODIGOPEDIDO NOT IN (SELECT DISTINCT TCODIGOPEDIDO FROM DPEDIDO WHERE TCODIGOPEDIDO IN(select DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO where tDocumento='" & sDocumento & "') AND ISNULL(TFACTURADO,'') <> 'P')"
                  
                  'GCAA 04032021 // ACTUALIZA EL CAMPO fEntregaClienteCD
                  'SOLO CUANDO ES UN PEDIDO TIPO 02 Y TIPO DE RECEPCION 1
                  Set RsSaberTipoPedido = Lib.OpenRecordset("select tTipoPedido ,tTiporecepcion  from MPEDIDO where tCodigoPedido in (select  tCodigoPedido  from DDOCUMENTO where tDocumento ='" & sDocumento & "')", Cn)
                 
                 If RsSaberTipoPedido!tTipoPedido = "02" And RsSaberTipoPedido!tTiporecepcion = 1 Then
                      Cn.Execute "UPDATE dbo.MPEDIDO set fEntregaClienteCD = getdate(), fLlegada = getdate(), fEntrega = {fn CURDATE()} + ' " & txtFechaDelivery & "' where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' ) and tTipoPedido='02' and isnull(fLlegada ,0)=0"
                 Else
                      Cn.Execute "UPDATE dbo.MPEDIDO set fLlegada = getdate(), fEntrega = {fn CURDATE()} + ' " & txtFechaDelivery & "' where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' ) and tTipoPedido='02' and isnull(fLlegada ,0)=0"
                 End If
                 
                  
                
                  
                  Cn.Execute "Update dbo.TMESA set tEstadoMesa = '04' where tCodigoMesa in (SELECT DISTINCT TMESA FROM MPEDIDO WHERE TCODIGOPEDIDO IN (SELECT DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO WHERE TDOCUMENTO='" & sDocumento & "') AND TCODIGOPEDIDO NOT IN (SELECT DISTINCT TCODIGOPEDIDO FROM DPEDIDO WHERE TCODIGOPEDIDO IN(select DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO where tDocumento='" & sDocumento & "') AND ISNULL(TFACTURADO,'') <> 'P'))"
                  'Juntar Mesa
                  Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
               End If
                              
               'Tarjeta de Credito vs Cliente
               If Calcular("select tTipoPedido as Codigo from DPEDIDO where tDocumento='" & sDocumento & "'", Cn) <> "" And RsGrilla.RecordCount > 0 Then
                  RsGrilla.MoveFirst
                  Cn.Execute "update TDELIVERY set tCodigoTarjeta='" & RsGrilla!tTipo & "', tFechaTarjeta='" & RsGrilla!tFecha & "', tNumeroTarjeta='" & IIf(IsNull(RsGrilla!tNumero), "", RsGrilla!tNumero) & "' where tCodigoDelivery='" & sCliente & "'"
               End If
               
               'Cuentas x Cobrar
               Dim clientePagoCP As String
               Dim totalDocumentoCP As Double
               Dim nConsumoActualCP As Double
               Dim totalCP As Double
               
               If xTipo = "PorPagar" Then
                  clientePagoCP = Calcular("Select tClientePago As Codigo From MDOCUMENTO Where tDocumento = '" & sDocumento & "'", Cn)
                  nConsumoActualCP = Calcular("Select isnull(nconsumoporcobrar,0) as codigo From TDELIVERY where tCodigoDelivery='" & clientePagoCP & "'", Cn)
                  totalDocumentoCP = Calcular("Select nVenta As Codigo From mdocumento where tdocumento = '" & sDocumento & "'", Cn)
                  totalCP = nConsumoActualCP - totalDocumentoCP
                  Isql = "Update TDELIVERY Set nConsumoPorCobrar = " & totalCP & " Where tCodigoDelivery='" & clientePagoCP & "'"
                  Cn.Execute Isql
               End If
               '------------------------------
               
              xTiporecepcion = Calcular("Select Distinct isnull( P.tTiporecepcion,0) as codigo from DDocumento as D Inner Join MPedido AS P On D.tCodigoPedido = P.tCodigoPedido Where tDocumento ='" & sDocumento & "'", Cn)
                    
               If xTiporecepcion <> 0 Then
  
                    If xTiporecepcion = 1 Then
                        Call ModifcarEstadoDeliveryCabecera(sDocumento)
                    ElseIf xTiporecepcion = 2 Then
                        'Call ModificarEstadoPedidoWebAp(sDocumento)
                    End If
                    
               End If
                              
                'GCAA 16032022 FE ARGENTINA
                 
                If pais = "003" And lFETusFacturasAPP = True And (Left(sDocumento, 1) = "A" Or Left(sDocumento, 1) = "B") Then
                      If lPagoAntesImpresion = False Then
                        If Not (INSERTA_FE_INFOREST_ARGE(sDocumento, 3, DateTime.Date)) Then
                          
                            MsgBox "Error de conexion con el servicio TUSFACTURASAPP, revisar el motor de facturacion, Integracion FE ", vbCritical, sMensaje
                            GoTo fin
                        End If
                        End If
                End If
                              
               Screen.MousePointer = vbDefault
               wEnter = True
               If lPagoAntesImpresion Then
                    lGrabaPagoAntes = True
               End If
  
               Unload Me
                              
          Case Is = 3  ' Cortesia
               If MsgBox("Seguro de Cancelarlo por Transferencia a Titulo Gratuito?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                              
               If Supervisor("04") = False Then
                  MsgBox "Clave no permitida", vbExclamation, sMensaje
                  Exit Sub
               End If
                                
               sTemp = ""
               Isql = "select * from vCortesia where lActivo = 1 Order by Descripcion"
               Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                               "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
                
                If wEnter = True Then
                    Limpia
                    sCortesia = sCodigo
                    txtCortesia.Caption = sDescrip
                    Bloquear True
                    
                    Dim nTotalCortesiaActual As Double
                    Dim nTopeCortesia As Double
                    Dim nTotalDocActual As Double
                    Dim nTope As Integer
                    
                    
                    'GCAA 02/02/2022
                    'REQ REST-C043-2021 - Grupo Hersil
                    nTope = Calcular("select isnull(Transferencia,0) as codigo from vcortesia where codigo='" & sCortesia & "'", Cn)
                    
                    If nTope = "1" Then
                        nTopeCortesia = Calcular("select isnull(tope,0) as codigo from vcortesia where codigo='" & sCortesia & "'", Cn)
                        If nTopeCortesia > 0 Then
                                'nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where tEstadoDocumento='02' and tcortesia='" & sCortesia & "' and month(fregistro)=month(getdate()) ", Cn)
                                nTotalCortesiaActual = Calcular("select sum(isnull(nventa,0))  as codigo From mDocumento where tEstadoDocumento='02' and tcortesia='" & sCortesia & "' and YEAR(FREGISTRO)=YEAR(GETDATE()) AND month(fregistro)=month(getdate()) ", Cn)
                                nTotalDocActual = val(txtCargo(0).Caption)
                                    If nTotalCortesiaActual + nTotalDocActual > nTopeCortesia Then
                                            MsgBox "Con esta Emisión se supera el Tope Mensual asignado para la Cortesia " & UCase(sDescrip) & vbCrLf & "Tope Mensual: " & nTopeCortesia & ". Ya Asignado : " & nTotalCortesiaActual, vbCritical
                                            Exit Sub
                                    End If
                        End If
                    End If
                                        
                Else
                   sCortesia = ""
                   txtCortesia.Caption = ""
                End If

                    Dim CODM As String
                    
                    CODM = Calcular("SELECT  DISTINCT TCODIGOPEDIDO AS CODIGO FROM DPEDIDO WHERE TDOCUMENTO='" & sDocumento & "'", Cn)
                 '---- mesa247 --- avisa a mesa que fue enviada a cuentas por cobrar
                If MESA247 Then
                    Cn.Execute "Update mpedido set lestadomesapedido = 5 where tcodigopedido ='" & CODM & "' and idpedidomesa<>'' "
                End If
                

                'Limpia
                txtNT.Caption = ""
                txtVT.Caption = ""
                txtMT.Caption = ""
                txtPT.Caption = ""
                sRefer = ""
                txtDocumento.Caption = ""
                txtReferencia.Caption = ""
                txtDocMonto.Caption = ""
                Cn.Execute "delete from " & sTarjeta
                RsGrilla.Requery
          
          Case Is = 4  ' Agregar Tarjeta
          
               If nET = 0 Then
                  MsgBox "El monto no puede ser cero", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If txtNT.Caption = "" And lActivaVoucher Then
                    MsgBox "Es necesario asignar el Nº de Voucher", vbExclamation, sMensaje
                  Exit Sub
               End If
                
               If lActivarValesConsumo Then
                Dim sCOD_Vale_VC As String
                sCOD_Vale_VC = Calcular("SELECT TOP 1 tCOD_VALE AS Codigo FROM VALE_CONSUMO WHERE tCOD_VALE IN (SELECT tCOD_VALE FROM " & sTarjetaVC & ") AND lESTADO_USO = 1 AND CONVERT(VARCHAR(20),fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) ", Cn)
                If sCOD_Vale_VC <> "0" Then
                   MsgBox "El nro Vale " & sCOD_Vale_VC & " ya se ha usado en otro comprobante.", vbInformation, sMensaje
                   Exit Sub
                End If
               End If

               lPinpadTarjeta = Calcular("select isnull(lpinpad,0) codigo from ttarjetacredito where tcodigotarjeta='" & sTipoTarjeta & "'", Cn)
                'Ecuador Tarjeta Socio
                If lSiab And lPinpadTarjeta Then
                    Mensaje "Estableciendo Conexion Esperando...", "SIAB", 10
                    Dim nMontoCalculado As Double
                    sSYBASE = Trim(LeerIni(App.path + "\INFOREST.INI", "CONEXION", "SYBASE", ""))
                    Set CnSiab = New Connection
                    CnSiab.Provider = SYPROV '"ASAProv.80"
                    CnSiab.CursorLocation = adUseServer
                    CnSiab.ConnectionString = sSYBASE
                    CnSiab.CommandTimeout = 250
                    On Error GoTo ErrorEcuador
                    CnSiab.Open
                    nMontoCalculado = Calcular("select saldo as Codigo from tarj_consumo_maestro where estado='A' and num_tarjeta = '" & txtNT.Caption & "'", CnSiab)
                    CnSiab.Close
                    On Error GoTo 0
                    If nMontoCalculado < nET Then
                       MsgBox "Tarjeta sin saldo suficiente", vbOKOnly + vbExclamation, sMensaje
                       Exit Sub
                    End If
                End If
                              
               nAbonoN = nAbonoN + nET
               nPropinaN = nPropinaN + nPT
               CalculaTotal
               nCorrela = Calcular("select max(tIndex) as Codigo from " & sTarjeta, Cn)
               If IsNull(nCorrela) Or nCorrela = "" Then
                  nCorrela = "001"
               Else
                  nCorrela = Lib.Correlativo(nCorrela, 3)
               End If

                If sCodigoIntegracion = "18" And lActivarValesConsumo Then 'Solo si está activado en ADM y enlazado a SERVILEXll
                    Isql = "insert into " & sTarjetaVC_Final _
                            & "( tIndex, tTICKET, tCOD_VALE, tMONEDA, nMONTO_VALE) " _
                            & " SELECT '" & nCorrela & "' ,tTICKET, " _
                            & " tCOD_VALE," _
                            & " tMONEDA," _
                            & " nMONTO_VALE " _
                            & " FROM " & sTarjetaVC
                    Cn.Execute Isql
                    Set RsGrillaVConsumo = Nothing
                    Cn.Execute "delete from " & sTarjetaVC & ""
               End If
               
               Isql = "insert into " & sTarjeta & _
                      "( tIndex, tTipo, tDescripcion, tNumero, tFecha, tReferencia, nMonto, nPropina, nIndex, lPinPad ) " & _
                      "Values(    '" & nCorrela & "', " _
                               & "'" & sTipoTarjeta & "'," _
                               & "'" & sTitulo & "'," _
                               & "'" & txtNT.Caption & "'," _
                               & "'" & txtVT.Caption & "'," _
                               & "''," _
                               & nET & ", " _
                               & nPT & ", " _
                               & nIndex & ", '" & IIf(lPinpadTarjeta = False, 0, 1) & "')"
               Cn.Execute Isql
               RsGrilla.Requery
               ActivaTarjeta False
               Bloquear False
               txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
               txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
               AsignarTotal
               txtNT.Caption = ""
               txtVT.Caption = ""
               txtMT.Caption = "0.00"
               txtPT.Caption = "0.00"
          
          Case Is = 5  ' Cancelar Tarjeta
               txtNT.Caption = ""
               txtVT.Caption = ""
               txtMT.Caption = "0.00"
               txtPT.Caption = "0.00"
               sRefer = ""
               ActivaTarjeta False
               Bloquear False
               lPinpadTarjeta = False
               lTarjetaGrilla = False
               
               
          Case Is = 6  ' Eliminar Tarjeta
               If RsGrilla.RecordCount <> 0 Then
                  If Not RsGrilla!lPinPad Then
                     nIndex = RsGrilla!nIndex
                     txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption - RsGrilla!nMonto, "###,###,##0.00")
                     txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption - RsGrilla!nPropina, "###,###,##0.00")
                     nAbonoN = nAbonoN - RsGrilla!nMonto
                     nPropinaN = nPropinaN - RsGrilla!nPropina
                     
                     Cn.Execute "delete from " & sTarjeta & " where tIndex = '" & RsGrilla!tIndex & "'"
                     
                     If lActivarValesConsumo Then
                        Cn.Execute "delete from " & sTarjetaVC_Final & " where tIndex = '" & RsGrilla!tIndex & "'"
                     End If
                     CalculaTotal
                     RsGrilla.Requery
                     lblTarjetaC.Caption = ""
                     AsignarTotal
                  Else
                    sMonto = Format(str(RsGrilla!nMonto), "0000000000.00")
                    sMonto = Mid(sMonto, 1, 10) & Mid(sMonto, 12, 2)
                    sOperacion = OP_FINANCIERA & "A" & sMonto & Chr$(FS) & _
                                                 "B" & "000000000000" & Chr$(FS) & _
                                                 "C" & "0" & Chr$(FS) & _
                                                 "D" & sEmpresa & Chr$(FS) & _
                                                 "E" & sCaja
                    nRet = fiStartOperation(sOperacion, 2, sMensaje)
                                                                       
                    If nRet = RET_OK Or nRet = RET_RUNNING Then
                       If Not Imprimir(sPreCuenta) Then
                          Exit Sub
                       End If
                       Printer.FontName = sFont
                       Printer.FontBold = False
                       sClave = ""
                       nContador = 0
                       lEmisor = True
                       lLoop = True
                       
                       Do
                         sRetorno = ""
                         nRet = fiGetStatus(sRetorno, 512)
                         lEmisor = ImprimeCabecera(sRetorno, lEmisor)
                         sClave = MensajePinPad(sRetorno)
                         If Mid(sClave, 1, 3) = "A00" Or Mid(sClave, 1, 3) = "A11" Then
                            sRefer = BuscaRetornoPinPad(sClave, "E")
                            nAbonoN = nAbonoN - Calcular("select sum(nMonto) as Codigo from  " & sTarjeta & " where tReferencia = '" & sRefer & "'", Cn)
                            nPropinaN = nPropinaN - Calcular("select sum(nPropina) as Codigo from  " & sTarjeta & " where tReferencia = '" & sRefer & "'", Cn)
                            Cn.Execute "delete from " & sTarjeta & " where tReferencia = '" & sRefer & "'"
                            Cn.Execute "update DPAGOTARJETA set tEstadoDocumento='04' where tDocumento='" & sDocumento & "' and tReferencia='" & sRefer & "'"
                            CalculaTotal
                            AsignarTotal
                            RsGrilla.Requery
                            txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
                            txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
                            AsignarTotal
                         Else
                            xError = BuscaRetornoPinPad(sClave, "B")
                            If Len(xError) > 0 Then
                               Mensaje xError, "VisaNet", 1000
                            End If
                         End If
                         
                         Mensaje "PinPad Listo. Esperando...", "PinPad", 500
                         nContador = nContador + 1
                         If nContador >= nTimeOut Then
                            If MsgBox("Tiempo de espera agotado, deseas mas tiempo?", vbExclamation + vbOKCancel, "VisaNet") = vbOK Then
                               lLoop = True
                               nContador = nTimeOut / 2
                            Else
                               lLoop = False
                            End If
                         End If
                          
                          If nRet <> "0" Then
                             nContador = 0
                          End If
                       Loop While (Mid$(sRetorno, 5, 2) <> "C1") And lLoop
                    Else
                       MsgBox "Error de conectividad", vbCritical, sMensaje
                       Exit Sub
                    End If
                  End If
               End If
                                  
          Case Is = 7
               sTipo = "Fecha"
               frmNumPad.Show vbModal
               If wEnter = True Then
                  If Len(sDescrip) = 4 And val(Mid(sDescrip, 1, 2)) <= 23 And val(Mid(sDescrip, 3, 2)) <= 59 Then
                     txtFechaDelivery.Caption = Format(sDescrip, "##:##")
                  Else
                     MsgBox "Formato de 24 Hrs. (HHMM)", vbInformation, sMensaje
                     txtFechaDelivery.Caption = ""
                  End If
               Else
                  txtFechaDelivery.Caption = ""
               End If
               
          Case Is = 8
                Dim antesvar As String
                antesvar = sVar1
                
                If lPasswordPorCobrar = True Then
                    If Supervisor("08") = False Then
                       MsgBox "Clave no permitida", vbExclamation, sMensaje
                       Exit Sub
                    End If
                Else
                    sVar1 = sUsuario
                End If
                
                sTemp = ""
                
                'GCAA 06122022 -  CORPORACION FERNADEZ
                '    If lActivaCreditoCoorporativo And lActivaRestriClienteDelivery Then
                '        Dim codigoDelivery As String

                '        codigoDelivery = Calcular("select isnull(tClienteDelivery,'') as codigo from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)


                '        Isql = "select * from vCompania where lActivo = 1 AND codigo = '" & codigoDelivery & "'  order by Descripcion "
                '    Else
                '        Isql = "select * from vCompania where lActivo = 1 order by Descripcion "
                '    End If
                
                Isql = "select * from vCompania where lActivo = 1 order by Descripcion "
                
                Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 900, 2, 0, "", _
                                                                "Identidad", 2, "Identidad", 1700, 0, 0, "", _
                                                                "Cliente", 2, "Descripcion", 3400, 0, 0, "", _
                                                                "Linea", 2, "nLineaPorCobrar", 1100, 1, 0, "##,##0.00", _
                                                                "Consumo", 2, "nConsumoPorCobrar", 1100, 1, 0, "##,##0.00")
                frmBusquedaRapida.nPredeterm = 1
                frmBusquedaRapida.Show vbModal
                
                If Not wEnter Then
                    wEnter = False
                   Exit Sub
                End If
                
                Dim nTotalPorCobrarCliente As Double
                Dim nLineaPorCobrarCliente As Double
                
                nTotalPorCobrarCliente = 0
                nLineaPorCobrarCliente = 0
                nLineaPorCobrarCliente = Calcular("select isnull(nlineaporcobrar,0) as codigo from vdelivery where codigo='" & sCodigo & "'", Cn)
                If nLineaPorCobrarCliente > 0 Then
                    nTotalPorCobrarCliente = Calcular("select isnull(nconsumoporcobrar,0) as codigo from vdelivery where codigo='" & sCodigo & "'", Cn)
                    Dim nVentaDocActual As Double
                    nVentaDocActual = Calcular("select isnull(nventa,0) as codigo from mdocumento where tdocumento='" & sDocumento & "' ", Cn)
                    If nTotalPorCobrarCliente + nVentaDocActual > nLineaPorCobrarCliente Then
                             MsgBox "Con esta asignación el cliente: " & sDescrip & " supera su linea aprobada para Cuentas por Cobrar " & vbCrLf & " Linea asignada :  " & nLineaPorCobrarCliente & ". Utilizado :  " & nTotalPorCobrarCliente, vbInformation
                             Exit Sub
                     End If
                End If
                
                If MsgBox("Esta seguro de Enviar el documento Nro: " & Format(sDocumento, "@-@@@@@-@@@@@@@@@") & _
                    Chr(13) & "a Cuentas por Cobrar por el Cliente " & sDescrip & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                    Exit Sub
                End If
                
                Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '03', tClientePago = '" & sCodigo & "', tUsuarioAutoriza='" & sVar1 & "',lreplica=1  where tDocumento ='" & sDocumento & "'"
               
                Cn.Execute "update tdelivery set nConsumoPorCobrar= isnull(nconsumoporcobrar,0) + " & nVentaDocActual & " where tcodigodelivery='" & sCodigo & "' "

                'Liberacion
                Dim cpci As Integer
                Dim rst1 As Recordset
                Dim RsT2 As Recordset
                Set rst1 = Lib.OpenRecordset("select tCodigoPedido from DDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
                'Cn.Execute "Update DPEDIDO set tFacturado ='P' where tDocumento ='" & sDocumento & "'"
                
                'Validacion Cuentas Por Cobrar
                Cn.Execute "Update dbo.DPEDIDO set tFacturado ='P', tCortesia='' where tDocumento ='" & sDocumento & "'"
                Cn.Execute "UPDATE dbo.MPEDIDO set tEstadoPedido='02', lReplica=1 where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' )  AND TCODIGOPEDIDO NOT IN (SELECT DISTINCT TCODIGOPEDIDO FROM DPEDIDO WHERE TCODIGOPEDIDO IN(select DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO where tDocumento='" & sDocumento & "') AND ISNULL(TFACTURADO,'') <> 'P')"
                
                If rst1.RecordCount <> 0 Then
                   rst1.MoveFirst
                   For cpci = 1 To rst1.RecordCount
                       Set RsT2 = Lib.OpenRecordset("select tCodigoPedido from DPEDIDO where tCodigoPedido ='" & rst1!tCodigoPedido & "' and (isnull(tFacturado,'0')='0' or tfacturado='F')", Cn)
                       If RsT2.RecordCount = 0 Then
                          Cn.Execute "Update MPEDIDO set tEstadoPedido ='02'  where tCodigoPedido ='" & rst1!tCodigoPedido & "'"
                          Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa in (select tMesa from MPEDIDO where tCodigoPedido='" & rst1!tCodigoPedido & "') "
                       End If
                       rst1.MoveNext
                   Next cpci
                End If
                
                Dim CODMES As String
                
                CODMES = Calcular("SELECT DISTINCT TCODIGOPEDIDO AS CODIGO FROM DPEDIDO WHERE tDocumento ='" & sDocumento & "'", Cn)
                
                '-- mesa247 --- avisa a mesa que fue enviada a cuentas por cobrar
                '    If MESA247 Then
                '        Cn.Execute "Update mpedido set lestadomesapedido = 3 where tcodigopedido ='" & CODMES & "' and idpedidomesa<>'' "
                '    End If
                
                sVar1 = antesvar
                wEnter = True
                lPagocortesiaAI = False
                lGrabaPagoAntes = True
                Unload Me
    
            Case Is = 11  ' Eliminar Recibo
            
                If RsGrillaRecibido.RecordCount <> 0 Then
                        If Not RsGrillaRecibido.BOF Then
                        sMonDoc = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & RsGrillaRecibido!codigo & "'", Cn)
        
                        If sMonDoc = "01" Then
                            nAbonoN = nAbonoN - RsGrillaRecibido!nMonto
                        Else
                            nAbonoE = nAbonoE - RsGrillaRecibido!nMonto
                        End If
        
                            'nIndex = RsGrillaRecibo!nIndex
                            'txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption - RsGrilla!nMonto, "###,###,##0.00")
                            ' txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption - RsGrilla!nPropina, "###,###,##0.00")
                            ' nAbonoN = nAbonoN - RsGrillaRecibido!nMonto
                            'nPropinaN = nPropinaN - RsGrilla!nPropina
                            Cn.Execute "delete from " & xTabla & " where Codigo = '" & RsGrillaRecibido!codigo & "'"
                            CalculaTotal
                            RsGrillaRecibido.Requery
                            'lblTarjetaC.Caption = ""
                            'AsignarTotal
                        If RsGrillaRecibido.RecordCount = 0 Then
                            cmdPago(12).Enabled = True
                        End If
                        End If
                End If
   
End Select
Exit Sub

ErrorEcuador:
    MsgBox "No existe conexión con el sistema SIAB", vbCritical, sMensaje
    Exit Sub

fin:
    Dim errorstring As String
    errorstring = Error
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, errorstring, "", "FALLA PROCESO PAGO DE DOCUMENTO", sUsuario)
    MsgBox "Error: " & errorstring & " | Proceso: Pagos", vbCritical, sMensaje

End Sub
Private Sub cmdOpcionPicture_Click(Index As Integer)
    cmdOpcion_Click Index
End Sub

Private Sub cmdPago_Click(Index As Integer)

On Error GoTo fin
   Select Case Index
          Case Is = 0  ' Efectivo Monto MN
            tCodigoIntegracion = 0
           If lCashDro Then
                If (MsgBox("¿Desea conectarse al CashDrow para el pago?", vbYesNo) = vbNo) Then
                    GoTo PAGOMANUAL
                End If
            
                DoEvents
                Me.lblMensajetexto.Caption = "Se esta esperando el pago del cash drow, favor de esperar....." & vbNewLine & vbNewLine & "TOTAL:" & Me.txtSaldo(0).Caption
                Me.frmMensaje.Visible = True
                DoEvents
                
                Dim IpCashDrow As String
                IpCashDrow = Calcular("select top 1 isnull(tIpCashDrow,'') Codigo from TCAJA where tCaja = '" & sCaja & "'", Cn)
                If Trim(IpCashDrow) <> "" Then
                    Cn.Execute "DELETE FROM  INTEGRACION_CASHDRO where documento='" & sDocumento & "'"
                    IniciarMotorCashDrow ("|Caja:" & sCaja & "-" & sUsuario & "-" & sDocumento & "|" & CDbl(Me.txtSaldo(0).Caption) * 100 & "|" & IpCashDrow & "|0")
                    Me.Timer_CashDro.Enabled = True
                    CmdCancelaCashDro.Visible = True
                Else
                     MsgBox "No hay configuracion de Ip para el cashdro!", vbExclamation, sMensaje
                     Me.frmMensaje.Visible = False
                     GoTo PAGOMANUAL
                End If

                Exit Sub
                
            End If
PAGOMANUAL:
            DoEvents
            Me.frmMensaje.Visible = False
            Frame1.Enabled = True
            Me.fraEfectivo.Enabled = True
          
            If nSaldo > 0 Or nAbonoN > 0 Then
               nAbonoN = nAbonoN - nEN
               sTipo = ""
               frmNumPad.Show vbModal
               nEN = IIf(wEnter = True, sDescrip, nEN)
               txtME(0).Caption = Format(nEN, "###,###,###,##0.00")
               nAbonoN = nAbonoN + nEN
               CalculaTotal
            Else
                MsgBox "No hay saldo a cancelar"
            End If
          Case Is = 2  ' Efectivo Monto ME
            If nSaldo > 0 Or nAbonoE > 0 Then
                nAbonoE = nAbonoE - nEE
                sTipo = ""
                frmNumPad.Show vbModal
                nEE = IIf(wEnter = True, sDescrip, nEE)
                txtME(1).Caption = Format(nEE, "###,###,###,##0.00")
                nAbonoE = nAbonoE + nEE
                CalculaTotal
            Else
                MsgBox "No hay saldo a cancelar"
            End If
               
                              
          Case Is = 4  ' Banco
               frmKeyBoard.txtResultado.Text = txtBCH.Caption
               frmKeyBoard.Show vbModal
               If wEnter Then
                  txtBCH.Caption = Mid(sDescrip, 1, 30)
               End If
          
          Case Is = 5  ' Numero
               sTipo = "Numero"
               frmNumPad.Show vbModal
               txtNCH.Caption = IIf(wEnter = True, Mid(sDescrip, 1, 15), txtNCH.Caption)
          
          Case Is = 6  ' Monto Cheque
               nAbonoN = nAbonoN - nCH
               sTipo = "Prepintado"
               sCodigo = nSaldo
               frmNumPad.Show vbModal
               nCH = IIf(wEnter = True, sDescrip, nCH)
               txtMCH(0).Caption = Format(nCH, "###,###,###,##0.00")
               nAbonoN = nAbonoN + nCH
               CalculaTotal
                     
          Case Is = 3  ' Monto Cheque Dolares
               nAbonoE = nAbonoE - nCE
               sTipo = ""
               frmNumPad.Show vbModal
               nCE = IIf(wEnter = True, sDescrip, nCE)
               txtMCH(1).Caption = Format(nCE, "###,###,###,##0.00")
               nAbonoE = nAbonoE + nCE
               CalculaTotal
               
          Case Is = 7  ' Numero Tarjeta
               sTipo = "Numero"
               frmNumPad.Show vbModal
               txtNT.Caption = IIf(wEnter = True, Mid(sDescrip, 1, 16), txtNT.Caption)
               
               If lClub = True And pais = "002" Then
                     'CHK TARJETA
                    Dim i As Integer
                    Dim tPrefijo As String
                    Dim xPrefijo As Boolean
                    Dim RSTarjetaCredito As Recordset
                    Dim RsTarjetaPrefijos As Recordset
               
                    xPrefijo = False
                
                    If lBotonTarjeta = False Then 'lBotonTarjeta = True
                                            
                           For i = 1 To 4
                                   tPrefijo = Mid(sDescrip, 1, 5 - i)
                                   Isql = "Select * From TTARJETACREDITO Where tPrefijo= '" & tPrefijo & "' "
                                   Set RSTarjetaCredito = Lib.OpenRecordset(Isql, Cn)
                    
                                   If RSTarjetaCredito.RecordCount > 0 Then
                                       lblTarjetaC.Caption = RSTarjetaCredito!tResumido
                                       sTipoTarjeta = RSTarjetaCredito!tCodigoTarjeta
                                       sTitulo = RSTarjetaCredito!tResumido
                                       nIndex = RSTarjetaCredito!nBoton
                                       xPrefijo = True
                                       Exit Sub
                                   End If
                           Next i
                    
                            If xPrefijo = False Then
                              MsgBox "Error: N° de Tarjeta no válido", vbExclamation, sMensaje
                              txtNT.Caption = ""
                              lblTarjetaC.Caption = ""
                              ActivaTarjeta False
                           End If
                           
                           
                    End If
               End If
               
          Case Is = 8  ' Fecha de Vencimiento
               sTipo = "Fecha"
               frmNumPad.Show vbModal
               
               If wEnter = True Then
                  If Len(sDescrip) = 4 And val(Mid(sDescrip, 1, 2)) <= 12 And val(Mid(sDescrip, 1, 2)) >= 1 Then
                     txtVT.Caption = sDescrip
                  Else
                     MsgBox "Formato no válido (MMAA)", vbInformation, sMensaje
                     txtVT.Caption = ""
                  End If
               Else
                  txtVT.Caption = ""
               End If
                              
          Case Is = 9  ' Monto Tarjeta
               sTipo = "Prepintado"
               sCodigo = nSaldo
               frmNumPad.Show vbModal
                    
               If Not wEnter Then
                  Exit Sub
               End If
               
               'frmNumPad.Show vbModal
               nET = IIf(wEnter = True, sDescrip, nET)
               txtMT.Caption = Format(nET, "###,###,###,##0.00")
          
          Case Is = 10  ' Propina Tarjeta
               sTipo = ""
               frmNumPad.Show vbModal
               nPT = IIf(wEnter = True, sDescrip, nPT)
               txtPT.Caption = Format(nPT, "###,###,###,##0.00")
               
               If sRefer <> "" Then
                  nPropinaN = nPropinaN + nPT
                  CalculaTotal
                  txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT - RsGrilla!nPropina, "###,###,##0.00")
                  Cn.Execute "update " & sTarjeta & " set nPropina=" & nPT & " where tReferencia='" & sRefer & "'"
                  RsGrilla.Requery
               End If
          
          Case Is = 11  ' Nota de Crédito
               sTemp = ""
               Isql = "select *, tNotaCredito as Codigo, nVenta as Descripcion from vNotaCredito where tEstadodocumento ='05'"
               Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Nota Credito", 2, "tNotaCredito", 1900, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1200, 0, 0, "dd/mm/yyyy", _
                                                               "Cliente", 2, "Cliente", 3900, 0, 0, "", _
                                                               "Monto", 2, "nVenta", 1200, 1, 0, "##,##0.00")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               If wEnter Then
                  txtDocumento.Caption = "Nota Credito"
                  sTipoDocumento = sCodigo
                  txtReferencia.Caption = sCodigo
                  sOtroTipoCancelacion = "002"
                  nAbonoN = nAbonoN - nDocumento
                  nDocumento = val(sDescrip)
                  sOtroTipo = sTipo
               Else
                  Exit Sub
               End If
               
               txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
               nAbonoN = nAbonoN + nDocumento
               CalculaTotal
               
          Case Is = 12  ' Recibos enlazados

                Isql = "insert into " & xTabla & " select  tRecibo as Codigo,tRecibo,fFecha,tDescripcion,Moneda,nMonto, nMonto as Descripcion from vIngreso where tEstadodocumento ='01' and lAnticipo=1 and Pedido in (select distinct tCodigoPedido from ddocumento where tdocumento ='" & sDocumento & "')"
                 Cn.Execute Isql
                 RsGrillaRecibido.Requery
                 
                 If RsGrillaRecibido.RecordCount > 0 Then
                    cmdPago(12).Enabled = False
                    
                     RsGrillaRecibido.MoveFirst
                     Do While Not RsGrillaRecibido.EOF
                        sMonDoc = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & RsGrillaRecibido!codigo & "'", Cn)
    
                        If sMonDoc = "01" Then
                           nAbonoN = nAbonoN + RsGrillaRecibido!nMonto
                        Else
                           nAbonoE = nAbonoE + RsGrillaRecibido!nMonto
                        End If
                        CalculaTotal
                        RsGrillaRecibido.MoveNext
                     Loop
                 Else
                    MsgBox "No hay Recibos Enlazados!!!", vbInformation
                 End If
                 RsGrillaRecibido.Requery
                 
               'frmBusquedaRapida.tipoBusqueda = "PAGO"
            '   Call ConfGrilla(6, frmBusquedaRapida.grdGrilla, "Recibo", 2, "tRecibo", 1200, 2, 0, "", _
            '                                                   "Fecha", 2, "fFecha", 1150, 0, 0, "dd/mm/yyyy", _
            '                                                   "Motivo", 2, "tDescripcion", 3000, 0, 0, "", _
            '                                                   "Mon", 2, "Moneda", 600, 2, 0, "", _
            '                                                   "Monto", 2, "nMonto", 1200, 1, 0, "##,##0.00", _
            '                                                   "Sel", 2, "Sel", 850, 2, 4, "")


            '            Set RsGrillaRecibo = Lib.OpenRecordset(Isql, Cn)
            '            If RsGrillaRecibo.RecordCount = 0 Then
            '                Set grdGrillaRecibo.DataSource = Nothing
            '                MsgBox ("No hay Recibos")
            '                SSTab1.Tab = 0
            '            Else
            '                Set grdGrillaRecibo.DataSource = RsGrillaRecibo
            '                SSTab1.Tab = 1
            '            End If

               
               
            '   frmBusquedaRapida.nPredeterm = 0
            '   frmBusquedaRapida.Show vbModal
                    
            '    If wEnter Then
            '      txtDocumento.Caption = "Recibo/Anticipo"
            '      stipodocumento = sCodigo
            '      txtReferencia.Caption = sCodigo
            '      sOtroTipoCancelacion = "001"
            '      sOtroTipo = sTipo

            '      If sMonDoc = "01" Then
            '         nAbonoN = nAbonoN - nDocumento
            '      Else
            '         nAbonoE = nAbonoE - nDocumento
            '      End If

            '      sMonDoc = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & sCodigo & "'", Cn)

            '      If sMonDoc = "01" Then
            '         nAbonoN = nAbonoN + Val(sDescrip)
            '      Else
            '         nAbonoE = nAbonoE + Val(sDescrip)
            '      End If
            '      nDocumento = Val(sDescrip)
            '   Else
            '      Exit Sub
            '    End If
            '   txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
            '   CalculaTotal

          Case Is = 15  ' Recibos
                        
                Dim RsGrillaRecibo2 As Recordset
                sTemp = ""

                Isql = "select *, tRecibo as Codigo, nMonto as Descripcion, convert(bit,0) as Sel from vIngreso where tEstadodocumento ='01' and lAnticipo=1 and pedido in ('') and  trecibo not in (select codigo from " & xTabla & ")"

               frmBusquedaRapida.tipoBusqueda = "PAGO"
               Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Recibo", 2, "tRecibo", 1200, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1150, 0, 0, "dd/mm/yyyy", _
                                                               "Motivo", 2, "tDescripcion", 3000, 0, 0, "", _
                                                               "Mon", 2, "Moneda", 700, 2, 0, "", _
                                                               "Monto", 2, "nMonto", 1200, 1, 0, "##,##0.00")


                    '    Set RsGrillaRecibo2 = Lib.OpenRecordset(Isql, Cn)
                    '    If RsGrillaRecibo.RecordCount = 0 Then
                    '        Set grdGrillaRecibo.DataSource = Nothing
                    '        MsgBox ("No hay Recibos")
                    '        SSTab1.Tab = 0
                    '    Else
                    '        Set grdGrillaRecibo.DataSource = RsGrillaRecibo
                    '        SSTab1.Tab = 1
                    '    End If


               frmBusquedaRapida.nPredeterm = 0
               frmBusquedaRapida.Show vbModal


               If wEnter Then
                '  Isql = "insert into " & xTabla & " select  tRecibo as Codigo,tRecibo,fFecha,tDescripcion,Moneda,nMonto, nMonto as Descripcion from vIngreso where tEstadodocumento ='01' and lAnticipo=1 and trecibo in ('" & sCodigo & "')"
                    Isql = "insert into " & xTabla & " select  tRecibo as Codigo,tRecibo,fFecha,tDescripcion,Moneda,nMonto, nMonto as Descripcion from vIngreso where tEstadodocumento ='01' and lAnticipo=1  and  trecibo in ('" & sCodigo & "')" ' and trecibo not in (select codigo from " & xTabla & " ) "
                    Cn.Execute Isql
                    RsGrillaRecibido.Requery

                  sMonDoc = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & sCodigo & "'", Cn)

                  If sMonDoc = "01" Then
                     nAbonoN = nAbonoN + val(sDescrip)
                  Else
                     nAbonoE = nAbonoE + val(sDescrip)
                  End If
                  'nDocumento = Val(sDescrip)
                   cmdPago(12).Enabled = False
               Else
                  Exit Sub
               End If
               'txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
               CalculaTotal

          Case Is = 13  ' otros formas de pag (Varios)
          
                'sUsuarioAutoriza = sUsuario
                If lPassOtrosPagos Then
                    tUsuActua = sUsuario
                    If Supervisor("28") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                    End If
                      sUsuario = tUsuActua
                      sUsuarioAutoriza = sVar1
                End If
                
               sTemp = ""
               Isql = "select Codigo, Descripcion from vTipoCancelacion Where lActivo='1' and Codigo<>'001' and Codigo<>'002'"
               Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                               "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
                              
               If wEnter = True Then
                  If sCodigo = "000" Then
                     frmKeyBoard.Caption = "Otro Tipo de Cancelación"
                     frmKeyBoard.Show vbModal
                     sOtroTipoCancelacion = "000"
                  Else
                     sOtroTipoCancelacion = sCodigo
                  End If
                  txtDocumento.Caption = sDescrip
               Else
                  Exit Sub
               End If

               nAbonoN = nAbonoN - nDocumento
               sTipo = ""
               frmNumPad.Show vbModal
               nDocumento = IIf(wEnter = True, sDescrip, nDocumento)
               
                Dim lActivaCredito As String
                lActivaCredito = Calcular("select isnull(nvalor,'0') Codigo from vTipoCancelacion Where lActivo='1' and Codigo='" & sOtroTipoCancelacion & "'", Cn)

               '-----------------------------------------------------------------
                'GCAA 06122022 -  CORPORACION FERNADEZ
                If lActivaCreditoCoorporativo And lActivaRestriClienteDelivery And lActivaCredito = "1" Then
 
                    Dim LineaTotal As Double, SaldoCliente As Double
                    Dim codigoDelivery As String
                    codigoDelivery = Calcular("select isnull(tClienteDelivery,'') as codigo from MPEDIDO where tCodigoPedido='" & sPedido & "'", Cn)

                    LineaTotal = 0
                    SaldoCliente = 0
                    LineaTotal = Calcular("select isnull(nLinea,0) as codigo from vCompania where lActivo = 1 AND codigo = '" & codigoDelivery & "'", Cn)
                    If LineaTotal > 0 Then
                        SaldoCliente = Calcular("select isnull(nSaldo,0) as codigo from vCompania where lActivo = 1 AND codigo = '" & codigoDelivery & "'", Cn)
                        If CDbl(nDocumento) > SaldoCliente Then
                            Dim ClienteCorporativo As String
                            ClienteCorporativo = Calcular("select isnull(Descripcion,'') as codigo from vCompania where lActivo = 1 AND codigo = '" & codigoDelivery & "'", Cn)

                            MsgBox "El cliente " & ClienteCorporativo & vbCrLf & " Linea asignada :  " & LineaTotal & vbCrLf & " Saldo: " & SaldoCliente & vbCrLf & "No cuenta con saldo y/o el valor ingresado supera al saldo!", vbInformation
                            nDocumento = 0
                            wEnter = False
                            Exit Sub
                         End If
                    Else
                       MsgBox "Cliente no cuenta con linea de credito corporativo!!!", vbExclamation, sMensaje
                       nDocumento = 0
                       wEnter = False
                       Exit Sub
                    End If

                End If
               
               '----------------
               txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
               nAbonoN = nAbonoN + nDocumento
               CalculaTotal
               
               sTipo = "Comanda"
               frmNumPad.Caption = "Referencia"
               frmNumPad.Show vbModal
               sTipoDocumento = IIf(wEnter = True, sDescrip, "")
               txtReferencia.Caption = sTipoDocumento
               
          Case Is = 14  ' Puntos
               Dim nDisponible As Double
               sTemp = ""
               Isql = "select Codigo, Cliente as Descripcion, nDisponible from vDelivery Where lActivo=1 and lPuntos=1 "
               Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1500, 2, 0, "", _
                                                               "Cliente", 2, "Descripcion", 5300, 0, 0, "", _
                                                               "Puntos", 2, "nDisponible", 1400, 1, 0, "###,##0.00")
               
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               nAbonoN = nAbonoN - nTotalPuntos
               nTotalPuntos = 0
               txtPunto.Caption = "0.00"
               CalculaTotal
               
               If wEnter = True Then
                  sClientePuntos = sCodigo
                  txtCliente.Caption = sDescrip
                  nDisponible = Calcular("select nDisponible as Codigo from vDelivery where Codigo='" & sCodigo & "'", Cn)
               Else
                  txtCliente.Caption = ""
                  Exit Sub
               End If
                              
               sTipo = ""
               frmNumPad.Show vbModal
               
               If Not wEnter Then
                  txtCliente.Caption = ""
                  Exit Sub
               End If
               
               If val(sDescrip) > nDisponible Then
                  MsgBox "Monto mayor a lo disponible", vbExclamation, sMensaje
                  txtCliente.Caption = ""
                  Exit Sub
               End If
               
               nTotalPuntos = IIf(wEnter = True, sDescrip, nTotalPuntos)
               txtPunto.Caption = Format(nTotalPuntos, "###,###,###,##0.00")
               nAbonoN = nAbonoN + nTotalPuntos
               CalculaTotal
          Case Is = 16  ' Efectivo Monto ME 2
                If Round(nSaldo, 2) > 0 Or nAbonoE2 > 0 Then
                nAbonoE2 = nAbonoE2 - nEE2
                sTipo = ""
                frmNumPad.Show vbModal
                nEE2 = IIf(wEnter = True, sDescrip, nEE2)
                txtME(2).Caption = Format(nEE2, "###,###,###,##0.00")
                nAbonoE2 = nAbonoE2 + nEE2
                CalculaTotal
                Else
                    MsgBox "No hay saldo a cancelar"
                End If
          Case Is = 17  ' Efectivo Monto ME 2
                If Round(nSaldo, 2) > 0 Or nAbonoE3 > 0 Then
                nAbonoE3 = nAbonoE3 - nEE3
                sTipo = ""
                frmNumPad.Show vbModal
                nEE3 = IIf(wEnter = True, sDescrip, nEE3)
                txtME(3).Caption = Format(nEE3, "###,###,###,##0.00")
                nAbonoE3 = nAbonoE3 + nEE3
                CalculaTotal
                Else
                    MsgBox "No hay saldo a cancelar"
                End If
    End Select
    
   On Error Resume Next
   
   If lActivaReciboSaldoQr Then
        txtQRRecibo.SetFocus
   End If
    
   Exit Sub
fin:

    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "FALLA PAGO DE DOCUMENTO ", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Private Sub cmdPagoEfectivo_Click()
On Error GoTo fin
    Screen.MousePointer = vbHourglass
    lblCIPpagoEfectivo.Caption = ""
    Cn.Execute "DELETE FROM DOCUMENTO_INTEGRACIONES WHERE DOCUMENTO='" & sDocumento & "'"
    
       If (EjecutaMotorIntegraciones(sDocumento, sUsuario, 7, 1) = True) Then
        Call ActivaDesactivaPagoEfectivo(False)
        Dim cIp As String
        cIp = Calcular("select Isnull(PagoEfectivoCip,'') as codigo from DOCUMENTO_INTEGRACIONES where Documento='" & sDocumento & "'", Cn)
        lblCIPpagoEfectivo.Caption = "CIP: " & cIp
        
        Dim url As String
        url = Calcular("select isnull(PagoEfectivoUrlCip,'') as Codigo from documento_integraciones where documento='" & sDocumento & "'", Cn)
        If url = "" Or url = "0" Or cIp = "0" Or cIp = "" Then
            MsgBox "No se pudo generar el Codigo CIP de pago efectivo, Favor intente nuevamente!!!", vbInformation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Dim X
        X = ShellExecute(Me.hwnd, "Open", url, &O0, &O0, SW_NORMAL)

        If lPagoAntesImpresion Then
            lGrabaPagoAntes = True
        End If
        cmdOpcion_Click (99)
        
'        Me.Timer_MercadoPago.Enabled = True
'        Me.Timer_MercadoPago.Interval = 1000
'        Me.Timer_MercadoPago.Enabled = True
        
       Else
            MsgBox "No se pudo generar el codigo CIP PagoEfectivo en la operacion!, Favor de contactar al administrador de su servicios", vbInformation
        Call ActivaDesactivaPagoEfectivo(True)
       End If
    
    Screen.MousePointer = vbDefault
Exit Sub
fin:
    Call ActivaDesactivaPagoEfectivo(True)
    MsgBox "Error: " & Error
End Sub
Private Sub cmdCancelarPagoEfectivo_Click()
    If (MsgBox("¿Esta seguro de cancelar el proceso?", vbYesNo) = vbYes) Then
        Call ActivaDesactivaPagoEfectivo(True)
        Me.Timer_MercadoPago.Enabled = False
        Me.Timer_MercadoPago.Interval = 0
        Me.Timer_MercadoPago.Enabled = False
    End If
End Sub

Private Sub lblCIPpagoEfectivo_Click()

    Dim url As String
    url = Calcular("select isnull(PagoEfectivoUrlCip,'') as Codigo from documento_integraciones where documento='" & sDocumento & "'", Cn)
    If url = "" Then
        MsgBox "EL documento no cuenta con codigo CIP de pago efectivo!!!", vbInformation
        Exit Sub
    End If
    Dim X
    X = ShellExecute(Me.hwnd, "Open", url, &O0, &O0, SW_NORMAL)

End Sub
Private Sub ActivaDesactivaPagoEfectivo(Valor As Boolean)
On Error GoTo fin
    fraEfectivo.Enabled = Valor
    'fraCheque.Enabled = valor
    fraPunto.Enabled = Valor
    Frame1.Enabled = Valor
    SSTab.Enabled = Valor
    fraOpciones.Enabled = Valor
    'cmdPagoEfectivo.Enabled = Valor
    lblclick.Visible = Not Valor
    lblInformacion.Visible = Not Valor
    lblCIPpagoEfectivo.Visible = Not Valor
    If lDesPagoCheque And sFormulario = "CajaRapida" Then
        Me.fraCheque.Enabled = False
    Else
        fraCheque.Enabled = Valor
    End If
Exit Sub
fin:
End Sub

Private Sub cmdTarjeta_Click(Index As Integer)
On Error GoTo fin
    
     lTarjetaGrilla = False
     lPinpadTarjeta = False
     
     lBotonTarjeta = True

     If Index = "8" And cmdTarjeta(8).Caption = "OTRAS" Then
                 
        sTemp = ""
        'Isql = "select tcodigotarjeta codigo, tdetallado,tresumido as descripcion,isnull(nfactorretencion,0) nfactorretencion, isnull(lpinpad,0) lpinpad, case when isnull(nboton,0)=0 then 99 else nboton end nboton, lactivo From TTARJETACREDITO Where lactivo = 1 order by tdetallado"
        Dim xCondicional As String
        xCondicional = ""
        Isql = "select TC.tcodigotarjeta codigo, TC.tdetallado,tresumido as descripcion,isnull(TC.nfactorretencion,0) nfactorretencion, isnull(TC.lpinpad,0) lpinpad, case when isnull(nboton,0)=0 then 99 else nboton end nboton, lactivo From TTARJETACREDITO TC left join TVISIBILIDADTARJETACREDITOXCANAL VTxC on VTxC.tCodigoTarjeta = TC.tCodigoTarjeta" _
         & " Where TC.lactivo = 1"
        Select Case sTipoPedido
            Case Is = "01"
                xCondicional = " and VTxC.lVisibilidad = 1 and VTxC.lCanal1 = 1 order by tdetallado"
            Case Is = "02"
                xCondicional = " and VTxC.lVisibilidad = 1 and VTxC.lCanal2 = 1 order by tdetallado"
            Case Is = "03"
                xCondicional = " and VTxC.lVisibilidad = 1 and VTxC.lCanal3 = 1 order by tdetallado"
            Case Is = "04"
                xCondicional = " and VTxC.lVisibilidad = 1 and VTxC.lCanal4 = 1 order by tdetallado"
            Case Is = "05"
                xCondicional = " and VTxC.lVisibilidad = 1 and VTxC.lCanal5 = 1 order by tdetallado"
        End Select
        Isql = Isql & xCondicional
        
        Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                        "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
        frmBusquedaRapida.nPredeterm = 1
        frmBusquedaRapida.Show vbModal
        If wEnter Then
        
            If Calcular("select COUNT(*) as codigo from TTARJETACREDITO where tCodigoTarjeta = '" & sCodigo & "' and tCodigoIntegracion = '19'", Cn) > 0 Then
               If Calcular("select COUNT(*) as codigo from TCAJA where tCaja = '" & sCaja & "' and lActivaIntegracionNiubiz = 1", Cn) = 0 Then
                   MsgBox "La caja actual no tiene habilitada la integración con la terminal de Niubiz.", vbExclamation, sMensaje
                   Exit Sub
               End If
            End If
     
            nIndex = 8
            sTipoTarjeta = sCodigo
            sTitulo = sDescrip
            lTarjetaGrilla = True
            lPinpadTarjeta = Calcular("select isnull(lpinpad,0) codigo from ttarjetacredito where tcodigotarjeta='" & sTipoTarjeta & "'", Cn)
            
            If lPinpadTarjeta And lVisaNet Then
                sTipo = "Prepintado"
                sCodigo = nSaldo
                frmNumPad.Show vbModal
                    
                If Not wEnter Then
                    Exit Sub
                End If
                
                If val(sDescrip) <= 0 Then
                    MsgBox "Error: El monto no puede ser cero", vbExclamation, sMensaje
                    Exit Sub
                End If
                
                If nCargo < val(Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)) + val(sDescrip) Then
                    MsgBox "El abono(s) de las tarjeta(s) no puede ser mayor al total del documento", vbExclamation, sMensaje
                    Exit Sub
                End If
                Mensaje "PinPad Listo. Esperando...", "PinPad", 10
                
                sMonto = Format(sDescrip, "0000000000.00")
                sMonto = Mid(sMonto, 1, 10) & Mid(sMonto, 12, 2)
                
                sOperacion = OP_FINANCIERA & "A" & sMonto & Chr$(FS) & _
                                            "B" & "000000000000" & Chr$(FS) & _
                                            "C" & "0" & Chr$(FS) & _
                                            "D" & sEmpresa & Chr$(FS) & _
                                            "E" & sCaja
                nRet = fiStartOperation(sOperacion, 2, sMensaje)
                
                If nRet = RET_OK Or nRet = RET_RUNNING Then
                    If Not Imprimir(sPreCuenta) Then
                    Exit Sub
                    End If
                    Printer.FontName = sFont
                    Printer.FontBold = False
                                
                    sClave = ""
                    lEmisor = True
                    lLoop = True
                    nContador = 0
                    
                    Do
                    sRetorno = ""
                    nRet = fiGetStatus(sRetorno, 512)
                    lEmisor = ImprimeCabecera(sRetorno, lEmisor)
                    sClave = MensajePinPad(sRetorno)
                    If Mid(sClave, 1, 3) = "A00" Or Mid(sClave, 1, 3) = "A11" Then
                        txtNT.Caption = BuscaRetornoPinPad(sClave, "F")
                        txtVT.Caption = BuscaRetornoPinPad(sClave, "G")
                        sRefer = BuscaRetornoPinPad(sClave, "E")
                        
                        'Agrego al temporal
                        nET = IIf(wEnter = True, sDescrip, nET)
                        nAbonoN = nAbonoN + nET
                        nPropinaN = nPropinaN + nPT
                        CalculaTotal
                            
                        Isql = "insert into " & sTarjeta & _
                                "( tIndex, tTipo, tDescripcion, tNumero, tFecha, tReferencia, nMonto, nPropina, nIndex, lPinPad ) " & _
                                "Values(    '" & nCorrela & "', " _
                                        & "'" & sTipoTarjeta & "'," _
                                        & "'" & sTitulo & "'," _
                                        & "'" & txtNT.Caption & "'," _
                                        & "'" & txtVT.Caption & "'," _
                                        & "'" & sRefer & "'," _
                                        & nET & ", " _
                                        & nPT & ", " _
                                        & nIndex & ", 1)"
                        Cn.Execute Isql
                        Isql = "insert into DPAGOTARJETA " & _
                                "( tDocumento, tReferencia, tTarjeta, tNumero, tFechaVencimiento, nMonto, nPropina, tEstadoDocumento ) " & _
                                "Values(    '" & sDocumento & "', " _
                                        & "'" & sRefer & "'," _
                                        & "'" & sTipoTarjeta & "'," _
                                        & "'" & txtNT.Caption & "'," _
                                        & "'" & txtVT.Caption & "'," _
                                        & nET & ", " _
                                        & nPT & ", '01')"
                        Cn.Execute Isql
                        RsGrilla.Requery
                        txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
                        txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
                        AsignarTotal
                    Else
                        xError = BuscaRetornoPinPad(sClave, "B")
                        If Len(xError) > 0 Then
                            Mensaje xError, "VisaNet", 500
                        End If
                    End If
                                
                    Mensaje "PinPad Listo. Esperando...", "PinPad", 500
                    nContador = nContador + 1
                    If nContador >= nTimeOut Then
                        If MsgBox("Tiempo de espera agotado, deseas mas tiempo?", vbExclamation + vbOKCancel, "VisaNet") = vbOK Then
                            lLoop = True
                            nContador = nTimeOut / 2
                        Else
                            lLoop = False
                        End If
                    End If
                    
                    If nRet <> "0" Then
                        nContador = 0
                    End If
                    Loop While (Mid$(sRetorno, 5, 2) <> "C1") And lLoop
                Else
                    MsgBox "Error de conectividad", vbCritical, sMensaje
                    Exit Sub
                End If
        
            Else
                ActivaTarjeta True
                txtNT.Caption = ""
                txtVT.Caption = ""
                txtMT.Caption = "0.00"
                txtPT.Caption = "0.00"
            End If
        End If
    Else
        
        sCodigoIntegracion = mTarjeta(Index, 4)
                    
        If sCodigoIntegracion = "18" And lActivarValesConsumo Then 'Solo si está activado en ADM y enlazado a SERVILEXll
            sTemp = ""
            ' Isql = "select tcodigotarjeta codigo, tdetallado,tresumido as descripcion,isnull(nfactorretencion,0) nfactorretencion, isnull(lpinpad,0) lpinpad, case when isnull(nboton,0)=0 then 99 else nboton end nboton, lactivo From TTARJETACREDITO Where lactivo = 1 order by tdetallado"
            ' Call ConfGrilla(2, frmBusquedaValeConsumo.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
            '                                                 "Descripcion", 2, "Descripcion", 7000, 0, 0, "")

            If lGuardarVC = False Then
                Set RsGrillaVConsumo = Nothing
                Cn.Execute "delete from " & sTarjetaVC & ""
            End If

            frmBusquedaValeConsumo.nPredeterm = 1
            frmBusquedaValeConsumo.Show vbModal

            If wEnter Then
                nIndex = Index
                ' sTipoTarjeta = sCodigo
                ' sTitulo = sDescrip
                sTipoTarjeta = mTarjeta(Index, 1)
                sTitulo = mTarjeta(Index, 2)
                lTarjetaGrilla = True
                
                ActivaTarjeta True
                cmdTarjeta(nIndex).Enabled = True
                cmdPago(9).Enabled = False 'Monto

                txtNT.Caption = ""
                txtVT.Caption = ""
                nET = sDescrip
                txtMT.Caption = Format(nET, "###,###,###,##0.00")
                txtPT.Caption = "0.00"

            End If

            Exit Sub
        Else

            sTipoTarjeta = mTarjeta(Index, 1)
            sTitulo = mTarjeta(Index, 2)
            sRefer = ""
            nIndex = Index
            nET = 0
            nPT = 0
            lblTarjetaC.Caption = sTitulo

        End If
        
        If mTarjeta(Index, 3) And lVisaNet Then
            sTipo = "Prepintado"
            sCodigo = nSaldo
            frmNumPad.Show vbModal
                
            If Not wEnter Then
                Exit Sub
            End If
            
            If val(sDescrip) <= 0 Then
                MsgBox "Error: El monto no puede ser cero", vbExclamation, sMensaje
                Exit Sub
            End If
            
            If nCargo < val(Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)) + val(sDescrip) Then
                MsgBox "El abono(s) de las tarjeta(s) no puede ser mayor al total del documento", vbExclamation, sMensaje
                Exit Sub
            End If
            Mensaje "PinPad Listo. Esperando...", "PinPad", 10
            
            sMonto = Format(sDescrip, "0000000000.00")
            sMonto = Mid(sMonto, 1, 10) & Mid(sMonto, 12, 2)
            
            sOperacion = OP_FINANCIERA & "A" & sMonto & Chr$(FS) & _
                                        "B" & "000000000000" & Chr$(FS) & _
                                        "C" & "0" & Chr$(FS) & _
                                        "D" & sEmpresa & Chr$(FS) & _
                                        "E" & sCaja
            nRet = fiStartOperation(sOperacion, 2, sMensaje)
            
            If nRet = RET_OK Or nRet = RET_RUNNING Then
                If Not Imprimir(sPreCuenta) Then
                Exit Sub
                End If
                Printer.FontName = sFont
                Printer.FontBold = False
                            
                sClave = ""
                lEmisor = True
                lLoop = True
                nContador = 0
                
                Do
                sRetorno = ""
                nRet = fiGetStatus(sRetorno, 512)
                lEmisor = ImprimeCabecera(sRetorno, lEmisor)
                sClave = MensajePinPad(sRetorno)
                If Mid(sClave, 1, 3) = "A00" Or Mid(sClave, 1, 3) = "A11" Then
                    txtNT.Caption = BuscaRetornoPinPad(sClave, "F")
                    txtVT.Caption = BuscaRetornoPinPad(sClave, "G")
                    sRefer = BuscaRetornoPinPad(sClave, "E")
                    
                    'Agrego al temporal
                    nET = IIf(wEnter = True, sDescrip, nET)
                    nAbonoN = nAbonoN + nET
                    nPropinaN = nPropinaN + nPT
                    CalculaTotal
                        
                    Isql = "insert into " & sTarjeta & _
                            "( tIndex, tTipo, tDescripcion, tNumero, tFecha, tReferencia, nMonto, nPropina, nIndex, lPinPad, tNumeroATarjeta ) " & _
                            "Values(    '" & nCorrela & "', " _
                                    & "'" & sTipoTarjeta & "'," _
                                    & "'" & sTitulo & "'," _
                                    & "'" & txtNT.Caption & "'," _
                                    & "'" & txtVT.Caption & "'," _
                                    & "'" & sRefer & "'," _
                                    & nET & ", " _
                                    & nPT & ", " _
                                    & nIndex & ", 1, '" & txtNAT.Caption & "')"
                    Cn.Execute Isql
                    Isql = "insert into DPAGOTARJETA " & _
                            "( tDocumento, tReferencia, tTarjeta, tNumero, tFechaVencimiento, nMonto, nPropina, tEstadoDocumento ) " & _
                            "Values(    '" & sDocumento & "', " _
                                    & "'" & sRefer & "'," _
                                    & "'" & sTipoTarjeta & "'," _
                                    & "'" & txtNT.Caption & "'," _
                                    & "'" & txtVT.Caption & "'," _
                                    & nET & ", " _
                                    & nPT & ", '01')"
                    Cn.Execute Isql
                    RsGrilla.Requery
                    txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
                    txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
                    AsignarTotal
                Else
                    xError = BuscaRetornoPinPad(sClave, "B")
                    If Len(xError) > 0 Then
                        Mensaje xError, "VisaNet", 500
                    End If
                End If
                            
                Mensaje "PinPad Listo. Esperando...", "PinPad", 500
                nContador = nContador + 1
                If nContador >= nTimeOut Then
                    If MsgBox("Tiempo de espera agotado, deseas mas tiempo?", vbExclamation + vbOKCancel, "VisaNet") = vbOK Then
                        lLoop = True
                        nContador = nTimeOut / 2
                    Else
                        lLoop = False
                    End If
                End If
                
                If nRet <> "0" Then
                    nContador = 0
                End If
                Loop While (Mid$(sRetorno, 5, 2) <> "C1") And lLoop
            Else
                MsgBox "Error de conectividad", vbCritical, sMensaje
                Exit Sub
            End If

        Else
            ActivaTarjeta True
            txtNT.Caption = ""
            txtVT.Caption = ""
            txtMT.Caption = "0.00"
            txtPT.Caption = "0.00"
            txtNAT.Caption = ""
        End If
    
    End If
    On Error Resume Next
    
    If lActivaReciboSaldoQr Then
        Me.txtQRRecibo.SetFocus
    End If
    
    Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO Tarjetas", "", "", sDocumento, Error, "", "FALLA PAGO DE DOCUMENTO Tarjeta ", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos Tarjetas", vbCritical, sMensaje
End Sub
Private Sub ActivaTarjetaCreditoxCanal(cantidadBoton As Integer)
    Dim i As Integer
    'Dim numCantidad As Integer
    Dim rsVisibilidadTarjetaxCanal As Recordset
    'Isql = "select count(*) as Codigo from TVISIBILIDADTARJETACREDITOXCANAL where tCodigoTarjeta = '02'"
    'numCantidad = Calcular(Isql, Cn)
    Isql = "select * from TVISIBILIDADTARJETACREDITOXCANAL"
    Set rsVisibilidadTarjetaxCanal = Lib.OpenRecordset(Isql, Cn)
    
    If rsVisibilidadTarjetaxCanal.RecordCount > 0 Then
        Select Case sTipoPedido
            Case Is = "01"
                For i = 1 To cantidadBoton
                    rsVisibilidadTarjetaxCanal.Filter = ""
                    rsVisibilidadTarjetaxCanal.MoveFirst
                    rsVisibilidadTarjetaxCanal.Filter = "tCodigoTarjeta = '" & mTarjeta(i, 1) & "'"
                    'verificar si rsVisibilidadTarjetaxCanal tiene contenido
                    If Not rsVisibilidadTarjetaxCanal.EOF Then
                        cmdTarjeta(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal1), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTMonto(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal1), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTPropina(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal1), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                    Else
                        cmdTarjeta(i).Enabled = False
                        txtTMonto(i).Enabled = False
                        txtTPropina(i).Enabled = False
                    End If
                Next i
            Case Is = "02"
                For i = 1 To cantidadBoton
                    rsVisibilidadTarjetaxCanal.Filter = ""
                    rsVisibilidadTarjetaxCanal.MoveFirst
                    rsVisibilidadTarjetaxCanal.Filter = "tCodigoTarjeta = '" & mTarjeta(i, 1) & "'"
                    'verificar si rsVisibilidadTarjetaxCanal tiene contenido
                    If Not rsVisibilidadTarjetaxCanal.EOF Then
                        cmdTarjeta(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal2), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTMonto(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal2), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTPropina(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal2), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                    Else
                        cmdTarjeta(i).Enabled = False
                        txtTMonto(i).Enabled = False
                        txtTPropina(i).Enabled = False
                    End If
                Next i
            Case Is = "03"
                For i = 1 To cantidadBoton
                    rsVisibilidadTarjetaxCanal.Filter = ""
                    rsVisibilidadTarjetaxCanal.MoveFirst
                    rsVisibilidadTarjetaxCanal.Filter = "tCodigoTarjeta = '" & mTarjeta(i, 1) & "'"
                    'verificar si rsVisibilidadTarjetaxCanal tiene contenido
                    If Not rsVisibilidadTarjetaxCanal.EOF Then
                        cmdTarjeta(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal3), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTMonto(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal3), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTPropina(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal3), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                    Else
                        cmdTarjeta(i).Enabled = False
                        txtTMonto(i).Enabled = False
                        txtTPropina(i).Enabled = False
                    End If
                Next i
            Case Is = "04"
                For i = 1 To cantidadBoton
                    rsVisibilidadTarjetaxCanal.Filter = ""
                    rsVisibilidadTarjetaxCanal.MoveFirst
                    rsVisibilidadTarjetaxCanal.Filter = "tCodigoTarjeta = '" & mTarjeta(i, 1) & "'"
                    'verificar si rsVisibilidadTarjetaxCanal tiene contenido
                    If Not rsVisibilidadTarjetaxCanal.EOF Then
                        cmdTarjeta(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal4), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTMonto(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal4), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTPropina(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal4), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                    Else
                        cmdTarjeta(i).Enabled = False
                        txtTMonto(i).Enabled = False
                        txtTPropina(i).Enabled = False
                    End If
                Next i
            Case Is = "05"
                For i = 1 To cantidadBoton
                    rsVisibilidadTarjetaxCanal.Filter = ""
                    rsVisibilidadTarjetaxCanal.MoveFirst
                    rsVisibilidadTarjetaxCanal.Filter = "tCodigoTarjeta = '" & mTarjeta(i, 1) & "'"
                    'verificar si rsVisibilidadTarjetaxCanal tiene contenido
                    If Not rsVisibilidadTarjetaxCanal.EOF Then
                        cmdTarjeta(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal5), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTMonto(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal5), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                        txtTPropina(i).Enabled = IIf(IsNull(rsVisibilidadTarjetaxCanal!lCanal5), False, IIf(rsVisibilidadTarjetaxCanal!lCanal1, True, False))
                    Else
                        cmdTarjeta(i).Enabled = False
                        txtTMonto(i).Enabled = False
                        txtTPropina(i).Enabled = False
                    End If
                Next i
        End Select

    Else
        For i = 1 To cantidadBoton
            cmdTarjeta(i).Enabled = False
            txtTMonto(i).Enabled = False
            txtTPropina(i).Enabled = False
            
            'cmdTarjeta(i).Enabled = True
            'txtTMonto(i).Enabled = True
            'txtTPropina(i).Enabled = True
        Next i
    End If
End Sub
Private Sub Form_Load()
On Error GoTo fin
    If (EstoyDivisionDoc = "1") Then
    Dim hMenu As Long
    '
hMenu = GetSystemMenu(hwnd, 0)
    ' Deshabilitar el menú cerrar del formulario
Call ModifyMenu(hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED, -10, "Close")
    End If
    
   lGuardarVC = False

   wEnter = False
   cmdPago(0).Caption = "Monto en " & sMonN
   cmdPago(6).Caption = "Monto en " & sMonN
   fraPunto.Caption = " " & sClub & " "
   cmdPago(12).Enabled = True ' para pagos por recibos de ingreso
   lGrabaPagoAntes = False
      
   If lCancelacion Then
      cmdOpcion(0).Enabled = False
   End If
    If lDesPagoCheque And sFormulario = "CajaRapida" Then
        Me.fraCheque.Enabled = False
    Else
        Me.fraCheque.Enabled = True
    End If
   
    If lDesPagoOtra And sFormulario = "CajaRapida" Then
        Me.fraDocumento.Enabled = False
    Else
        Me.fraDocumento.Enabled = True
    End If
   If pais = "003" Then
    label(11).Caption = "TC: " & sMonE & ":" & nTC & ", " & sMonE2 & ":" & nTC2 & ", " & sMonE3 & ":" & nTC3
    lblTc.Visible = False
    fraExtranjero.Visible = True
    'cmdOpcion(8).Visible = False
   Else
    lblTc = Format(nTC, "##0.000")
   End If
   If sMonE <> "" And sMonN <> sMonE Then
      cmdPago(2).Caption = "Monto en " & sMonE
      cmdPago(3).Caption = "Monto en " & sMonE
      cmdPago(16).Caption = "Monto en " & sMonE2
      cmdPago(17).Caption = "Monto en " & sMonE3
   Else
      cmdPago(2).Visible = False
      cmdPago(3).Visible = False
      cmdPago(16).Visible = False
      cmdPago(17).Visible = False
      
      If pais = "005" Then
            lblTc.Visible = False
            label(11).Visible = False
            txtME(1).Visible = False
            txtMCH(1).Visible = False
            fraExtranjero.Visible = True
      End If
   End If
   
   If lDesactivaNCFP Then
    cmdPago(11).Enabled = False
   Else
    cmdPago(11).Enabled = True
   End If
   
   
   Set RsCanalesVenta = Lib.OpenRecordset("select * from TCANALVENTA where tCodigoCanalVenta='" + sTipoPedido + "'", Cn)
   
   If Not RsCanalesVenta.BOF Then
    lCuentasXCobrar = IIf(IsNull(RsCanalesVenta!lActivaCuentasXCobrar), False, RsCanalesVenta!lActivaCuentasXCobrar)
   Else
        Dim RsDdoc As Recordset
        Set RsDdoc = Lib.OpenRecordset("select top 1 * from DDOCUMENTO where tDocumento='" + sDocumento + "'", Cn)
        
        Dim RsDPedido As Recordset
        Set RsDPedido = Lib.OpenRecordset("select top 1 * from DPEDIDO where tCodigoPedido='" + RsDdoc!tCodigoPedido + "'", Cn)
        
        If Not RsDPedido.EOF Then
            sTipoPedido = RsDPedido!tTipoPedido
            Set RsCanalesVenta = Lib.OpenRecordset("select * from TCANALVENTA where tCodigoCanalVenta='" + sTipoPedido + "'", Cn)
            lCuentasXCobrar = IIf(IsNull(RsCanalesVenta!lActivaCuentasXCobrar), False, RsCanalesVenta!lActivaCuentasXCobrar)
        Else
            lCuentasXCobrar = False
        End If
        
        Set RsDdoc = Nothing
        Set RsDPedido = Nothing
'    lCuentasXCobrar = False
'    cmdOpcion(8).Enabled = True
   End If
   

  
  If frmDocumentoCorrelativo.SaberPunto = 1 Then
'      cmdOpcion(8).Enabled = False
        If lCuentasXCobrar Then
            cmdOpcion(8).Enabled = True
        Else
            cmdOpcion(8).Enabled = False
        End If
  Else
        If lPagoAntesImpresion Then
                If lCuentasXCobrar Then
                    cmdOpcion(8).Enabled = True
                Else
                    cmdOpcion(8).Enabled = False
                End If
            Else
                cmdOpcion(8).Enabled = False
        End If
  End If
  
   
   frmPago.Caption = "Cancelación del Documento " & Format(sDocumento, "@-@@@@@-@@@@@@@@@")
   
   'para recibos ingreso
     xTabla = dbTemporal(sCaja, 7, "Codigo", "nVarChar(10)", _
                                        "tRecibo", "nVarChar(10)", _
                                            "fFecha", "datetime", _
                                                "tDescripcion", "nVarChar(100)", _
                                                "Moneda", "nVarChar(12)", _
                                                "nMonto", "Float", _
                                                "Descripcion", "Float")
                                                
    Set RsGrillaRecibido = Lib.OpenRecordset("select * from " & xTabla, Cn)
    
   Call ConfGrilla(5, grdGrillaRecibo, "Recibo", 2, "tRecibo", 1000, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1000, 0, 0, "dd/mm/yyyy", _
                                                               "Motivo", 2, "tDescripcion", 1500, 0, 0, "", _
                                                               "Mon", 2, "Moneda", 600, 2, 0, "", _
                                                               "Monto", 2, "nMonto", 900, 1, 0, "##,##0.00")
                                                               
                                                               
    Set grdGrillaRecibo.DataSource = RsGrillaRecibido
   
   'Tarjetas de Crédito
   sTarjeta = dbTemporal(sCaja, 11, "tIndex", "nVarChar(3)", _
                                   "tTipo", "nVarChar(2)", _
                                   "tDescripcion", "nVarChar(30)", _
                                   "tNumero", "nVarChar(16)", _
                                   "tFecha", "nVarChar(4)", _
                                   "nMonto", "Float", _
                                   "nPropina", "Float", _
                                   "nIndex", "Int", _
                                   "lPinPad", "bit", _
                                   "tReferencia", "nVarChar(4)", _
                                   "tNumeroATarjeta", "nVarChar(10)")
                                        
   Set RsGrilla = Lib.OpenRecordset("select * from " & sTarjeta, Cn)
   Call ConfGrilla(7, grdGrilla, "Tarjeta", 2, "tDescripcion", 1000, 0, 0, "", _
                                 "Numero", 2, "tNumero", 800, 0, 0, "", _
                                 "Fec.Venc", 2, "tFecha", 900, 1, 0, "", _
                                 "Monto", 2, "nMonto", 950, 1, 0, "####,###,##0.00", _
                                 "Propina", 2, "nPropina", 850, 1, 0, "####,###,##0.00", _
                                 "NumeroAut.", 2, "tNumeroATarjeta", 600, 0, 0, "", _
                                 "Ref.", 2, "tReferencia", 500, 0, 0, "")
   Set grdGrilla.DataSource = RsGrilla

    'Tarjetas de Crédito VALE CONSUMO
    
    If lActivarValesConsumo Then
    sTarjetaVC = dbTemporal(sCaja, 5, "tIndex", "nVarChar(3)", _
                                    "tTICKET", "nVarChar(50)", _
                                    "tCOD_VALE", "nVarChar(20)", _
                                    "tMONEDA", "nVarChar(10)", _
                                    "nMONTO_VALE", "Float")

    sTarjetaVC_Final = dbTemporal(sCaja, 5, "tIndex", "nVarChar(3)", _
                                    "tTICKET", "nVarChar(50)", _
                                    "tCOD_VALE", "nVarChar(20)", _
                                    "tMONEDA", "nVarChar(10)", _
                                    "nMONTO_VALE", "Float")
    End If
   sMonDoc = "01"
 
    'Tarjetas de Crédito
    Dim cantBoton As Integer
    cantBoton = 0
    If Calcular("select count(*) as codigo from ttarjetacredito where lactivo=1", Cn) > 8 Then
        Isql = "select TTARJETACREDITO.tcodigotarjeta,lCanal1,lCanal2,lCanal3,lCanal4,lCanal5, tdetallado,tresumido,isnull(nfactorretencion,0) nfactorretencion,case when isnull(nboton,0)=0 then 8 else nboton end nboton, lactivo, isnull(lpinpad,0) lpinpad,case when isnull(tCodigoIntegracion,0) = 0 then 8 else tCodigoIntegracion end tCodigoIntegracion from TTARJETACREDITO INNER JOIN TVISIBILIDADTARJETACREDITOXCANAL ON TTARJETACREDITO.tCodigoTarjeta = TVISIBILIDADTARJETACREDITOXCANAL.tCodigoTarjeta  where lactivo=1 "
        Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
        'mastarjetas
        If RsTarjeta.RecordCount <= 8 Then
            Call AsignaTarjeta(8, RsTarjeta, cmdTarjeta())
            ' FERM - Activa Tarjetas x Canal Venta
            cantBoton = 8
        Else
            Call AsignaTarjeta(7, RsTarjeta, cmdTarjeta())
            ' FERM - Activa Tarjetas x Canal Venta
            cantBoton = 7
            
            cmdTarjeta(8).Caption = "OTRAS"
            mTarjeta(8, 1) = 99
            mTarjeta(8, 2) = "OTRAS"
            mTarjeta(8, 3) = "0"
            mTarjeta(8, 4) = "0"
            cmdTarjeta(8).Visible = True
        End If

    Else
   
        Isql = "select * from TTARJETACREDITO where nBoton>0 and lActivo = 1 Order by nBoton"
        Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
        'mastarjetas
         Call AsignaTarjeta(8, RsTarjeta, cmdTarjeta())
         ' FERM - Activa Tarjetas x Canal Venta
            cantBoton = 8
    End If
    'mas tarjetas
      
   ActivaTarjeta False
   Call ActivaTarjetaCreditoxCanal(cantBoton)
   Limpia
   Bloquear False
   If sVar1 = "02" Then
      Set RsPrepago = Lib.OpenRecordset("select * from DPREPAGO where tdocumento='" & sDocumento & "'", Cn)
      If RsPrepago.RecordCount <> 0 Then
      
         RsPrepago.MoveFirst
         Do While Not RsPrepago.EOF
            Select Case RsPrepago!tTipoPago
                   Case Is = "01" 'Efectivo
                        If RsPrepago!tMONEDA = "01" Then
                           nEN = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                           txtME(0).Caption = Format(nEN, "###,###,###,##0.00")
                           nAbonoN = nAbonoN + nEN
                        Else
                           nEE = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                           txtME(1).Caption = Format(nEE, "###,###,###,##0.00")
                           nAbonoE = nAbonoE + nEE
                        End If
                        
                   Case Is = "02" 'Tarjeta
                        nET = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                        txtMT.Caption = Format(nET, "###,###,###,##0.00")
                        nAbonoN = nAbonoN + nET
                        nPT = IIf(IsNull(RsPrepago!nPropina), 0, RsPrepago!nPropina)
                        nPropinaN = nPropinaN + nPT
                        
                        txtPT.Caption = Format(nPT, "###,###,###,##0.00")
                        sTipoTarjeta = IIf(IsNull(RsPrepago!tTarjeta), 0, RsPrepago!tTarjeta)
                        txtVT.Caption = IIf(IsNull(RsPrepago!tFechaVencimiento), "", RsPrepago!tFechaVencimiento)
                        txtNT.Caption = IIf(IsNull(RsPrepago!tNumero), 0, RsPrepago!tNumero)
                        sRefer = IIf(IsNull(RsPrepago!tReferencia), 0, RsPrepago!tReferencia)
                        
                        RsTarjeta.MoveFirst
                        RsTarjeta.Find ("tCodigoTarjeta='" & sTipoTarjeta & "'")
                        If Not RsTarjeta.EOF Then
                           sTitulo = IIf(IsNull(RsTarjeta!tDetallado), "", RsTarjeta!tDetallado)
                           nIndex = IIf(IsNull(RsTarjeta!nBoton), "", RsTarjeta!nBoton)
                        Else
                           sTitulo = ""
                           nIndex = 1
                        End If
                                                     
                        nCorrela = Calcular("select max(tIndex) as Codigo from " & sTarjeta, Cn)
                        If IsNull(nCorrela) Or nCorrela = "" Then
                           nCorrela = "001"
                        Else
                           nCorrela = Lib.Correlativo(nCorrela, 3)
                        End If
                            
                        Isql = "insert into " & sTarjeta & _
                               "( tIndex, tTipo, tDescripcion, tNumero, tFecha, tReferencia, nMonto, nPropina, nIndex, lPinPad ) " & _
                               "Values(    '" & nCorrela & "', " _
                                        & "'" & sTipoTarjeta & "'," _
                                        & "'" & sTitulo & "'," _
                                        & "'" & txtNT.Caption & "'," _
                                        & "'" & txtVT.Caption & "'," _
                                        & "'" & sRefer & "'," _
                                        & nET & ", " _
                                        & nPT & ", " _
                                        & nIndex & ", " & IIf(mTarjeta(nIndex, 3), 1, 0) & ")"
                        Cn.Execute Isql
                        RsGrilla.Requery
               
                        txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
                        txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
                   
                   Case Is = "03" 'Cheque
                        
                        If RsPrepago!tMONEDA = "01" Then
                            nCH = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                            txtMCH(0).Caption = Format(nCH, "###,###,###,##0.00")
                            nAbonoN = nAbonoN + nCH
                            txtNCH.Caption = IIf(IsNull(RsPrepago!tNumero), "", RsPrepago!tNumero)
                            txtBCH.Caption = IIf(IsNull(RsPrepago!tBanco), "", RsPrepago!tBanco)
                        Else
                            nCE = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                            txtMCH(1).Caption = Format(nCE, "###,###,###,##0.00")
                            nAbonoE = nAbonoE + nCE
                            txtNCH.Caption = IIf(IsNull(RsPrepago!tNumero), "", RsPrepago!tNumero)
                            txtBCH.Caption = IIf(IsNull(RsPrepago!tBanco), "", RsPrepago!tBanco)
                        End If
                                           
                   Case Is = "04" 'Documento
                        nDocumento = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                        txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
                        nAbonoN = nAbonoN + nDocumento
                        sOtroTipoCancelacion = RsPrepago!tOtroTipoPago
                        txtDocumento.Caption = Calcular("select Descripcion as Codigo from vTipoCancelacion where Codigo='" & sOtroTipoCancelacion & "'", Cn)
                        
                   Case Is = "05" 'Puntos
                        nTotalPuntos = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                        txtPunto.Caption = Format(nTotalPuntos, "###,###,###,##0.00")
                        nAbonoN = nAbonoN + nTotalPuntos
                        txtCliente.Caption = Calcular("select Cliente as codigo from vDELIVERY where Codigo='" & Trim(IIf(IsNull(RsPrepago!tNumero), "", RsPrepago!tNumero)) & "'", Cn)
            
            End Select
            CalculaTotal
            RsPrepago.MoveNext
         Loop
      End If
   Else
     Dim RsPagoTarjeta As Recordset
     Set RsPagoTarjeta = Lib.OpenRecordset("select * from DPAGOTARJETA where tDocumento='" & sDocumento & "' and tEstadoDocumento='01'", Cn)
     If RsPagoTarjeta.RecordCount > 0 Then
        Do While Not RsPagoTarjeta.EOF
            nET = IIf(IsNull(RsPagoTarjeta!nMonto), 0, RsPagoTarjeta!nMonto)
            txtMT.Caption = Format(nET, "###,###,###,##0.00")
            nAbonoN = nAbonoN + nET
            nPT = IIf(IsNull(RsPagoTarjeta!nPropina), 0, RsPagoTarjeta!nPropina)
            nPropinaN = nPropinaN + nPT
            txtPT.Caption = Format(nPT, "###,###,###,##0.00")
            sTipoTarjeta = IIf(IsNull(RsPagoTarjeta!tTarjeta), 0, RsPagoTarjeta!tTarjeta)
            txtVT.Caption = IIf(IsNull(RsPagoTarjeta!tFechaVencimiento), "", RsPagoTarjeta!tFechaVencimiento)
            txtNT.Caption = IIf(IsNull(RsPagoTarjeta!tNumero), 0, RsPagoTarjeta!tNumero)
            sRefer = IIf(IsNull(RsPagoTarjeta!tReferencia), 0, RsPagoTarjeta!tReferencia)
                             
            RsTarjeta.MoveFirst
            RsTarjeta.Find ("tCodigoTarjeta='" & sTipoTarjeta & "'")
            If Not RsTarjeta.EOF Then
               sTitulo = IIf(IsNull(RsTarjeta!tDetallado), "", RsTarjeta!tDetallado)
               nIndex = IIf(IsNull(RsTarjeta!nBoton), "", RsTarjeta!nBoton)
            Else
               sTitulo = ""
               nIndex = 1
            End If
                                          
            nCorrela = Calcular("select max(tIndex) as Codigo from " & sTarjeta, Cn)
            If IsNull(nCorrela) Or nCorrela = "" Then
               nCorrela = "001"
            Else
               nCorrela = Lib.Correlativo(nCorrela, 3)
            End If
                 
            Isql = "insert into " & sTarjeta & _
                   "( tIndex, tTipo, tDescripcion, tNumero, tFecha, tReferencia, nMonto, nPropina, nIndex, lPinPad ) " & _
                   "Values(    '" & nCorrela & "', " _
                            & "'" & sTipoTarjeta & "'," _
                            & "'" & sTitulo & "'," _
                            & "'" & txtNT.Caption & "'," _
                            & "'" & txtVT.Caption & "'," _
                            & "'" & sRefer & "'," _
                            & nET & ", " _
                            & nPT & ", " _
                            & nIndex & ", 1)"
             Cn.Execute Isql
             RsGrilla.Requery
    
             txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
             txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
             CalculaTotal
             RsPagoTarjeta.MoveNext
         Loop
     End If
   End If
   
   If lVisaNet Then
'      If Not ValidaIP(IpPinPad, IpPort) Then
'         MsgBox "Error de conexión", vbCritical, "VisaNet"
         lVisaNet = False
'      End If
   End If
      
   ObtenerNetos
   
   'CHK TARJETA
   If lClub = True And pais = "002" Then
        cmdPago(1).Visible = True
        txtNAT.Visible = True
        
        lTarjetaActiva = Calcular("select lActivaTarjeta As codigo from TPARAMETRO", Cn)
        If lTarjetaActiva Then
           lBotonTarjeta = False
           ActivaTarjeta True
        End If
   End If
   
   If MESA247 Then
   
    Set RsPrepago = Lib.OpenRecordset("select * from DPREPAGO where tdocumento='" & sPedido & "' OR tcodigopedido='" & sPedido & "'", Cn)
      If RsPrepago.RecordCount <> 0 Then
      
         RsPrepago.MoveFirst
         Do While Not RsPrepago.EOF
            Select Case RsPrepago!tTipoPago
                   Case Is = "02" 'Tarjeta
                        nET = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                        txtMT.Caption = Format(nET, "###,###,###,##0.00")
                        nAbonoN = nAbonoN + nET
                        nPT = IIf(IsNull(RsPrepago!nPropina), 0, RsPrepago!nPropina)
                        nPropinaN = nPropinaN + nPT
                        
                        txtPT.Caption = Format(nPT, "###,###,###,##0.00")
                        sTipoTarjeta = IIf(IsNull(RsPrepago!tTarjeta), 0, RsPrepago!tTarjeta)
                        txtVT.Caption = IIf(IsNull(RsPrepago!tFechaVencimiento), "", RsPrepago!tFechaVencimiento)
                        txtNT.Caption = IIf(IsNull(RsPrepago!tNumero), 0, RsPrepago!tNumero)
                        sRefer = IIf(IsNull(RsPrepago!tReferencia), 0, RsPrepago!tReferencia)
                        
                        RsTarjeta.MoveFirst
                        RsTarjeta.Find ("tCodigoTarjeta='" & sTipoTarjeta & "'")
                        If Not RsTarjeta.EOF Then
                           sTitulo = IIf(IsNull(RsTarjeta!tDetallado), "", RsTarjeta!tDetallado)
                           nIndex = IIf(IsNull(RsTarjeta!nBoton), "", RsTarjeta!nBoton)
                        Else
                           sTitulo = ""
                           nIndex = 1
                        End If
                                                     
                        nCorrela = Calcular("select max(tIndex) as Codigo from " & sTarjeta, Cn)
                        If IsNull(nCorrela) Or nCorrela = "" Then
                           nCorrela = "001"
                        Else
                           nCorrela = Lib.Correlativo(nCorrela, 3)
                        End If
                            
                        Isql = "insert into " & sTarjeta & _
                               "( tIndex, tTipo, tDescripcion, tNumero, tFecha, tReferencia, nMonto, nPropina, nIndex, lPinPad ) " & _
                               "Values(    '" & nCorrela & "', " _
                                        & "'" & sTipoTarjeta & "'," _
                                        & "'" & sTitulo & "'," _
                                        & "'" & txtNT.Caption & "'," _
                                        & "'" & txtVT.Caption & "'," _
                                        & "'" & sRefer & "'," _
                                        & nET & ", " _
                                        & nPT & ", " _
                                        & nIndex & ", " & IIf(mTarjeta(nIndex, 3), 1, 0) & ")"
                        Cn.Execute Isql
                        RsGrilla.Requery
               
                        txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
                        txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")

                End Select
            CalculaTotal
            RsPrepago.MoveNext
         Loop
      End If
      
   End If
   
   If pais = "003" Then
    fraCheque.Visible = False
    fraEfectivo.Height = 3945
   End If
   
   If lIntMercadoPagoQREstatico Then
        Me.cmdMercadoPago.Visible = True
   End If
   
   If lIntPagoEfectivo Then
        frmPagoEfectivo.Enabled = True
   End If
    
    If IntFPAY Then
        frmFPAY.Enabled = True
   End If
   
   On Error Resume Next
   
   If lActivaReciboSaldoQr Then
        txtQRRecibo.SetFocus
   End If
   
   '----------
   Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "Load_Pago", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Public Sub Limpia()

 Set RsGrillaVConsumo = Nothing
 
 If lActivarValesConsumo Then
 Cn.Execute "delete from " & sTarjetaVC & ""
 Cn.Execute "delete from " & sTarjetaVC_Final & ""
 End If

 nEN = 0
 nEE = 0
 nEE2 = 0
 nEE3 = 0
 nCH = 0
 nCE = 0
 nET = 0
 nPT = 0
 nDocumento = 0
 nTotalPuntos = 0
 nAbonoN = 0
 nAbonoE = 0
 nAbonoE2 = 0
 nAbonoE3 = 0
 nSaldo = nCargo - nAbonoN
 nVueltoN = 0
 nVueltoE = 0
 nVueltoE2 = 0
 nVueltoE3 = 0
 nPropinaN = 0
 nTotalN = 0
 nVueltoE = 0
 nPropinaE = 0
 nTotalE = 0
 
 sCortesia = ""
 txtCortesia.Caption = ""
 
 txtCargo(0).Caption = Format(nCargo, "###,###,##0.00")
 txtAbono(0).Caption = Format(nAbonoN, "###,###,##0.00")
 txtSaldo(0).Caption = Format(nSaldo, "###,###,##0.00")
 txtVuelto(0).Caption = Format(nVueltoN, "###,###,##0.00")
 TxtTotal(0).Caption = Format(nTotalN, "###,###,##0.00")
 txtCliente.Caption = ""
 txtDocumento.Caption = ""
 
 If sMonE <> "" And sMonN <> sMonE Then
    txtCargo(1).Caption = Format(nCargo / nTC, "###,###,##0.00")
    txtAbono(1).Caption = Format(nAbonoN / nTC, "###,###,##0.00")
    txtSaldo(1).Caption = Format(nSaldo / nTC, "###,###,##0.00")
    txtVuelto(1).Caption = Format(nVueltoN / nTC, "###,###,##0.00")
    TxtTotal(1).Caption = Format(nTotalN / nTC, "###,###,##0.00")
 Else
    txtCargo(1).Visible = False
    txtAbono(1).Visible = False
    txtSaldo(1).Visible = False
    txtVuelto(1).Visible = False
    TxtTotal(1).Visible = False
 End If
 
 txtME(0).Caption = "0.00"
 txtME(1).Caption = "0.00"
 txtME(2).Caption = "0.00"
 txtME(3).Caption = "0.00"
 txtMT.Caption = "0.00"
 txtPT.Caption = "0.00"
 
 txtMCH(0).Caption = "0.00"
 txtMCH(1).Caption = "0.00"
 txtBCH.Caption = ""
 txtNCH.Caption = ""
 Dim i As Integer
 For i = 1 To 8
     txtTMonto(i).Caption = "0.00"
     txtTPropina(i).Caption = "0.00"
 Next i
 
 If nPuerto > 0 And sFormulario = "CajaRapida" Then
    Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(nAbonoN + (nAbonoE * nTC) + (nAbonoE2 * nTC2) + (nAbonoE3 * nTC3), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(nVueltoN, "##,##0.00"), 8), nPuerto, "N"
 End If

End Sub

Public Sub CalculaTotal()
   txtCargo(0).Caption = Format(nCargo, "###,###,##0.00")
   txtCargo(1).Caption = Format(nCargo / nTC, "###,###,##0.00")
    
   If pais = "003" Then
    txtAbono(0).Caption = Format(nAbonoN + (nAbonoE * nTC) + (nAbonoE2 * nTC2) + (nAbonoE3 * nTC3), "###,###,##0.00")
   Else
    txtAbono(0).Caption = Format(nAbonoN, "###,###,##0.00")
   End If
   txtAbono(1).Caption = Format(nAbonoE, "###,###,##0.00")
      
   nSaldo = nCargo - nAbonoN
   If nAbonoE > 0 Then
      nSaldo = nSaldo - (nAbonoE * nTC)
   End If
    If nAbonoE2 > 0 Then
      nSaldo = nSaldo - (nAbonoE2 * nTC2)
   End If
    If nAbonoE3 > 0 Then
      nSaldo = nSaldo - (nAbonoE3 * nTC3)
   End If
   
   'nSaldo = Round(IIf((nCargo - nAbonoN - (nAbonoE * nTC)) >= 0, nCargo - nAbonoN - (nAbonoE * nTC), 0), 2)
   nSaldo = Round(IIf(nSaldo >= 0, nSaldo, 0), 2)
   
   txtSaldo(0).Caption = Format(nSaldo, "###,###,##0.00")
   txtSaldo(1).Caption = Format(nSaldo / nTC, "###,###,##0.00")
   
   nTotalN = nAbonoN + nPropinaN + (nAbonoE * nTC) + (nPropinaE * nTC) + (nAbonoE2 * nTC2) + (nAbonoE3 * nTC3)
   nTotalE = (nAbonoN / nTC) + (nPropinaN / nTC) + nAbonoE + nPropinaE
   
   TxtTotal(0).Caption = Format(nTotalN, "###,###,##0.00")
   TxtTotal(1).Caption = Format(nTotalE, "###,###,##0.00")
   
   txtPropina(0).Caption = Format(nPropinaN, "###,###,##0.00")
   txtPropina(1).Caption = Format(nPropinaE, "###,###,##0.00")
   
   nVueltoN = IIf((nAbonoN + (nAbonoE * nTC) + (nAbonoE2 * nTC2) + (nAbonoE3 * nTC3)) - nCargo > 0, (nAbonoN + (nAbonoE * nTC) + (nAbonoE2 * nTC2) + (nAbonoE3 * nTC3)) - nCargo, 0)
   nVueltoE = IIf((nAbonoE + (nAbonoN / nTC)) - nCargo / nTC > 0, (nAbonoE + (nAbonoN / nTC)) - nCargo / nTC, 0)
   If pais = "003" Then
   Dim CalculoVuel As String
   CalculoVuel = (nAbonoE2 + (nAbonoN / nTC2) + (nAbonoE * nTC) / nTC2) - nCargo / nTC2
   nVueltoE2 = IIf(CalculoVuel > 0, CalculoVuel, 0)
   
   CalculoVuel = (nAbonoE3 + (nAbonoN / nTC3) + (nAbonoE2 * nTC2) / nTC3 + (nAbonoE * nTC) / nTC3) - nCargo / nTC3
   nVueltoE3 = IIf(CalculoVuel > 0, CalculoVuel, 0)
   End If
   txtVuelto(0).Caption = Format(nVueltoN, "###,###,##0.00")
   txtVuelto(1).Caption = Format(nVueltoE, "###,###,##0.00")
   
   If nPuerto > 0 And sFormulario = "CajaRapida" Then
      Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(nAbonoN + (nAbonoE * nTC) + (nAbonoE2 * nTC2) + (nAbonoE3 * nTC3), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(nVueltoN, "##,##0.00"), 8), nPuerto, "N"
   End If
End Sub

Public Sub ActivaTarjeta(Activa As Boolean)
On Error GoTo fin
   Dim i As Integer
   cmdPago(1).Enabled = Activa    'Numero Autorizacion Tarjeta
   cmdPago(7).Enabled = Activa    'Numero
   cmdPago(8).Enabled = Activa    'Vencimiento
   cmdPago(9).Enabled = Activa    'Monto
   cmdPago(10).Enabled = Activa   'Propina
   
   cmdOpcion(4).Enabled = Activa  'Agregar
   cmdOpcion(5).Enabled = Activa  'Modificar
   
   For i = 1 To 8
       cmdTarjeta(i).Enabled = Not Activa
   Next i
   
   If lClub = True And pais = "002" Then
        If lTarjetaActiva = True Then
             Dim RsTarjetaTemp As Recordset
             Isql = "select tcodigotarjeta,case when isnull(nboton,0)=0 then 99 else nboton end nboton from TTARJETACREDITO where lactivo=1 and ISNULL(tPrefijo,'') = ''"
             Set RsTarjetaTemp = Lib.OpenRecordset(Isql, Cn)
             
             For i = 1 To RsTarjetaTemp.RecordCount
                 cmdTarjeta(RsTarjetaTemp.Fields(1)).Enabled = True
                 RsTarjetaTemp.MoveNext
             Next i
        End If
   End If
   Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "FALLA ActivaTarjeta ", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Public Sub AsignarTotal()
On Error GoTo fin
   If RsGrilla.RecordCount <> 0 Then
      txtMT.Caption = Format(RsGrilla!nMonto, "###,###,##0.00")
      txtVT.Caption = RsGrilla!tFecha
      sRefer = RsGrilla!tReferencia
      txtPT.Caption = Format(RsGrilla!nPropina, "###,###,##0.00")
      txtNT.Caption = RsGrilla!tNumero
      cmdPago(10).Enabled = IIf(sRefer <> "", True, False)
      nIndex = RsGrilla!nIndex
      lblTarjetaC.Caption = RsGrilla!tDescripcion
   Else
      txtMT.Caption = ""
      txtVT.Caption = ""
      txtPT.Caption = ""
      txtNT.Caption = ""
      txtNAT.Caption = ""
      
      cmdPago(10).Enabled = False
      sRefer = ""
      nIndex = 1
   End If
   Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "AsignarTotal", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sTarjeta
   Set RsTarjeta = Nothing
   Set RsGrilla = Nothing
   Set rst1 = Nothing
   Set RsT2 = Nothing
   Set frmPago = Nothing
  
   
   
'     If (EstoyDivisionDoc = "1") Then
'        Dim xSuma As String
'        'xSuma = Calcular("select count(*) as Codigo FROM MDOCUMENTO where tEstadoItem = 'N' and tCodigoPedido='" & sPedido & "'", Cn)
'        xSuma = Calcular("select count(*) as Codigo from MDOCUMENTO where tEstadoDocumento = '01' and tDocumento  in ( select distinct tDocumento from DDOCUMENTO where tCodigoPedido='" & sPedido & "')", Cn)
'          If xSuma > 0 Then
'              If MsgBox("Aun se tiene pendiente por pagar " + xSuma + " Documentos, esta seguro que desea salir del proceso?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
'                    Exit Sub
'                Else
'                    EstoyDivisionDoc = "0"
'                    wEnter = False
'                    'Unload Me
'              End If
'          End If
'     Else
          If lPagoAntesImpresion Then
            If lGrabaPagoAntes Then
                wEnter = True
            Else
                wEnter = False
            End If
         End If
    'End If
   
   
   
End Sub


Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   AsignarTotal
End Sub

Public Sub AsignaTarjeta(nBoton As Integer, RsAsigna As Recordset, cmdBoton As Object)
On Error GoTo fin
   Screen.MousePointer = vbHourglass
   Dim i As Integer
   With RsAsigna
        If .RecordCount > 0 Then
           For i = 1 To nBoton
               .MoveFirst
               .Find "nboton = " & Trim(str(i))
               If Not .EOF Then
                  mTarjeta(i, 1) = !tCodigoTarjeta
                  mTarjeta(i, 2) = IIf(IsNull(!tResumido), "", !tResumido)
                  mTarjeta(i, 3) = IIf(IsNull(!lPinPad), 0, !lPinPad)
                  mTarjeta(i, 4) = IIf(IsNull(!tCodigoIntegracion), 0, !tCodigoIntegracion)
                  cmdBoton(i).Visible = True
                  cmdBoton(i).Caption = mTarjeta(i, 2)
                Else
                  cmdBoton(i).Visible = False
                  txtTMonto(i).Visible = False
                  txtTPropina(i).Visible = False
                End If
           Next i
        Else
           For i = 1 To nBoton
               cmdBoton(i).Visible = False
               txtTMonto(i).Visible = False
               txtTPropina(i).Visible = False
           Next i
        End If
   End With
   Screen.MousePointer = vbDefault
   Exit Sub
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "AsignaTarjeta", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Sub

Public Sub Bloquear(Estado As Boolean)
   Dim RTemps As Recordset
   Dim i As Integer
   Dim xttipopedido As String
   Dim cantTarj As Integer
   xttipopedido = Calcular("Select ttipopedido as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
   
   Isql = "SELECT lVisibilidad,lCanal1,lCanal2,lCanal3,lCanal4,lCanal5, " & _
           "TTC.tCodigoTarjeta,TTC.tDetallado,TTC.nBoton " & _
           "FROM TVISIBILIDADTARJETACREDITOXCANAL TVT " & _
           "INNER JOIN TTARJETACREDITO TTC ON TVT.tCodigoTarjeta = TTC.tcodigotarjeta"
        Set RTemps = Lib.OpenRecordset(Isql, Cn)
   
   cantTarj = RTemps.RecordCount
   
   For i = 1 To 8

        RTemps.Filter = "nBoton = " & i
        
        If RTemps.EOF Then
            Set RTemps = Nothing
            Exit For
        End If
        
        If RTemps!lVisibilidad Then
        
            Select Case xttipopedido
            
                    Case "01"
                        If RTemps!lCanal1 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                            If i = 8 And cantTarj > 8 Then
                                cmdTarjeta(i).Enabled = Not Estado
                            End If
                        End If
                        
                    Case "02"
                        If RTemps!lCanal2 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                            If i = 8 And cantTarj > 8 Then
                                cmdTarjeta(i).Enabled = Not Estado
                            End If
                        End If
                    
                    Case "03"
                        If RTemps!lCanal3 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                            If i = 8 And cantTarj > 8 Then
                                cmdTarjeta(i).Enabled = Not Estado
                            End If
                        End If
                    
                    Case "04"
                        If RTemps!lCanal4 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                            If i = 8 And cantTarj > 8 Then
                                cmdTarjeta(i).Enabled = Not Estado
                            End If
                        End If
                    
                    Case "05"
                        If RTemps!lCanal5 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                            If i = 8 And cantTarj > 8 Then
                                cmdTarjeta(i).Enabled = Not Estado
                            End If
                        End If
            End Select
        Else
            cmdTarjeta(i).Enabled = Estado
        End If
   Next i
   
   cmdPago(0).Enabled = Not Estado
   cmdPago(2).Enabled = Not Estado
   cmdPago(4).Enabled = Not Estado
   cmdPago(5).Enabled = Not Estado
   cmdPago(6).Enabled = Not Estado
   cmdPago(3).Enabled = Not Estado
   cmdPago(11).Enabled = Not Estado
   cmdPago(12).Enabled = Not Estado
   cmdPago(13).Enabled = Not Estado
   cmdPago(14).Enabled = Not Estado
End Sub

Private Sub ModifcarEstadoDeliveryCabecera(ByVal qDocumento As String) 'pp
    On Error GoTo ErrorHandler
    'Central Delivery-Motorizado--------------------------------------------------pp
    Isql = "Select Distinct P.tCodigoPedidoCD from DDocumento as D Inner Join MPedido AS P On D.tCodigoPedido = P.tCodigoPedido Where tDocumento = '" + qDocumento + "'"
    Dim RsCodigoPCD As ADODB.Recordset
    Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
    If Not RsCodigoPCD.EOF Then
        If Not IsNull(RsCodigoPCD!tCodigoPedidoCD) Then
'            Call ModifcarEstadoDeliveryCabecera(RsCodigoPCD!tCodigoPedidoCD, "3", txtMotorizado.Caption)
             
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
                
            CnCD.Execute "usp_CD_Modificar_EstadoDelivery_Cabecera 4, '" + RsCodigoPCD!tCodigoPedidoCD + "',''"
            CnCD.Close
            
        End If
    End If
    ''''''''''''''''''''''''''''''
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub


Private Sub ModificarEstadoPedidoWebAp(ByVal qDocumento As String)
    On Error GoTo ErrorHandler
'    Isql = "Select Distinct P.tCodigoPedidoWebAp from DDocumento as D Inner Join MPedido AS P On D.tCodigoPedido = P.tCodigoPedido Where tDocumento = '" + qDocumento + "'"
'    Dim RsCodigoPCD As ADODB.Recordset
'    Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
'    If Not RsCodigoPCD.EOF Then
'        If Not IsNull(RsCodigoPCD!tCodigoPedidoWebAp) Then
'            Dim CnWAP As Connection
'            Set CnWAP = New Connection
'            CnWAP.Provider = "SQLOLEDB"
'            CnWAP.CursorLocation = adUseServer
'            CnWAP.ConnectionString = "User ID=" & sUserName & _
'                                    ";password=" & sUserPassword & _
'                                    ";Data Source=" & sRutaWAP & _
'                                    ";Initial Catalog=" & sMDBWAP
'            CnWAP.CommandTimeout = 250
'            CnWAP.Open
'
'            CnWAP.Execute "APP_ACTUALIZAR_MPEDIDO_ESTADO '" & RsCodigoPCD!tCodigoPedidoWebAp & "','8'"
'            CnWAP.Close
'        End If
'    End If
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub


'control de pagos
Public Function validadCantidadMontosPagos() As Boolean
On Error GoTo fin
    validadCantidadMontosPagos = True
    Dim efectivo As Boolean
    Dim dolares As Boolean
    Dim cheke As Boolean
    Dim chekedolares As Boolean
    Dim montotarjeta As Double
    Dim tarjeta As Boolean
    Dim Cantidad As Integer
    Dim PUNTO As Boolean
    Dim otroDoc As Boolean
    Cantidad = 0
    If val(txtME(0).Caption) > 0 Then
        efectivo = True
        Cantidad = Cantidad + 1
    End If
    
    If val(txtME(1).Caption) > 0 Then
        dolares = True
        Cantidad = Cantidad + 1
    End If
    If val(txtMCH(0).Caption) > 0 Then
        cheke = True
        Cantidad = Cantidad + 1
    End If
    If val(txtMCH(1).Caption) > 0 Then
        chekedolares = True
        Cantidad = Cantidad + 1
    End If
    If Calcular("select sum(nmonto) as codigo from " & sTarjeta, Cn) > 0 Then
        tarjeta = True
        Cantidad = Cantidad + 1
    End If
    If val(Me.txtPunto.Caption) > 0 Then
        PUNTO = True
        Cantidad = Cantidad + 1
    End If
    If val(Me.txtDocMonto.Caption) > 0 Then
        otroDoc = True
        Cantidad = Cantidad + 1
    End If

    If Cantidad > 1 Then
        If efectivo = True Then
            If val(txtME(0).Caption) >= nCargo Then
                    validadCantidadMontosPagos = False
                    Exit Function
            End If
        End If
    
        If dolares = True Then
            If val(txtME(1).Caption) * nTC >= nCargo Then
                    validadCantidadMontosPagos = False
                    Exit Function
            End If
        End If
        
        If cheke = True Then
            If val(txtMCH(0).Caption) >= nCargo Then
                    validadCantidadMontosPagos = False
                    Exit Function
            End If
        End If
        If chekedolares = True Then
            If val(txtMCH(1).Caption) >= nCargo Then
                    validadCantidadMontosPagos = False
                    Exit Function
            End If
        End If
        If tarjeta = True Then
             If Calcular("select sum(nmonto)  as codigo from " & sTarjeta, Cn) >= nCargo Then
                    validadCantidadMontosPagos = False
                    Exit Function
                End If
        End If
        If PUNTO = True Then
            If val(Me.txtPunto.Caption) >= nCargo Then
                    validadCantidadMontosPagos = False
                    Exit Function
            End If
        
        End If
        
        If otroDoc = True Then
            If val(Me.txtDocMonto.Caption) >= nCargo Then
                    validadCantidadMontosPagos = False
                    Exit Function
            End If
        End If
    Else
        validadCantidadMontosPagos = True
        Exit Function
    End If
    Exit Function
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "ValidadCantidadMontosPagos", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Function
'control de pagos
Public Function validadCantidadMontosPagos2() As Boolean
On Error GoTo fin
    validadCantidadMontosPagos2 = True
    Dim montotarjeta As Double
 
    montotarjeta = Calcular("select sum(isnull(nmonto,0)) as codigo from " & sTarjeta, Cn)
    
    'If (Val(Me.txtMCH(0).Caption) + Val(Me.txtPunto.Caption) + Val(Me.txtDocMonto.Caption)) >= nCargo And (Val(Me.txtME(0).Caption) + montotarjeta + (Val(Me.txtME(1).Caption) * nTC) > 0) Then
    If (val(Me.txtMCH(0).Caption) + (val(Me.txtMCH(1).Caption) * nTC) + val(Me.txtPunto.Caption) + val(Me.txtDocMonto.Caption)) >= nCargo And (val(Me.txtME(0).Caption) + montotarjeta + (val(Me.txtME(1).Caption) * nTC) > 0) Then
            validadCantidadMontosPagos2 = False
    End If
    Exit Function
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "VlidadCantidadMontosPagos2", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Function

'control de pagos
Public Function validadCantidadMontosPagos3() As Boolean
On Error GoTo fin
    validadCantidadMontosPagos3 = True
    Dim montotarjeta As Double
 
    montotarjeta = Calcular("select sum(isnull(nmonto,0)) as codigo from " & sTarjeta, Cn)
    
    If (val(Me.txtMCH(0).Caption) + (val(Me.txtMCH(1).Caption) * nTC) + val(Me.txtPunto.Caption) + val(Me.txtDocMonto.Caption)) > 0 And (val(Me.txtME(0).Caption) + montotarjeta + (val(Me.txtME(1).Caption) * nTC) >= nCargo) Then
            validadCantidadMontosPagos3 = False
    End If
    Exit Function
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "ValidadCantidadMontosPagos3", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Function

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
End Function
'diaContable

Public Function InsertaSiab(tTarjeta As String, nMonto As Double) As Boolean
    sSYBASE = Trim(LeerIni(App.path + "\INFOREST.INI", "CONEXION", "SYBASE", ""))
    Set CnSiab = New Connection
    CnSiab.Provider = SYPROV '"ASAProv.80"
    CnSiab.CursorLocation = adUseServer
    CnSiab.ConnectionString = sSYBASE
    CnSiab.CommandTimeout = 250
    On Error GoTo ErrorEcuador
    CnSiab.Open
    Dim sZona As String
    Dim sDocCar As String
    Dim nSecuencial As Double
    
    sZona = Calcular("select destino as Codigo from EQUIVALENCIA where Relacion = 'ZONA' and origen = '" & sCaja & "'", CnSiab)
    'sDocCar = Calcular("select destino as Codigo from EQUIVALENCIA where Relacion = 'DOC_CAR' and origen = '" & sCaja & "'", CnSiab)
    sDocCar = "F10"
    nSecuencial = Calcular("select max(num_trans) as Codigo from TARJ_CONSUMO_MOV", CnSiab)
    'update TARJ_CONSUMO_MAESTRO set saldo = saldo-@nMonto,Utilizado = Utilizado+@nMonto where NUM_TARJETA = @tTarjeta
    
    Isql = "insert into TARJ_CONSUMO_MOV " & _
           "(CIA_CODIGO,NUM_TRANS, TIPO_TARJ,NUM_TARJETA,TIPO_TRANSAC,AGENCIA_ID, ZONA_ID,FECHA_TRANS,VALOR,OBSERVACION,ESTADO," & _
           "USER_ID,ULT_MODIF,TIP_DOC_CAJA,NUM_DOC_CAJ,LINEA_DOC_CAJ, TIP_DOC_CAR,NUM_DOC_CAR) values('01'," & nSecuencial + 1 & ",'01', " & _
           "'" & tTarjeta & "','VTA',1, '" & sZona & "', today(*), " & nMonto & ", 'FACTURA: " & Mid(sDocumento, 2, 6) & Mid(sDocumento, 8) & "','A'," & _
           "'" & sUsuario & "', getdate(*),null,null,null, '" & sDocCar & "' , " & Mid(sDocumento, 5, 3) & Mid(sDocumento, 10, 7) & ")"
    On Error GoTo ErrorEcuador2
    CnSiab.Execute Isql
    CnSiab.Close
    On Error GoTo 0
    InsertaSiab = True
Exit Function
ErrorEcuador:
    InsertaSiab = False
ErrorEcuador2:
    MsgBox "Error fatal: " & Err.Description & Chr(13) & Isql, vbCritical, "Error de insercion en el SIAB"
    InsertaSiab = False
End Function

Public Function ObtenerNetos()
 On Error GoTo fin
    Dim RsMDocumentox As Recordset
    Dim RsDocumentox As Recordset
    Dim RsParametrox As Recordset
    Dim simp1 As String
    Dim simp2 As String
    Dim IMP1 As Double
    Dim Imp2 As Double
    Dim impxT As Double
    Dim impx1 As Double
    Dim impx2 As Double
    Dim impG As Double
    Dim xDescuento As Double
    Dim xRecargo As Double
    Dim xSubTotal As Double
    Dim xDiferencia As Double
    Dim xAumento As Double
    Dim zImpuesto1 As Double
    Dim zImpuesto2 As Double
    Dim zImpuesto3 As Double
    Dim zTotal As Double
    
    Set RsParametrox = Lib.OpenRecordset("select * from TPARAMETRO", Cn)
    
    lActivaReciboSaldoQr = IIf(IsNull(RsParametrox!lActivaReciboSaldoQr), 0, RsParametrox!lActivaReciboSaldoQr)
    
    simp1 = RsParametrox!tImpuesto1
    simp2 = RsParametrox!tImpuesto2
    IMP1 = RsParametrox!IMPUESTO1
    Imp2 = RsParametrox!IMPUESTO2
    impxT = 1 + (IMP1 / 100) + (Imp2 / 100)
    impx1 = 1 + (IMP1 / 100)
    impx2 = 1 + (Imp2 / 100)
    
    impG = 0
    xDiferencia = 0
    xDescuento = 0
    xAumento = 0
    xRecargo = 0
    zImpuesto1 = 0
    zImpuesto2 = 0
    zImpuesto3 = 0
    zTotal = 0
    
    Set RsDocumentox = Lib.OpenRecordset("select * from DDOCUMENTO where tDocumento='" & sDocumento & "'", Cn)
    Set RsMDocumentox = Lib.OpenRecordset("select * from MDOCUMENTO where tDocumento='" & sDocumento & "'", Cn)
    
    If RsDocumentox.RecordCount > 0 Then
        Do While Not RsDocumentox.EOF
            If RsDocumentox!nprecioImpuesto1 > 0 And RsDocumentox!nprecioImpuesto2 > 0 Then
                  impG = impxT
            ElseIf RsDocumentox!nprecioImpuesto1 > 0 And RsDocumentox!nprecioImpuesto2 = 0 Then
                  impG = impx1
            ElseIf RsDocumentox!nprecioImpuesto1 = 0 And RsDocumentox!nprecioImpuesto2 > 0 Then
                  impG = impx2
            Else
                  impG = 1
            End If
            
            xSubTotal = xSubTotal + ((RsDocumentox!nPrecioOficial * RsDocumentox!nCantidad) / impG)
            xDiferencia = ((RsDocumentox!nPrecioOficial * RsDocumentox!nCantidad) / impG) - (RsDocumentox!nVenta - RsDocumentox!nImpuesto1 - RsDocumentox!nImpuesto2)
            xDescuento = xDescuento + xDiferencia
            xAumento = (RsDocumentox!nVenta - RsDocumentox!nImpuesto1 - RsDocumentox!nImpuesto2) - ((RsDocumentox!nPrecioOficial * RsDocumentox!nCantidad) / impG)
            xRecargo = xRecargo + xAumento
            
            zImpuesto1 = zImpuesto1 + RsDocumentox!nImpuesto1
            zImpuesto2 = zImpuesto2 + RsDocumentox!nImpuesto2
            zImpuesto3 = zImpuesto3 + RsDocumentox!nImpuesto3
            
            zTotal = zTotal + RsDocumentox!nVenta
            
            RsDocumentox.MoveNext
        Loop
    End If
   
    
    txtSubTotal.Caption = Format(xSubTotal, "##,###,##0.00")
    
    If RsMDocumentox!nDescuento > 0 Then
       TxtDescuento.Caption = Format(xDescuento, "##,###,##0.00")
    Else
       TxtDescuento.Caption = Format(0, "##,###,##0.00")
    End If
    
    If RsMDocumentox!nRecargo > 0 Then
       TxtRecargo.Caption = Format(xRecargo, "##,###,##0.00")
    Else
       TxtRecargo.Caption = Format(0, "##,###,##0.00")
    End If
    
    If zImpuesto1 > 0 Then
       lblImpuesto1.Caption = simp1 & " " & Format(IMP1, "##0.00") & "% : "
       txtImpuesto1.Caption = Format(zImpuesto1, "##,###,##0.00")
    Else
       lblImpuesto1.Caption = simp1 & " " & Format(IMP1, "##0.00") & "% : "
       txtImpuesto1.Caption = Format(0, "##,###,##0.00")
    End If
    
    If zImpuesto2 > 0 Then
       lblImpuesto2.Caption = simp2 & " " & Format(Imp2, "##0.00") & "% : "
       txtImpuesto2.Caption = Format(zImpuesto2, "##,###,##0.00")
    Else
       lblImpuesto2.Caption = simp2 & " " & Format(Imp2, "##0.00") & "% : "
       txtImpuesto2.Caption = Format(0, "##,###,##0.00")
    End If
    
    txtTotalNeto.Caption = Format(zTotal, "##,###,##0.00")
    
    dsTotalNeto = txtTotalNeto.Caption
    dsImpuesto2 = txtImpuesto2.Caption
    dsImpuesto1 = txtImpuesto1.Caption
    dsRecargo = TxtRecargo.Caption
    dsDescuento = TxtDescuento.Caption
    dsSubTotal = txtSubTotal.Caption
    Exit Function
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "ObtenerNetos", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje

End Function

Public Function obtieneFilasSeleccionadas() As Integer
On Error GoTo fin
    obtieneFilasSeleccionadas = 0
    Dim m As Integer
    Dim k As Integer
    'Dim nTotal As Double
   ' nTotal = 0
    k = 0
      For m = 1 To Me.grdGrilla.ApproxCount
               If grdGrilla.Columns(9).CellValue(m) = True Then
                    k = k + 1
                    'nTotal = nTotal + Val(grdGrilla.Columns(7).CellValue(m))
                End If
        Next m
    obtieneFilasSeleccionadas = k
    Exit Function
fin:
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGODOCUMENTO", "PAGO DE DOCUMENTO", "", "", sDocumento, Error, "", "obtineFilasSeleccionadas", sUsuario)
    MsgBox "Error: " & Error + " | Proceso: Pagos", vbCritical, sMensaje
    obtieneFilasSeleccionadas = 0
End Function



Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
    If lActivaReciboSaldoQr Then
        Me.txtQRRecibo.SetFocus
    End If
End Sub

Private Sub Timer_CashDro_Timer()
On Error GoTo fin
    If Calcular("If exists (select 1 from INTEGRACION_CASHDRO where documento='" & sDocumento & "') select '1' codigo else select '0' codigo;", Cn) = "1" Then
        
        Me.Timer_CashDro.Enabled = False
        Me.frmMensaje.Visible = False
        
        sDescrip = Calcular("select isnull(Monto,'0') Codigo from INTEGRACION_CASHDRO where documento='" & sDocumento & "'", Cn)
        If (sDescrip <> "0") Then
            tCodigoIntegracion = 4 '-- Codigo de la tabla tintegraciones
            wEnter = True
            nEN = IIf(wEnter = True, sDescrip, nEN)
            txtME(0).Caption = Format(nEN, "###,###,###,##0.00")
            nAbonoN = nAbonoN + nEN
            CalculaTotal
            'MsgBox "Se realizo el deposito correcto en el CashDrow", vbInformation
            cmdOpcion_Click (2)
        Else
            sDescrip = ""
        End If
    End If
fin:
End Sub

Private Sub Timer_Fpay_Timer()
On Error GoTo fin
    If Calcular("select count(*) as codigo from dpagodocumento where tdocumento='" & sDocumento & "'", Cn) > 0 Then
        MsgBox "Pago de documento " & sDocumento & " completado!!!", vbInformation
        Me.Timer_Fpay.Enabled = False
        Me.Timer_Fpay.Interval = 0
        Me.BtnCancelarFPAY.Enabled = False
        If lPagoAntesImpresion Then
            lGrabaPagoAntes = True
        End If
        cmdOpcion_Click (99)
    ElseIf Calcular("select count(*) as codigo from TRANSACCIONES_FPAY where tipo='EXPIRADO' and estado <> 3 and documento='" & sDocumento & "'", Cn) > 0 Then
        MsgBox "El QR asociado al " & sDocumento & " a expirado, debe generar otra intencion de pago !!!", vbExclamation
        frmMensaje.Visible = False
        Me.Timer_Fpay.Enabled = False
        Me.Timer_Fpay.Interval = 0
        
        Cn.Execute "update TRANSACCIONES_FPAY set estado = 3 where estado = 0 and documento ='" & sDocumento & "'"
    
    ElseIf Calcular("select count(*) as codigo from TRANSACCIONES_FPAY where tipo='RECHAZADO' and estado <> 3 and documento='" & sDocumento & "'", Cn) > 0 Then
        MsgBox "El QR asociado al " & sDocumento & " a sido Rechazado, debe generar otra intencion de pago !!!", vbExclamation
        frmMensaje.Visible = False
        Me.Timer_Fpay.Enabled = False
        Me.Timer_Fpay.Interval = 0
        Frame1.Enabled = True
        Me.fraEfectivo.Enabled = True
        Cn.Execute "update TRANSACCIONES_FPAY set estado = 3 where estado = 0 and documento ='" & sDocumento & "'"
            
    End If
Exit Sub
fin:
End Sub

Private Sub Timer_MercadoPago_Timer()
On Error GoTo fin
    If Calcular("select count(*) as codigo from dpagodocumento where tdocumento='" & sDocumento & "'", Cn) > 0 Then
        MsgBox "Pago de documento " & sDocumento & " completado!!!", vbInformation
        Me.Timer_MercadoPago.Enabled = False
        Me.Timer_MercadoPago.Interval = 0
        Me.cmdCancelarMercadopago.Enabled = False
        If lPagoAntesImpresion Then
            lGrabaPagoAntes = True
        End If
        cmdOpcion_Click (99)
    End If
Exit Sub
fin:

End Sub


Private Sub txtQRRecibo_KeyPress(KeyAscii As Integer)
On Error GoTo fin
    If KeyAscii = 13 And lActivaReciboSaldoQr And Trim(txtQRRecibo.Text) <> "" Then
         Dim cadena() As String
         cadena = Split(Trim(txtQRRecibo.Text), "|")
    
        If cadena(3) <> "" Then
            
            If (Calcular("select count(*) as codigo from MINGRESO where trecibo='" & cadena(3) & "'", Cn) = 0) Then
                MsgBox "Recibo #" & cadena(3) & ", No existe!!", vbExclamation, sMensaje
                Me.txtQRRecibo.Text = ""
                Exit Sub
           End If
           If (Calcular("select count(*) as codigo from " & xTabla & " where Codigo='" & cadena(3) & "'", Cn) > 0) Then
                MsgBox "Recibo #" & cadena(3) & ", ya esta siendo usado!!", vbExclamation, sMensaje
                Me.txtQRRecibo.Text = ""
                Exit Sub
           End If
           If cadena(0) <> Format(FechaServidor(), "yyyyMMdd") Then
                MsgBox "Recibo #" & cadena(3) & ", fuera de fecha de uso!" & vbNewLine & "Fecha recibo: " & cadena(0), vbExclamation, sMensaje
                Me.txtQRRecibo.Text = ""
                Exit Sub
           End If
           
           SSTab1.Tab = 1
           Isql = "insert into " & xTabla & " select  tRecibo as Codigo,tRecibo,fFecha,tDescripcion,Moneda,nMonto - (select isnull(sum(monto),0) from DINGRESO where trecibo = '" & cadena(3) & "') nMonto, nMonto as Descripcion from vIngreso where   lAnticipo=1  and  trecibo = '" & cadena(3) & "' "
            Cn.Execute Isql
            RsGrillaRecibido.Requery

            sMonDoc = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & cadena(3) & "'", Cn)
            Dim montorecibo As Double
            
            montorecibo = Calcular("select nmonto as codigo from " & xTabla & " where codigo = '" & cadena(3) & "'", Cn)
            
            If (Round(montorecibo, 2) = 0) Then
                MsgBox "Recibo #" & cadena(3) & ", ya no cuenta con saldo!" & _
                vbNewLine & "Fecha recibo: " & cadena(0) & _
                vbNewLine & "Saldo: " & CStr(montorecibo), vbExclamation, sMensaje
                Me.txtQRRecibo.Text = ""
                Isql = " delete from " & xTabla & " where codigo = '" & cadena(3) & "'"
                Cn.Execute Isql
                
                RsGrillaRecibido.Requery
                
                Exit Sub
            End If
            
            If sMonDoc = "01" Then
                If CDbl(txtSaldo(0).Caption) >= CDbl(montorecibo) Then
                    nAbonoN = nAbonoN + CDbl(montorecibo)
                Else
                    nAbonoN = nAbonoN + CDbl(txtSaldo(0).Caption)
                    Cn.Execute "update " & xTabla & " set nMonto= " & txtSaldo(0).Caption & ", Descripcion='" & txtSaldo(0).Caption & "' where  Codigo='" & cadena(3) & "'"
                    RsGrillaRecibido.Requery
                End If
            
               'nAbonoN = nAbonoN + CDbl(montorecibo)
            Else
                If CDbl(txtSaldo(1).Caption) >= CDbl(montorecibo) Then
                    'nAbonoN = nAbonoN + CDbl(montorecibo)
                    nAbonoE = nAbonoE + CDbl(montorecibo)
                Else
                    nAbonoE = nAbonoE + CDbl(txtSaldo(1).Caption)
                    Cn.Execute "update " & xTabla & " set nMonto= " & txtSaldo(1).Caption & ", Descripcion='" & txtSaldo(1).Caption & "' where  Codigo='" & cadena(3) & "'"
                    RsGrillaRecibido.Requery
                End If
               'nAbonoE = nAbonoE + val(sDescrip)
            End If
             cmdPago(12).Enabled = False
             Me.txtQRRecibo.Text = ""
       Else
          Exit Sub
       End If
       CalculaTotal
    End If
    Exit Sub
fin:
    Dim stringerror As String
    stringerror = Error
    Me.txtQRRecibo.Text = ""
    Screen.MousePointer = vbDefault
    Call Log_Inforest("PAGORECIBOQR", "PAGO DE DOCUMENTO RECIBO QR", "", "", sDocumento, stringerror, "", "FALLA PAGO DE DOCUMENTO RECIBO QR ", sUsuario)
    MsgBox "Error: " & stringerror + " | Proceso: Pagos QR RECIBO", vbCritical, sMensaje
End Sub







