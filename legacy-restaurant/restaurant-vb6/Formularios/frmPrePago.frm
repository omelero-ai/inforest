VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPrePago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmPrePago.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPunto 
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
      Height          =   1215
      Left            =   0
      TabIndex        =   94
      Top             =   3510
      Width           =   3330
      Begin VB.CommandButton cmdPago 
         Caption         =   "Monto"
         Height          =   555
         Index           =   14
         Left            =   135
         TabIndex        =   95
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label txtPunto 
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
         Left            =   1530
         TabIndex        =   97
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label txtCliente 
         BackColor       =   &H00E0E0E0&
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
         Left            =   135
         TabIndex        =   96
         Top             =   810
         Width           =   3090
      End
   End
   Begin VB.Frame fraDocumento 
      Caption         =   " Documento "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   30
      TabIndex        =   83
      Top             =   4725
      Width           =   8175
      Begin VB.CommandButton cmdPago 
         Caption         =   "Otros"
         Height          =   555
         Index           =   13
         Left            =   2790
         TabIndex        =   90
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Nota de Crédito"
         Height          =   555
         Index           =   11
         Left            =   1440
         TabIndex        =   89
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Recibo"
         Height          =   555
         Index           =   12
         Left            =   120
         TabIndex        =   87
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label txtReferencia 
         BackColor       =   &H00E0E0E0&
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
         Left            =   6600
         TabIndex        =   99
         Top             =   585
         Width           =   1485
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Left            =   4860
         TabIndex        =   98
         Top             =   622
         Width           =   1635
      End
      Begin VB.Label txtDocumento 
         BackColor       =   &H00E0E0E0&
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
         Left            =   4230
         TabIndex        =   88
         Top             =   195
         Width           =   2235
      End
      Begin VB.Label txtDocMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   6600
         TabIndex        =   86
         Top             =   195
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   5655
      Left            =   8250
      TabIndex        =   60
      Top             =   0
      Width           =   3555
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Observación :"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   92
         Top             =   4890
         Width           =   990
      End
      Begin VB.Label txtObservacion 
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
         Height          =   465
         Left            =   150
         TabIndex        =   91
         Top             =   5100
         Width           =   3300
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
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
         Left            =   210
         TabIndex        =   85
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTc 
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
         Left            =   1770
         TabIndex        =   84
         Top             =   240
         Width           =   705
      End
      Begin VB.Label txtCortesia 
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
         ForeColor       =   &H00004080&
         Height          =   285
         Left            =   150
         TabIndex        =   82
         Top             =   4530
         Width           =   3300
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
         Index           =   0
         Left            =   150
         TabIndex        =   81
         Top             =   3945
         Width           =   1500
      End
      Begin VB.Label txtVuelto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   80
         Top             =   3360
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
         Index           =   0
         Left            =   150
         TabIndex        =   79
         Top             =   2781
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
         Index           =   0
         Left            =   150
         TabIndex        =   78
         Top             =   2199
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
         Index           =   0
         Left            =   150
         TabIndex        =   77
         Top             =   1617
         Width           =   1500
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
         Index           =   0
         Left            =   150
         TabIndex        =   76
         Top             =   1035
         Width           =   1500
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Cortesia :"
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   75
         Top             =   4320
         Width           =   660
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Recibido :"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   74
         Top             =   3735
         Width           =   1125
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Vuelto :"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   73
         Top             =   3150
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propinas :"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   72
         Top             =   2565
         Width           =   705
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   71
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Abonos :"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   70
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Cargos :"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   69
         Top             =   810
         Width           =   585
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
         Left            =   1950
         TabIndex        =   68
         Top             =   3945
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
         Height          =   315
         Index           =   1
         Left            =   1950
         TabIndex        =   67
         Top             =   3360
         Visible         =   0   'False
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
         Left            =   1950
         TabIndex        =   66
         Top             =   2781
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
         Left            =   1950
         TabIndex        =   65
         Top             =   2199
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
         Left            =   1950
         TabIndex        =   64
         Top             =   1617
         Width           =   1500
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
         Left            =   1950
         TabIndex        =   63
         Top             =   1035
         Width           =   1500
      End
      Begin VB.Label Label 
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
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   62
         Top             =   510
         Width           =   1305
      End
      Begin VB.Label Label 
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
         Height          =   315
         Index           =   1
         Left            =   1950
         TabIndex        =   61
         Top             =   495
         Width           =   1395
      End
   End
   Begin VB.Frame fraOpciones 
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
      Height          =   2520
      Left            =   10260
      TabIndex        =   55
      Top             =   5700
      Width           =   1545
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
         Left            =   90
         Picture         =   "frmPrePago.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1905
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpcion 
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
         Height          =   555
         Index           =   1
         Left            =   90
         Picture         =   "frmPrePago.frx":03FC
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   795
         Width           =   1335
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
         Left            =   90
         Picture         =   "frmPrePago.frx":053E
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1350
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cortesía"
         Height          =   555
         Index           =   3
         Left            =   90
         TabIndex        =   56
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   4725
      Left            =   3390
      TabIndex        =   28
      Top             =   0
      Width           =   4830
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   1
         Left            =   150
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   2
         Left            =   150
         TabIndex        =   35
         Top             =   780
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   3
         Left            =   150
         TabIndex        =   34
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   4
         Left            =   150
         TabIndex        =   33
         Top             =   1860
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   5
         Left            =   150
         TabIndex        =   32
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   6
         Left            =   150
         TabIndex        =   31
         Top             =   2940
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   7
         Left            =   150
         TabIndex        =   30
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   545
         Index           =   8
         Left            =   150
         TabIndex        =   29
         Top             =   4020
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Index           =   9
         Left            =   2025
         TabIndex        =   54
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propina"
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
         Index           =   10
         Left            =   3540
         TabIndex        =   53
         Top             =   180
         Width           =   540
      End
      Begin VB.Label txtTMonto 
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
         Index           =   1
         Left            =   1575
         TabIndex        =   52
         Top             =   470
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
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
         Left            =   1575
         TabIndex        =   51
         Top             =   1010
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
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
         Left            =   1575
         TabIndex        =   50
         Top             =   1550
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
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
         Index           =   4
         Left            =   1575
         TabIndex        =   49
         Top             =   2090
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
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
         Index           =   5
         Left            =   1575
         TabIndex        =   48
         Top             =   2630
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
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
         Index           =   6
         Left            =   1575
         TabIndex        =   47
         Top             =   3170
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
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
         Index           =   7
         Left            =   1575
         TabIndex        =   46
         Top             =   3710
         Width           =   1515
      End
      Begin VB.Label txtTMonto 
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
         Index           =   8
         Left            =   1575
         TabIndex        =   45
         Top             =   4250
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   1
         Left            =   3180
         TabIndex        =   44
         Top             =   470
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   2
         Left            =   3180
         TabIndex        =   43
         Top             =   1010
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   3
         Left            =   3180
         TabIndex        =   42
         Top             =   1550
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   4
         Left            =   3180
         TabIndex        =   41
         Top             =   2090
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   5
         Left            =   3180
         TabIndex        =   40
         Top             =   2630
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   6
         Left            =   3180
         TabIndex        =   39
         Top             =   3170
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   7
         Left            =   3180
         TabIndex        =   38
         Top             =   3710
         Width           =   1515
      End
      Begin VB.Label txtTPropina 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Index           =   8
         Left            =   3180
         TabIndex        =   37
         Top             =   4250
         Width           =   1515
      End
   End
   Begin VB.Frame fraGrilla 
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
      Height          =   2535
      Left            =   30
      TabIndex        =   12
      Top             =   5700
      Width           =   10215
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   8895
         Picture         =   "frmPrePago.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   8895
         Picture         =   "frmPrePago.frx":133A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   8895
         Picture         =   "frmPrePago.frx":1C04
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   8895
         Picture         =   "frmPrePago.frx":24CE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1905
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
         Height          =   555
         Index           =   6
         Left            =   4725
         TabIndex        =   19
         Top             =   1905
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Cancelar"
         Height          =   555
         Index           =   5
         Left            =   6045
         TabIndex        =   18
         Top             =   1905
         Width           =   1245
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
         Height          =   555
         Index           =   4
         Left            =   7365
         TabIndex        =   17
         Top             =   1905
         Width           =   1245
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Propina"
         Height          =   555
         Index           =   10
         Left            =   75
         TabIndex        =   16
         Top             =   1905
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Monto"
         Height          =   555
         Index           =   9
         Left            =   75
         TabIndex        =   15
         Top             =   1350
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Fecha Vencimiento"
         Height          =   555
         Index           =   8
         Left            =   75
         TabIndex        =   14
         Top             =   795
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Número"
         Height          =   555
         Index           =   7
         Left            =   75
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   1680
         Left            =   3210
         TabIndex        =   93
         Top             =   180
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   2963
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
         Caption         =   "Cancelación con Tarjetas"
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
      Begin VB.Label txtNT 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1410
         TabIndex        =   27
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label txtVT 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1410
         TabIndex        =   26
         Top             =   945
         Width           =   1725
      End
      Begin VB.Label txtMT 
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
         Left            =   1560
         TabIndex        =   25
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label txtPT 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   1560
         TabIndex        =   24
         Top             =   2055
         Width           =   1575
      End
   End
   Begin VB.Frame fraCheque 
      Caption         =   " Cheque "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   30
      TabIndex        =   5
      Top             =   1485
      Width           =   3330
      Begin VB.CommandButton cmdPago 
         Caption         =   "Banco"
         Height          =   555
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Número"
         Height          =   555
         Index           =   5
         Left            =   150
         TabIndex        =   7
         Top             =   825
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Monto"
         Height          =   555
         Index           =   6
         Left            =   150
         TabIndex        =   6
         Top             =   1380
         Width           =   1275
      End
      Begin VB.Label txtBCH 
         BackColor       =   &H00E0E0E0&
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
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label txtNCH 
         BackColor       =   &H00E0E0E0&
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
         Top             =   945
         Width           =   1695
      End
      Begin VB.Label txtMCH 
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
         Left            =   1530
         TabIndex        =   9
         Top             =   1500
         Width           =   1695
      End
   End
   Begin VB.Frame fraEfectivo 
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
      Height          =   1485
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   3330
      Begin VB.CommandButton cmdPago 
         Caption         =   "Monto"
         Height          =   555
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Monto"
         Height          =   555
         Index           =   2
         Left            =   150
         TabIndex        =   2
         Top             =   810
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
         Index           =   0
         Left            =   1530
         TabIndex        =   4
         Top             =   360
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
         Index           =   1
         Left            =   1530
         TabIndex        =   3
         Top             =   930
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmPrePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTarjeta As String
Dim RsTarjeta As Recordset
Dim RsGrilla As Recordset
Dim rst1 As Recordset
Dim RsT2 As Recordset
Dim RsPrepago As Recordset

Dim xSinPrepago As String
Dim tCodigoIntegracion As Integer
Dim xTiporecepcion As Integer

Dim nEN As Double
Dim nEE As Double
Dim nCH As Double
Dim nET As Double
Public nPT As Double '25/01/2023 GCAA
Dim nDocumento As Double

Dim nAbonoN As Double
Dim nAbonoE As Double
Dim nTotalN As Double
Dim nTotalE As Double

Dim nVueltoN As Double
Dim nVueltoE As Double
Dim nPropinaN As Double
Dim nPropinaE As Double

Dim nSaldo As Double
Dim mTarjeta(8, 3)
Public sTipoTarjeta As String
Public sTitulo As String
Public nIndex As Integer
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
Dim lDelivery  As Boolean

'mastarjetas
Dim lTarjetaGrilla As Boolean
Dim lPinpadTarjeta As Boolean
Dim xttipopedido As String
 
'mastarjetas


Private Sub cmdNavegar_Click(Index As Integer)
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
   
   End Select
End Sub

Public Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0  ' Salir
               wEnter = False
               Unload Me
               
          Case Is = 1  ' Observacion
               frmKeyBoard.txtResultado.Text = txtObservacion.Caption
               frmKeyBoard.Show vbModal
               If wEnter Then
                  txtObservacion.Caption = sDescrip
               End If
          
          Case Is = 2  ' Grabar
               'Tipo de Cambio
               Dim nCorrelativo As Integer
                  
               If nTC = 0 Then
                  MsgBox "Tipo de Cambio no ingresado", vbCritical, sMensaje
                  Exit Sub
               End If
               
               If nCargo < nEN And (nCargo / nTC) < nEE Then
                  MsgBox "Ambos Efectivos no deben ser mayor al cargo", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               If nSaldo > 0 And sCortesia = "" Then
                  MsgBox "Debe Cancelarse el total del documento", vbExclamation, sMensaje
                  Exit Sub
               End If
                             
               If nCargo < val(Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)) Then
                  MsgBox "El abono(s) de las tarjeta(s) no puede ser mayor al total del documento", vbExclamation, sMensaje
                  Exit Sub
               End If
                             
               Screen.MousePointer = vbHourglass
               wEnter = True
               
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
               
                  Isql = "insert into DPREPAGO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, nVuelto, tObservacion, tUsuario, fRegistro,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & sTurno & "'," _
                                  & "'01'," _
                                  & "'01'," _
                                  & nTC & ", " _
                                  & nEN & ", " _
                                  & nVueltoN & ", " _
                                  & "'" & Mid(txtObservacion.Caption, 1, 100) & "', " _
                                  & "'" & sUsuario & "'," _
                                  & "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "' )"
                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               
               'Efectivo Moneda Extranjera
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
                  
                  Isql = "insert into DPREPAGO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto,  nVuelto, tObservacion, tUsuario, fRegistro,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & sTurno & "'," _
                                  & "'01'," _
                                  & "'02'," _
                                  & nTC & ", " _
                                  & nEE & ", " _
                                  & nVueltoN & ", " _
                                  & "'" & Mid(txtObservacion.Caption, 1, 100) & "', " _
                                  & "'" & sUsuario & "'," _
                                  & "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "')"

                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               
               'Cheque
               If nCH <> 0 Then
                  Isql = "insert into DPREPAGO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto,  nVuelto, tBanco, tNumero, tObservacion, tUsuario, fRegistro,fDiaContable) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & sTurno & "'," _
                                  & "'03'," _
                                  & "'01'," _
                                  & nTC & ", " _
                                  & nCH & ", " _
                                  & nVueltoN & ", " _
                                  & "'" & txtBCH.Caption & "', " _
                                  & "'" & txtNCH.Caption & "', " _
                                  & "'" & Mid(txtObservacion.Caption, 1, 100) & "', " _
                                  & "'" & sUsuario & "'," _
                                  & "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "' )"

                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
               
               'Tarjeta
               If RsGrilla.RecordCount > 0 And sCortesia = "" Then
                  grdGrilla.MoveFirst
                  Dim i As Integer
                  For i = 1 To RsGrilla.RecordCount
                      Isql = "insert into DPREPAGO " & _
                             "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio, nMonto,  nVuelto, nPropina, tTarjeta, tNumero, tFechaVencimiento, tObservacion, tUsuario, fRegistro,fDiaContable ) " & _
                             "Values(    '" & sDocumento & "'," _
                                      & "'" & Trim(str(nCorrelativo)) & "'," _
                                      & "'" & sTurno & "'," _
                                      & "'02'," _
                                      & "'01'," _
                                      & "'" & IIf(IsNull(RsGrilla!tReferencia), "", RsGrilla!tReferencia) & "', " _
                                      & nTC & ", " _
                                      & IIf(IsNull(RsGrilla!nMonto), 0, RsGrilla!nMonto) & ", " _
                                      & nVueltoN & ", " _
                                      & IIf(IsNull(RsGrilla!nPropina), 0, RsGrilla!nPropina) & ", " _
                                      & "'" & RsGrilla!tTipo & "', " _
                                      & "'" & IIf(IsNull(RsGrilla!tNumero), "", RsGrilla!tNumero) & "', " _
                                      & "'" & IIf(IsNull(RsGrilla!tFecha), "", RsGrilla!tFecha) & "', " _
                                      & "'" & Mid(txtObservacion.Caption, 1, 100) & "', " _
                                      & "'" & sUsuario & "'," _
                                      & "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "' )"
                      Cn.Execute Isql
                      nCorrelativo = nCorrelativo + 1
                      grdGrilla.MoveNext
                  Next i
               End If
            
               'Documento
               If nDocumento <> 0 Then
                  Isql = "insert into DPREPAGO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, tUsuario, fRegistro,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & sTurno & "'," _
                                  & "'04'," _
                                  & "'" & sOtroTipoCancelacion & "'," _
                                  & "'" & sMonDoc & "'," _
                                  & nTC & ", " _
                                  & nDocumento & ", " _
                                  & "'" & sTipoDocumento & "', " _
                                  & "'" & IIf(sOtroTipoCancelacion = "000", txtDocumento.Caption, "") & "', " _
                                  & "'" & sUsuario & "'," _
                                  & "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "' )"
                  Cn.Execute Isql
                  If sOtroTipo = "Recibo" Then
                     Cn.Execute "update MINGRESO set tEstadoDocumento ='02' where tRecibo ='" & txtDocumento.Caption & "'"
                  ElseIf sOtroTipo = "NotaCredito" Then
                     Cn.Execute "update MNOTACREDITO set tEstadoDocumento ='02' ,lreplica=1 where tNotaCredito ='" & txtDocumento.Caption & "'"
                  End If
                  nCorrelativo = nCorrelativo + 1
                  
               End If
                                       
               'Puntos
               If nTotalPuntos <> 0 Then
                  Isql = "insert into DPREPAGO " & _
                         "( tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda, nTipoCambio, nMonto, tNumero, tBanco, tUsuario, fRegistro,fDiaContable ) " & _
                         "Values(   '" & sDocumento & "'," _
                                  & "'" & nCorrelativo & "'," _
                                  & "'" & sTurno & "'," _
                                  & "'05'," _
                                  & "''," _
                                  & "'" & sMonDoc & "'," _
                                  & nTC & ", " _
                                  & nTotalPuntos & ", " _
                                  & "'" & sClientePuntos & "', " _
                                  & "'', " _
                                  & "'" & sUsuario & "'," _
                                  & "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "' )"
                                  
                  Cn.Execute Isql
                  nCorrelativo = nCorrelativo + 1
               End If
                                                                         
                                                                         
               '09102020
               
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
                                                                         
               'Liberacion
               If xTipo = "" Then
                  Cn.Execute "Update DPEDIDO set tFacturado ='P' where tDocumento ='" & sDocumento & "'"
                  Set rst1 = Lib.OpenRecordset("select tCodigoPedido from DDOCUMENTO where tDocumento ='" & sDocumento & "'", Cn)
                  If rst1.RecordCount <> 0 Then
                     rst1.MoveFirst
                     For i = 1 To rst1.RecordCount
                         Set RsT2 = Lib.OpenRecordset("select tCodigoPedido from DPEDIDO where tCodigoPedido ='" & rst1!tCodigoPedido & "' and (isnull(tFacturado,'0')='0' or tfacturado='F' or len(ltrim(tFacturado))=0)", Cn)
                         If RsT2.RecordCount = 0 Then
                            Cn.Execute "Update MPEDIDO set tEstadoPedido ='02' where tCodigoPedido ='" & rst1!tCodigoPedido & "'"
                            Cn.Execute "Update TMESA set tEstadoMesa = '04' where tCodigoMesa in (select tMesa from MPEDIDO where tCodigoPedido='" & rst1!tCodigoPedido & "') "
                         End If
                         rst1.MoveNext
                     Next i
                  End If
               End If
               
                xttipopedido = Calcular("Select ttipopedido as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
                xTiporecepcion = Calcular("Select isnull(tTipoRecepcion,0) as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
                If (lPagoDirectoCD) Or (xttipopedido = "01" And xTiporecepcion = 3) Then
                    Dim FDIAC As Date
                    Dim fActual As Date
                    Dim flag As String
                    Dim fFechaPago As Date
                    'aqui se produce el error
                    flag = Calcular("select max(fRegistro) as Codigo from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'", Cn)
                    fActual = Calcular("select getdate() as Codigo    ", Cn)
                    fFechaPago = IIf(flag = "0", Format(fActual, "yyyy/mm/dd hh:nn"), flag)
                    FDIAC = Format(Calcular("select isnull(max(fDiaContable),0) as Codigo from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'", Cn), "yyyy/mm/dd hh:nn")
                 
                     Dim txtFechaDelivery As String
                     txtFechaDelivery = Format("", "##:##")
                 
                   If nTC = 0 Then
                      MsgBox "Tipo de Cambio no ingresado", vbCritical, sMensaje
                      Exit Sub
                   End If
                   
'                   If lFechaEntregaDelivery And txtFechaDelivery = "" Then
'                      MsgBox "La hora de Entrega Delviery no fue Ingresada", vbCritical, sMensaje
'                      Exit Sub
'                   End If
                   
                   If nCargo < nEN And (nCargo / nTC) < nEE Then
                      MsgBox "Ambos Efectivos no deben ser mayor al cargo", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
    '               If nCargo < nAbonoN Then
    '                  MsgBox "El valor del Pago no puede ser mayor al Cargo ", vbExclamation, sMensaje
    '                  Exit Sub
    '               End If
                   
                   If nSaldo > 0 And sCortesia = "" Then
                      MsgBox "Debe Cancelarse el total del documento", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                   'If nCargo < Val(Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)) Then
                   If Round(nCargo, 2) < Round(val(Calcular("select sum(nMonto) as Codigo from " & sTarjeta, Cn)), 2) Then
    
                      MsgBox "El abono(s) de las tarjeta(s) no puede ser mayor al total del documento", vbExclamation, sMensaje
                      Exit Sub
                   End If
                   
                     'JL Correccion
                    If lMCPV Then
                         If MultiCajeroOk = False Then
                             Exit Sub
                         End If
                    End If
                     
                    Screen.MousePointer = vbHourglass
                    wEnter = True
                   
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
                   
                    nCorrelativo = 1
               'Dim nEfectivo As Double
               
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
               
        
               
               'Tarjeta
               If RsGrilla.RecordCount > 0 And sCortesia = "" Then
                  grdGrilla.MoveFirst
                   
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
                                      & "'" & sUsuario & "','" & IIf(xTipo = "Modificacion", Format(FDIAC, "yyyyMMdd"), Format(obtieneDiaContable, "yyyyMMdd")) & "', '" & IIf(IsNull(RsGrilla!tNumero), "", RsGrilla!tNumero) & "')"
                  
                      Cn.Execute Isql
                      nCorrelativo = nCorrelativo + 1
                      grdGrilla.MoveNext
                  Next i
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
                  Cn.Execute "UPDATE dbo.MPEDIDO set fLlegada=getdate(), fEntrega = {fn CURDATE()} + ' " & txtFechaDelivery & "' where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' ) and tTipoPedido='02' and isnull(fLlegada ,0)=0"
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
               
               
            xTiporecepcion = Calcular("Select Distinct isnull( P.tTiporecepcion,0) as codigo from DDocumento as D Inner Join MPedido AS P On D.tCodigoPedido = P.tCodigoPedido Where tDocumento ='" & sDocumento & "'", Cn)
                    
               If xTiporecepcion <> 0 Then
  
                    If xTiporecepcion = 1 Then
                        Call ModifcarEstadoDeliveryCabecera(sDocumento)
                    ElseIf xTiporecepcion = 2 Then
                        'Call ModificarEstadoPedidoWebAp(sDocumento)
                    End If
                    
               End If
                              
'               Screen.MousePointer = vbDefault
'               wEnter = True
'               If lPagoAntesImpresion Then
'                    lGrabaPagoAntes = True
'               End If
'
               
               '16112020 GCAA
                   
                   
                End If
                              
                              
               Screen.MousePointer = vbDefault
               Unload Me
               
          Case Is = 3  ' Cortesia
               If MsgBox("Seguro de Cancelarlo por Cortesia?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
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
                  sCortesia = sCodigo
                  txtCortesia.Caption = sDescrip
               Else
                  sCortesia = ""
                  txtCortesia.Caption = ""
               End If
               
                Limpia
                txtNT.Caption = ""
                txtVT.Caption = ""
                txtMT.Caption = ""
                txtPT.Caption = ""
                Cn.Execute "delete from " & sTarjeta
                RsGrilla.Requery
          
          Case Is = 4  ' Agregar Tarjeta
               If nET = 0 Then
                  MsgBox "El monto no puede ser cero", vbExclamation, sMensaje
                  Exit Sub
               End If
                              
               'Oscar Ortega
               nAbonoN = nAbonoN + nET
               nPropinaN = nPropinaN + nPT
               CalculaTotal
               
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
                               & "''," _
                               & nET & ", " _
                               & nPT & ", " _
                               & nIndex & ",0)"
               Cn.Execute Isql
               RsGrilla.Requery
               ActivaTarjeta False
               
               txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption + nET, "###,###,##0.00")
               txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT, "###,###,##0.00")
               AsignarTotal
               txtNT.Caption = ""
               txtVT.Caption = ""
               txtMT.Caption = "0.00"
               txtPT.Caption = "0.00"
               Bloquear False
                         
          Case Is = 5  ' Cancelar Tarjeta
               txtNT.Caption = ""
               txtVT.Caption = ""
               txtMT.Caption = "0.00"
               txtPT.Caption = "0.00"
               sRefer = ""
               ActivaTarjeta False
               Bloquear False
          
          Case Is = 6  ' Eliminar Tarjeta
               If RsGrilla.RecordCount <> 0 Then
                  If Not RsGrilla!lPinPad Then
                     nIndex = RsGrilla!nIndex
                     txtTMonto(nIndex).Caption = Format(txtTMonto(nIndex).Caption - RsGrilla!nMonto, "###,###,##0.00")
                     txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption - RsGrilla!nPropina, "###,###,##0.00")
                     nAbonoN = nAbonoN - RsGrilla!nMonto
                     nPropinaN = nPropinaN - RsGrilla!nPropina
                     Cn.Execute "delete from " & sTarjeta & " where tIndex = '" & RsGrilla!tIndex & "'"
                     CalculaTotal
                     RsGrilla.Requery
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

          
   End Select
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
Private Sub cmdPago_Click(Index As Integer)
   Select Case Index
          Case Is = 0  ' Efectivo Monto MN
               tCodigoIntegracion = 0
               nAbonoN = nAbonoN - nEN
               sTipo = ""
               frmNumPad.Show vbModal
               nEN = IIf(wEnter = True, sDescrip, nEN)
               txtME(0).Caption = Format(nEN, "###,###,###,##0.00")
               nAbonoN = nAbonoN + nEN
               CalculaTotal
                    
          Case Is = 2  ' Efectivo Monto ME
               nAbonoE = nAbonoE - nEE
               sTipo = ""
               frmNumPad.Show vbModal
               nEE = IIf(wEnter = True, sDescrip, nEE)
               txtME(1).Caption = Format(nEE, "###,###,###,##0.00")
               nAbonoE = nAbonoE + nEE
               CalculaTotal
                              
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
               sTipo = ""
               frmNumPad.Show vbModal
               nCH = IIf(wEnter = True, sDescrip, nCH)
               txtMCH.Caption = Format(nCH, "###,###,###,##0.00")
               nAbonoN = nAbonoN + nCH
               CalculaTotal
                         
          Case Is = 7  ' Numero Tarjeta
               sTipo = "Numero"
               frmNumPad.Show vbModal
               txtNT.Caption = IIf(wEnter = True, Mid(sDescrip, 1, 16), txtNT.Caption)
          
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
               sTipoDocumento = txtDocumento.Caption
               CalculaTotal
          
          Case Is = 12  ' Recibos
               sTemp = ""
               Isql = "select *, tRecibo as Codigo, nMonto as Descripcion from vIngreso where tEstadodocumento ='01' and lAnticipo=1"
               Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Recibo", 2, "tRecibo", 1500, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1250, 0, 0, "dd/mm/yyyy", _
                                                               "Motivo", 2, "tDescripcion", 3500, 0, 0, "", _
                                                               "Mon", 2, "Moneda", 700, 2, 0, "", _
                                                               "Monto", 2, "nMonto", 1300, 1, 0, "##,##0.00")
               frmBusquedaRapida.nPredeterm = 0
               frmBusquedaRapida.Show vbModal
               
               If wEnter Then
                  txtDocumento.Caption = "Recibo/Anticipo"
                  sTipoDocumento = sCodigo
                  txtReferencia.Caption = sCodigo
                  sOtroTipoCancelacion = "001"
                  sOtroTipo = sTipo
                  
                  If sMonDoc = "01" Then
                     nAbonoN = nAbonoN - nDocumento
                  Else
                     nAbonoE = nAbonoE - nDocumento
                  End If
                  
                  sMonDoc = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & sCodigo & "'", Cn)
                  
                  If sMonDoc = "01" Then
                     nAbonoN = nAbonoN + val(sDescrip)
                  Else
                     nAbonoE = nAbonoE + val(sDescrip)
                  End If
                  nDocumento = val(sDescrip)
               Else
                  Exit Sub
               End If
                              
               txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
               sOtroTipoCancelacion = "001"
               sOtroTipo = sTipo
               sTipoDocumento = txtDocumento.Caption
               CalculaTotal
          
          Case Is = 13  ' Varios
               sTemp = ""
               Isql = "select Codigo, Descripcion from vTipoCancelacion Where lActivo='1'"
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
                  nDisponible = nVar1
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
   End Select
End Sub

Public Sub cmdTarjeta_Click(Index As Integer)

     lTarjetaGrilla = False
     lPinpadTarjeta = False

     If Index = "8" And cmdTarjeta(8).Caption = "OTRAS" Then
        
     
               sTemp = ""
               Isql = "select tcodigotarjeta codigo, tdetallado,tresumido as descripcion,isnull(nfactorretencion,0) nfactorretencion, isnull(lpinpad,0) lpinpad, case when isnull(nboton,0)=0 then 8 else nboton end nboton, lactivo From TTARJETACREDITO Where lactivo = 1 order by tdetallado"
               Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                               "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               If wEnter Then
                'gcaa 30072020
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
                   
                   Else
                      ActivaTarjeta True
                      txtNT.Caption = ""
                      txtVT.Caption = ""
                      txtMT.Caption = "0.00"
                      txtPT.Caption = "0.00"
                  End If
     
     
             End If
     
     
     
               
     Else
     'GCAA 300707
           '' If Index > 7 Then
           ''     Index = 8
           ''  End If
                   sTipoTarjeta = mTarjeta(Index, 1)
                   sTitulo = mTarjeta(Index, 2)
                   nIndex = Index
                   nET = 0
                   nPT = 0
                   
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
                   
                   Else
                      ActivaTarjeta True
                      txtNT.Caption = ""
                      txtVT.Caption = ""
                      txtMT.Caption = "0.00"
                      txtPT.Caption = "0.00"
                  End If
     
     
     
     End If
   
End Sub

Public Sub Form_Load()
   wEnter = False
   cmdPago(0).Caption = "Monto en " & sMonN
   fraPunto.Caption = " " & sClub & " "
   lblTc = Format(nTC, "##0.00")
   lDelivery = True
   lPagocortesiaAI = False
   If lCancelacion Then
      cmdOpcion(0).Enabled = False
   End If
   
'    If lDesPagoCheque Then
'        Me.fraCheque.Enabled = False
'    Else
'        Me.fraCheque.Enabled = True
'    End If
'
'    If lDesPagoOtra Then
'        Me.fraDocumento.Enabled = False
'    Else
'        Me.fraDocumento.Enabled = True
'    End If
    If lDesactivaNCFP Then
        cmdPago(11).Enabled = False
    Else
        cmdPago(11).Enabled = True
    End If
   
   
   If sMonE <> "" And sMonN <> sMonE Then
      cmdPago(2).Caption = "Monto en " & sMonE
   Else
      cmdPago(2).Visible = False
    '  cmdPago(3).Visible = False
   End If
   frmPrePago.Caption = "Cancelación del Documento " & Format(sDocumento, "@-@@@@@-@@@@@@@@@")
   
   'Tarjetas de Crédito
   sTarjeta = dbTemporal(sCaja, 10, "tIndex", "nVarChar(3)", _
                                   "tTipo", "nVarChar(2)", _
                                   "tDescripcion", "nVarChar(30)", _
                                   "tNumero", "nVarChar(16)", _
                                   "tFecha", "nVarChar(4)", _
                                   "nMonto", "Float", _
                                   "nPropina", "Float", _
                                   "nIndex", "Int", _
                                   "lPinPad", "bit", _
                                   "tReferencia", "nVarChar(4)")
                                        
   Set RsGrilla = Lib.OpenRecordset("select * from " & sTarjeta, Cn)
   Call ConfGrilla(6, grdGrilla, "Tarjeta", 2, "tDescripcion", 1000, 0, 0, "", _
                               "Numero", 2, "tNumero", 800, 0, 0, "", _
                               "Fec.Venc", 2, "tFecha", 900, 1, 0, "", _
                               "Monto", 2, "nMonto", 950, 1, 0, "####,###,##0.00", _
                               "Propina", 2, "nPropina", 850, 1, 0, "####,###,##0.00", _
                               "Ref.", 2, "tReferencia", 500, 0, 0, "")
   Set grdGrilla.DataSource = RsGrilla
   sMonDoc = "01"
   
   
    'Tarjetas de Crédito
    If Calcular("select count(*) as codigo from ttarjetacredito where lactivo=1", Cn) > 8 Then
        Isql = "select tcodigotarjeta, tdetallado,tresumido,isnull(nfactorretencion,0) nfactorretencion,case when isnull(nboton,0)=0 then 8 else nboton end nboton, lactivo, isnull(lpinpad,0) lpinpad from TTARJETACREDITO where lactivo=1"
        Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
        'mastarjetas
        If RsTarjeta.RecordCount <= 8 Then
             Call AsignaTarjeta(8, RsTarjeta, cmdTarjeta())
        Else
            Call AsignaTarjeta(7, RsTarjeta, cmdTarjeta())
            cmdTarjeta(8).Caption = "OTRAS"
            mTarjeta(8, 1) = 99
            mTarjeta(8, 2) = "OTRAS"
            mTarjeta(8, 3) = "0"
            cmdTarjeta(8).Visible = True
        End If

    Else
   
        Isql = "select * from TTARJETACREDITO where nBoton>0 and lActivo = 1 Order by nBoton"
        Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
        'mastarjetas
     
         Call AsignaTarjeta(8, RsTarjeta, cmdTarjeta())
        
    End If
    'mas tarjetas
      
      
      
   ActivaTarjeta False
   Limpia
   'lDelivery = False
     
   Set RsPrepago = Lib.OpenRecordset("select * from DPREPAGO where tdocumento='" & sDocumento & "'", Cn)
   If RsPrepago.RecordCount <> 0 Then
      
      RsPrepago.MoveFirst
      txtObservacion.Caption = IIf(IsNull(RsPrepago!tObservacion), "", RsPrepago!tObservacion)
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
                     nCH = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                     txtMCH.Caption = Format(nCH, "###,###,###,##0.00")
                     nAbonoN = nAbonoN + nCH
                     txtNCH.Caption = IIf(IsNull(RsPrepago!tNumero), "", RsPrepago!tNumero)
                     txtBCH.Caption = IIf(IsNull(RsPrepago!tBanco), "", RsPrepago!tBanco)
                     
                Case Is = "04" 'Documento
                     nDocumento = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                     txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
                     nAbonoN = nAbonoN + nDocumento
                     sOtroTipoCancelacion = RsPrepago!tOtroTipoPago
                     txtDocumento.Caption = Calcular("select Cliente as codigo from vDELIVERY where Codigo='" & Trim(IIf(IsNull(RsPrepago!tNumero), "", RsPrepago!tNumero)) & "'", Cn)
                                   
                Case Is = "05" 'Puntos
                     nTotalPuntos = IIf(IsNull(RsPrepago!nMonto), 0, RsPrepago!nMonto)
                     txtPunto.Caption = Format(nTotalPuntos, "###,###,###,##0.00")
                     nAbonoN = nAbonoN + nTotalPuntos
                     txtCliente.Caption = Calcular("select Cliente as codigo from vDELIVERY where Codigo='" & Trim(IIf(IsNull(RsPrepago!tNumero), "", RsPrepago!tNumero)) & "'", Cn)
         End Select
         CalculaTotal
         RsPrepago.MoveNext
      Loop
   Else
      Dim RsDelivery As Recordset
      lDelivery = False
      Dim xCliente As String
      xCliente = Calcular("SELECT dbo.MPEDIDO.tClienteDelivery AS Codigo FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido WHERE (dbo.DPEDIDO.tDocumento = '" & sDocumento & "')", Cn)
      
      Isql = "select tCodigoTarjeta, tNumeroTarjeta, tFechaTarjeta from TDELIVERY where tCodigoDelivery ='" & xCliente & "'"
      Set RsDelivery = Lib.OpenRecordset(Isql, Cn)
        
      If RsDelivery.RecordCount > 0 Then
         If IIf(IsNull(RsDelivery!tCodigoTarjeta), "", RsDelivery!tCodigoTarjeta) <> "" Then
            sTipoTarjeta = IIf(IsNull(RsDelivery!tCodigoTarjeta), "", RsDelivery!tCodigoTarjeta)
            txtNT.Caption = IIf(IsNull(RsDelivery!tNumeroTarjeta), "", RsDelivery!tNumeroTarjeta)
            txtVT.Caption = IIf(IsNull(RsDelivery!tFechaTarjeta), "", RsDelivery!tFechaTarjeta)
            nET = nCargo
            txtMT.Caption = Format(nET, "###,###,###,##0.00")
            
            RsTarjeta.MoveFirst
            RsTarjeta.Find ("tCodigoTarjeta='" & sTipoTarjeta & "'")
            If Not RsTarjeta.EOF Then
               sTitulo = IIf(IsNull(RsTarjeta!tDetallado), "", RsTarjeta!tDetallado)
               nIndex = IIf(IsNull(RsTarjeta!nBoton), "", RsTarjeta!nBoton)
               fraGrilla.Caption = " Tarjeta de Credito " & RTrim(sTitulo) & " "
            Else
               sTitulo = "(No Definido)"
               nIndex = 1
            End If
            ActivaTarjeta True
         End If
      End If
   End If
   
   If lVisaNet Then
'      If Not ValidaIP(IpPinPad, IpPort) Then
'         MsgBox "Error de conexión", vbCritical, "VisaNet"
         lVisaNet = False
'      End If
   End If
   
   If pais = "003" Then
    fraCheque.Visible = False
   End If

  'GCAA 16112020
   'xSinPrepago = Calcular("select isnull(lPagoDirectoCD,0)AS codigo from tcaja Where tcaja='" & sCaja & "'", Cn)
   Bloquear False
   
End Sub

Public Sub Limpia()
 nEN = 0
 nEE = 0
 nCH = 0
 nET = 0
 nPT = 0
 nDocumento = 0
 nTotalPuntos = 0
 nAbonoN = 0
 nAbonoE = 0
 nSaldo = nCargo - nAbonoN
 nVueltoN = 0
 nVueltoE = 0
 nPropinaN = 0
 nTotalN = 0
 nVueltoE = 0
 nPropinaE = 0
 nTotalE = 0
    
 txtCargo(0).Caption = Format(nCargo, "###,###,##0.00")
 txtAbono(0).Caption = Format(nAbonoN, "###,###,##0.00")
 txtSaldo(0).Caption = Format(nSaldo, "###,###,##0.00")
 txtVuelto(0).Caption = Format(nVueltoN, "###,###,##0.00")
 txtTotal(0).Caption = Format(nTotalN, "###,###,##0.00")
 txtCliente.Caption = ""
 txtDocumento.Caption = ""
 
 If sMonE <> "" And sMonN <> sMonE Then
    txtCargo(1).Caption = Format(nCargo / nTC, "###,###,##0.00")
    txtAbono(1).Caption = Format(nAbonoN / nTC, "###,###,##0.00")
    txtSaldo(1).Caption = Format(nSaldo / nTC, "###,###,##0.00")
    txtVuelto(1).Caption = Format(nVueltoN / nTC, "###,###,##0.00")
    txtTotal(1).Caption = Format(nTotalN / nTC, "###,###,##0.00")
 Else
    txtCargo(1).Visible = False
    txtAbono(1).Visible = False
    txtSaldo(1).Visible = False
    txtVuelto(1).Visible = False
    txtTotal(1).Visible = False
 End If
 
 txtME(0).Caption = "0.00"
 txtME(1).Caption = "0.00"
 txtMT.Caption = "0.00"
 txtPT.Caption = "0.00"
 
 txtMCH.Caption = "0.00"
 txtBCH.Caption = ""
 txtNCH.Caption = ""
 Dim i As Integer
 For i = 1 To 8
     txtTMonto(i).Caption = "0.00"
     txtTPropina(i).Caption = "0.00"
 Next i
End Sub

Public Sub CalculaTotal()
   txtCargo(0).Caption = Format(nCargo, "###,###,##0.00")
   txtCargo(1).Caption = Format(nCargo / nTC, "###,###,##0.00")
      
   txtAbono(0).Caption = Format(nAbonoN, "###,###,##0.00")
   txtAbono(1).Caption = Format(nAbonoE, "###,###,##0.00")
      
   nSaldo = nCargo - nAbonoN
   If nAbonoE > 0 Then
      nSaldo = nSaldo - (nAbonoE * nTC)
   End If
   nSaldo = Round(IIf(nSaldo >= 0, nSaldo, 0), 2)
           
   txtSaldo(0).Caption = Format(nSaldo, "###,###,##0.00")
   txtSaldo(1).Caption = Format(nSaldo / nTC, "###,###,##0.00")
   
   nTotalN = nAbonoN + nPropinaN + (nAbonoE * nTC) + (nPropinaE * nTC)
   nTotalE = (nAbonoN / nTC) + (nPropinaN / nTC) + nAbonoE + nPropinaE
   
   txtTotal(0).Caption = Format(nTotalN, "###,###,##0.00")
   txtTotal(1).Caption = Format(nTotalE, "###,###,##0.00")
   
   txtPropina(0).Caption = Format(nPropinaN, "###,###,##0.00")
   txtPropina(1).Caption = Format(nPropinaE, "###,###,##0.00")
   
   nVueltoN = IIf((nAbonoN + (nAbonoE * nTC)) - nCargo > 0, (nAbonoN + (nAbonoE * nTC)) - nCargo, 0)
   nVueltoE = IIf((nAbonoE + (nAbonoN / nTC)) - nCargo / nTC > 0, (nAbonoE + (nAbonoN / nTC)) - nCargo / nTC, 0)
   txtVuelto(0).Caption = Format(nVueltoN, "###,###,##0.00")
   txtVuelto(1).Caption = Format(nVueltoE, "###,###,##0.00")
End Sub

Public Sub ActivaTarjeta(Activa As Boolean)
   Dim i As Integer
   cmdPago(7).Enabled = Activa    'Numero
   cmdPago(8).Enabled = Activa    'Vencimiento
   cmdPago(9).Enabled = Activa    'Monto
   cmdPago(10).Enabled = Activa   'Propina
   
   cmdOpcion(4).Enabled = Activa  'Agregar
   cmdOpcion(5).Enabled = Activa  'Modificar
   
   For i = 1 To 8
       cmdTarjeta(i).Enabled = Not Activa
   Next i
End Sub

Public Sub AsignarTotal()
   If lDelivery Then
      If RsGrilla.RecordCount <> 0 Then
         txtMT.Caption = Format(RsGrilla!nMonto, "###,###,##0.00")
         txtVT.Caption = IIf(IsNull(RsGrilla!tFecha), "", RsGrilla!tFecha)
         sRefer = RsGrilla!tReferencia
         txtPT.Caption = Format(RsGrilla!nPropina, "###,###,##0.00")
         txtNT.Caption = RsGrilla!tNumero
         cmdPago(10).Enabled = IIf(sRefer <> "", True, False)
         nIndex = RsGrilla!nIndex
      Else
         txtMT.Caption = ""
         txtVT.Caption = ""
         txtPT.Caption = ""
         txtNT.Caption = ""
         cmdPago(10).Enabled = False
         sRefer = ""
         nIndex = 1
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cn.Execute "drop table " & sTarjeta
   Set RsTarjeta = Nothing
   Set RsGrilla = Nothing
   Set rst1 = Nothing
   Set RsT2 = Nothing
   Set frmPrePago = Nothing
End Sub

Private Sub grdGrilla_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   AsignarTotal
End Sub

Public Sub AsignaTarjeta(nBoton As Integer, RsAsigna As Recordset, cmdBoton As Object)
   Screen.MousePointer = vbHourglass
   Dim i As Integer
   With RsAsigna
        If .RecordCount > 0 Then
           For i = 1 To nBoton
               .MoveFirst
               .Find "nboton = " & Trim(str(i))
               If Not .EOF Then
                  cmdBoton(i).Visible = True
                  cmdBoton(i).Caption = !tResumido
                  mTarjeta(i, 1) = !tCodigoTarjeta
                  mTarjeta(i, 2) = !tResumido
                  mTarjeta(i, 3) = IIf(IsNull(!lPinPad), 0, !lPinPad)
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
End Sub

Public Property Let PnET(valuew As String) ' Monto Tarjeta
    sTipo = ""
    nET = valuew
    txtMT.Caption = Format(nET, "###,###,###,##0.00")
    'nAbonoN = nAbonoN + nET
End Property

Public Property Let PnPT(valuew As String) ' Propina Tarjeta
    sTipo = ""
    nPT = valuew
    txtPT.Caption = Format(nPT, "###,###,###,##0.00")
    If sRefer <> "" Then
       nPropinaN = nPropinaN + nPT
       CalculaTotal
       txtTPropina(nIndex).Caption = Format(txtTPropina(nIndex).Caption + nPT - RsGrilla!nPropina, "###,###,##0.00")
       Cn.Execute "update " & sTarjeta & " set nPropina=" & nPT & " where tReferencia='" & sRefer & "'"
       RsGrilla.Requery
    End If
End Property

Public Property Let PnEN(valuew As String) ' Efectivo Monto MN
    'nAbonoN = nAbonoN - nEN
    sTipo = ""
    nEN = valuew
    txtME(0).Caption = Format(nEN, "###,###,###,##0.00")
    nAbonoN = nAbonoN + nEN
    CalculaTotal
End Property

Public Property Let PnEE(valuew As String) ' Efectivo Monto ME
    'nAbonoE = nAbonoE - nEE
    sTipo = ""
    nEE = valuew
    txtME(1).Caption = Format(nEE, "###,###,###,##0.00")
    nAbonoE = nAbonoE + nEE
    CalculaTotal
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function CargarDocumento(ByVal otrotipocancelacion As String, ByVal strDocumento As String, ByVal dblDocumento As String, ByVal gTipoDocumento As String) 'pp
    sOtroTipoCancelacion = otrotipocancelacion
    txtDocumento.Caption = strDocumento
    nAbonoN = nAbonoN - nDocumento
    sTipo = ""
    nDocumento = dblDocumento
    txtDocMonto.Caption = dblDocumento
    nAbonoN = nAbonoN + nDocumento
    CalculaTotal
    sTipoDocumento = gTipoDocumento
    txtReferencia.Caption = sTipoDocumento
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

Public Sub Bloquear(Estado As Boolean)
   Dim RTemps As Recordset
   Dim i As Integer
   Dim xttipopedido As String
   xttipopedido = Calcular("Select ttipopedido as codigo from MPEDIDO Where tCodigoPedido='" & sPedido & "'", Cn)
   
   Isql = "SELECT lVisibilidad,lCanal1,lCanal2,lCanal3,lCanal4,lCanal5, " & _
           "TTC.tCodigoTarjeta,TTC.tDetallado,TTC.nBoton " & _
           "FROM TVISIBILIDADTARJETACREDITOXCANAL TVT " & _
           "INNER JOIN TTARJETACREDITO TTC ON TVT.tCodigoTarjeta = TTC.tcodigotarjeta"
        Set RTemps = Lib.OpenRecordset(Isql, Cn)
   
   
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
                        End If
                        
                    Case "02"
                        If RTemps!lCanal2 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                        End If
                    
                    Case "03"
                        If RTemps!lCanal3 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                        End If
                    
                    Case "04"
                        If RTemps!lCanal4 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                        End If
                    
                    Case "05"
                        If RTemps!lCanal5 Then
                            cmdTarjeta(i).Enabled = Not Estado
                        Else
                            cmdTarjeta(i).Enabled = Estado
                        End If
            End Select
        End If
   Next i
   
   cmdPago(0).Enabled = Not Estado
   cmdPago(2).Enabled = Not Estado
   cmdPago(4).Enabled = Not Estado
   cmdPago(5).Enabled = Not Estado
   cmdPago(6).Enabled = Not Estado
   'cmdPago(3).Enabled = Not Estado
   cmdPago(11).Enabled = Not Estado
   cmdPago(12).Enabled = Not Estado
   cmdPago(13).Enabled = Not Estado
   cmdPago(14).Enabled = Not Estado
End Sub


