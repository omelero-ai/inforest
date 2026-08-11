VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLiquidacionMozo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Turno       "
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frmLiquidacionMozo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Observacion"
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
      Index           =   7
      Left            =   9225
      Picture         =   "frmLiquidacionMozo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4860
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Paloteo"
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
      Index           =   5
      Left            =   9225
      Picture         =   "frmLiquidacionMozo.frx":044C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3780
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Liquidación"
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
      Index           =   4
      Left            =   9225
      Picture         =   "frmLiquidacionMozo.frx":097E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3105
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Comprobante"
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
      Left            =   9225
      Picture         =   "frmLiquidacionMozo.frx":0EB0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5535
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cierre"
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
      Index           =   0
      Left            =   9225
      Picture         =   "frmLiquidacionMozo.frx":13E2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6210
      Width           =   1275
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
      Height          =   600
      Index           =   1
      Left            =   9225
      Picture         =   "frmLiquidacionMozo.frx":14E4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7290
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7530
      Left            =   45
      TabIndex        =   10
      Top             =   405
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   13282
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Resumen"
      TabPicture(0)   =   "frmLiquidacionMozo.frx":15D6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(32)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtTotalIngresoN"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(23)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPuntoN"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtInicioN"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtEfectivoN"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtChequeN"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtMN"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTotalTarjeta"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtTotalOtroN"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPuntoE"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label(31)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label(27)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTotalOtroE"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtME"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtEfectivoE"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtInicioE"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTotalTarjetaPropina"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtChequeE"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label(10)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtTotalIngresoE"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblReciboN"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblReciboE"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdMovimiento(8)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdMovimiento(7)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdMovimiento(6)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdMovimiento(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdMovimiento(4)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdMovimiento(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdMovimiento(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Tarjetas de Crédito"
      TabPicture(1)   =   "frmLiquidacionMozo.frx":15F2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPropina(8)"
      Tab(1).Control(1)=   "cmdTarjeta(8)"
      Tab(1).Control(2)=   "cmdPropina(6)"
      Tab(1).Control(3)=   "cmdTarjeta(6)"
      Tab(1).Control(4)=   "cmdPropina(5)"
      Tab(1).Control(5)=   "cmdTarjeta(5)"
      Tab(1).Control(6)=   "cmdPropina(4)"
      Tab(1).Control(7)=   "cmdTarjeta(4)"
      Tab(1).Control(8)=   "cmdPropina(3)"
      Tab(1).Control(9)=   "cmdTarjeta(3)"
      Tab(1).Control(10)=   "cmdPropina(2)"
      Tab(1).Control(11)=   "cmdTarjeta(2)"
      Tab(1).Control(12)=   "cmdPropina(1)"
      Tab(1).Control(13)=   "cmdTarjeta(1)"
      Tab(1).Control(14)=   "cmdPropina(7)"
      Tab(1).Control(15)=   "cmdTarjeta(7)"
      Tab(1).Control(16)=   "Label(14)"
      Tab(1).Control(17)=   "txtTarjeta(1)"
      Tab(1).Control(18)=   "txtTarjeta(2)"
      Tab(1).Control(19)=   "txtTarjeta(3)"
      Tab(1).Control(20)=   "txtTarjeta(4)"
      Tab(1).Control(21)=   "txtTarjeta(5)"
      Tab(1).Control(22)=   "txtTarjeta(6)"
      Tab(1).Control(23)=   "txtTarjeta(8)"
      Tab(1).Control(24)=   "txtPropina(1)"
      Tab(1).Control(25)=   "txtPropina(2)"
      Tab(1).Control(26)=   "txtPropina(3)"
      Tab(1).Control(27)=   "txtPropina(4)"
      Tab(1).Control(28)=   "txtPropina(5)"
      Tab(1).Control(29)=   "txtPropina(6)"
      Tab(1).Control(30)=   "txtPropina(8)"
      Tab(1).Control(31)=   "lblTarjeta(1)"
      Tab(1).Control(32)=   "lblTarjeta(2)"
      Tab(1).Control(33)=   "lblTarjeta(6)"
      Tab(1).Control(34)=   "lblTarjeta(3)"
      Tab(1).Control(35)=   "lblTarjeta(5)"
      Tab(1).Control(36)=   "lblTarjeta(8)"
      Tab(1).Control(37)=   "lblTarjeta(4)"
      Tab(1).Control(38)=   "Label(20)"
      Tab(1).Control(39)=   "txtTarjeta(7)"
      Tab(1).Control(40)=   "txtPropina(7)"
      Tab(1).Control(41)=   "lblTarjeta(7)"
      Tab(1).Control(42)=   "txtTTarjeta"
      Tab(1).Control(43)=   "txtTPropina"
      Tab(1).Control(44)=   "lblTarjeta(0)"
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "Otros Tipos de Pago"
      TabPicture(2)   =   "frmLiquidacionMozo.frx":160E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdOtroE(10)"
      Tab(2).Control(1)=   "cmdOtroN(10)"
      Tab(2).Control(2)=   "cmdOtroE(9)"
      Tab(2).Control(3)=   "cmdOtroN(9)"
      Tab(2).Control(4)=   "cmdOtroN(7)"
      Tab(2).Control(5)=   "cmdOtroN(1)"
      Tab(2).Control(6)=   "cmdOtroN(3)"
      Tab(2).Control(7)=   "cmdOtroN(4)"
      Tab(2).Control(8)=   "cmdOtroN(5)"
      Tab(2).Control(9)=   "cmdOtroN(6)"
      Tab(2).Control(10)=   "cmdOtroN(8)"
      Tab(2).Control(11)=   "cmdOtroE(8)"
      Tab(2).Control(12)=   "cmdOtroE(6)"
      Tab(2).Control(13)=   "cmdOtroE(5)"
      Tab(2).Control(14)=   "cmdOtroE(4)"
      Tab(2).Control(15)=   "cmdOtroE(3)"
      Tab(2).Control(16)=   "cmdOtroE(1)"
      Tab(2).Control(17)=   "cmdOtroE(7)"
      Tab(2).Control(18)=   "cmdOtroE(2)"
      Tab(2).Control(19)=   "cmdOtroN(2)"
      Tab(2).Control(20)=   "lblOtro(9)"
      Tab(2).Control(21)=   "lblOtro(10)"
      Tab(2).Control(22)=   "txtOtroE(10)"
      Tab(2).Control(23)=   "txtOtroN(10)"
      Tab(2).Control(24)=   "txtOtroE(9)"
      Tab(2).Control(25)=   "txtOtroN(9)"
      Tab(2).Control(26)=   "lblOtro(11)"
      Tab(2).Control(27)=   "txtOtroN(11)"
      Tab(2).Control(28)=   "lblOtro(7)"
      Tab(2).Control(29)=   "txtOtroN(7)"
      Tab(2).Control(30)=   "lblOtro(4)"
      Tab(2).Control(31)=   "lblOtro(8)"
      Tab(2).Control(32)=   "lblOtro(5)"
      Tab(2).Control(33)=   "lblOtro(3)"
      Tab(2).Control(34)=   "lblOtro(6)"
      Tab(2).Control(35)=   "lblOtro(1)"
      Tab(2).Control(36)=   "txtOtroN(8)"
      Tab(2).Control(37)=   "txtOtroN(6)"
      Tab(2).Control(38)=   "txtOtroN(5)"
      Tab(2).Control(39)=   "txtOtroN(4)"
      Tab(2).Control(40)=   "txtOtroN(3)"
      Tab(2).Control(41)=   "txtOtroN(1)"
      Tab(2).Control(42)=   "lblMontoN"
      Tab(2).Control(43)=   "txtOtroE(1)"
      Tab(2).Control(44)=   "txtOtroE(3)"
      Tab(2).Control(45)=   "txtOtroE(4)"
      Tab(2).Control(46)=   "txtOtroE(5)"
      Tab(2).Control(47)=   "txtOtroE(6)"
      Tab(2).Control(48)=   "txtOtroE(8)"
      Tab(2).Control(49)=   "txtOtroE(7)"
      Tab(2).Control(50)=   "txtOtroE(11)"
      Tab(2).Control(51)=   "txtOtroE(2)"
      Tab(2).Control(52)=   "txtOtroN(2)"
      Tab(2).Control(53)=   "lblOtro(2)"
      Tab(2).Control(54)=   "lblMontoE"
      Tab(2).ControlCount=   55
      TabCaption(3)   =   "Observación"
      TabPicture(3)   =   "frmLiquidacionMozo.frx":162A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtObservacion"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   10
         Left            =   -67485
         TabIndex        =   139
         Top             =   6405
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   10
         Left            =   -72300
         TabIndex        =   138
         Top             =   6405
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   9
         Left            =   -67485
         TabIndex        =   135
         Top             =   5775
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   9
         Left            =   -72300
         TabIndex        =   134
         Top             =   5775
         Width           =   1275
      End
      Begin VB.TextBox txtObservacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6810
         Left            =   -74820
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   133
         Top             =   540
         Width           =   8700
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   7
         Left            =   -72300
         TabIndex        =   103
         Top             =   4515
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   -72300
         TabIndex        =   102
         Top             =   765
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   3
         Left            =   -72300
         TabIndex        =   101
         Top             =   2010
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   -72300
         TabIndex        =   100
         Top             =   2640
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   5
         Left            =   -72300
         TabIndex        =   99
         Top             =   3270
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   -72300
         TabIndex        =   98
         Top             =   3900
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   -72300
         TabIndex        =   97
         Top             =   5145
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   -67485
         TabIndex        =   96
         Top             =   5145
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   -67485
         TabIndex        =   95
         Top             =   3900
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   5
         Left            =   -67485
         TabIndex        =   94
         Top             =   3270
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   -67485
         TabIndex        =   93
         Top             =   2640
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   3
         Left            =   -67485
         TabIndex        =   92
         Top             =   2010
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   -67485
         TabIndex        =   91
         Top             =   765
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   7
         Left            =   -67485
         TabIndex        =   90
         Top             =   4515
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroE 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   2
         Left            =   -67485
         TabIndex        =   89
         Top             =   1395
         Width           =   1275
      End
      Begin VB.CommandButton cmdOtroN 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   2
         Left            =   -72300
         TabIndex        =   88
         Top             =   1395
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   -67485
         TabIndex        =   58
         Top             =   5865
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   -72300
         TabIndex        =   57
         Top             =   5865
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   -67485
         TabIndex        =   56
         Top             =   4470
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   -72300
         TabIndex        =   55
         Top             =   4470
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   5
         Left            =   -67485
         TabIndex        =   54
         Top             =   3780
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   5
         Left            =   -72300
         TabIndex        =   53
         Top             =   3780
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   -67485
         TabIndex        =   52
         Top             =   3075
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   -72300
         TabIndex        =   51
         Top             =   3075
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   3
         Left            =   -67485
         TabIndex        =   50
         Top             =   2385
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   3
         Left            =   -72300
         TabIndex        =   49
         Top             =   2385
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   2
         Left            =   -67485
         TabIndex        =   48
         Top             =   1680
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   2
         Left            =   -72300
         TabIndex        =   47
         Top             =   1680
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   -67485
         TabIndex        =   46
         Top             =   990
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   -72300
         TabIndex        =   45
         Top             =   990
         Width           =   1275
      End
      Begin VB.CommandButton cmdPropina 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   7
         Left            =   -67485
         TabIndex        =   44
         Top             =   5160
         Width           =   1275
      End
      Begin VB.CommandButton cmdTarjeta 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   7
         Left            =   -72300
         TabIndex        =   43
         Top             =   5160
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Ventas Efectivo MN"
         Height          =   555
         Index           =   0
         Left            =   2835
         TabIndex        =   17
         Top             =   1845
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   1
         Left            =   7470
         TabIndex        =   16
         Top             =   1845
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   4
         Left            =   2835
         TabIndex        =   15
         Top             =   3975
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Tarjetas de Crédito"
         Height          =   555
         Index           =   2
         Left            =   2835
         TabIndex        =   14
         Top             =   2610
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   6
         Left            =   2835
         TabIndex        =   13
         Top             =   3300
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Extranjera"
         Height          =   555
         Index           =   7
         Left            =   7470
         TabIndex        =   12
         Top             =   3300
         Width           =   1275
      End
      Begin VB.CommandButton cmdMovimiento 
         Caption         =   "Retiro Moneda Nacional"
         Height          =   555
         Index           =   8
         Left            =   2835
         TabIndex        =   11
         Top             =   4650
         Width           =   1275
      End
      Begin VB.Label lblReciboE 
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
         Left            =   5895
         TabIndex        =   145
         Top             =   2205
         Width           =   1410
      End
      Begin VB.Label lblReciboN 
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
         Left            =   4275
         TabIndex        =   144
         Top             =   2205
         Width           =   1410
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Left            =   -74775
         TabIndex        =   143
         Top             =   5775
         Width           =   990
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Left            =   -74775
         TabIndex        =   142
         Top             =   6405
         Width           =   990
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   10
         Left            =   -69135
         TabIndex        =   141
         Top             =   6405
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   10
         Left            =   -70785
         TabIndex        =   140
         Top             =   6405
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   9
         Left            =   -69135
         TabIndex        =   137
         Top             =   5775
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   9
         Left            =   -70785
         TabIndex        =   136
         Top             =   5775
         Width           =   1410
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Total Otros Pagos :"
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
         Left            =   -74775
         TabIndex        =   132
         Top             =   7065
         Width           =   1665
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Index           =   11
         Left            =   -70785
         TabIndex        =   131
         Top             =   7065
         Width           =   1410
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   7
         Left            =   -74775
         TabIndex        =   130
         Top             =   4515
         Width           =   990
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   7
         Left            =   -70785
         TabIndex        =   129
         Top             =   4515
         Width           =   1410
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   4
         Left            =   -74775
         TabIndex        =   128
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   8
         Left            =   -74775
         TabIndex        =   127
         Top             =   5145
         Width           =   990
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   5
         Left            =   -74775
         TabIndex        =   126
         Top             =   3270
         Width           =   990
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   3
         Left            =   -74775
         TabIndex        =   125
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   6
         Left            =   -74775
         TabIndex        =   124
         Top             =   3900
         Width           =   990
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   1
         Left            =   -74775
         TabIndex        =   123
         Top             =   765
         Width           =   990
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   8
         Left            =   -70785
         TabIndex        =   122
         Top             =   5145
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   6
         Left            =   -70785
         TabIndex        =   121
         Top             =   3900
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   5
         Left            =   -70785
         TabIndex        =   120
         Top             =   3270
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   4
         Left            =   -70785
         TabIndex        =   119
         Top             =   2640
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   3
         Left            =   -70785
         TabIndex        =   118
         Top             =   2010
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   1
         Left            =   -70785
         TabIndex        =   117
         Top             =   765
         Width           =   1410
      End
      Begin VB.Label lblMontoN 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -69915
         TabIndex        =   116
         Top             =   450
         Width           =   540
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   1
         Left            =   -69135
         TabIndex        =   115
         Top             =   765
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   3
         Left            =   -69135
         TabIndex        =   114
         Top             =   2010
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   4
         Left            =   -69135
         TabIndex        =   113
         Top             =   2640
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   5
         Left            =   -69135
         TabIndex        =   112
         Top             =   3270
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   6
         Left            =   -69135
         TabIndex        =   111
         Top             =   3900
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   8
         Left            =   -69135
         TabIndex        =   110
         Top             =   5145
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   7
         Left            =   -69135
         TabIndex        =   109
         Top             =   4515
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Index           =   11
         Left            =   -69135
         TabIndex        =   108
         Top             =   7065
         Width           =   1410
      End
      Begin VB.Label txtOtroE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   2
         Left            =   -69135
         TabIndex        =   107
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Label txtOtroN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   2
         Left            =   -70785
         TabIndex        =   106
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Label lblOtro 
         AutoSize        =   -1  'True
         Caption         =   "Otro Pago :"
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
         Index           =   2
         Left            =   -74775
         TabIndex        =   105
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label lblMontoE 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -68265
         TabIndex        =   104
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   14
         Left            =   -69915
         TabIndex        =   87
         Top             =   585
         Width           =   540
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   1
         Left            =   -70785
         TabIndex        =   86
         Top             =   990
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   2
         Left            =   -70785
         TabIndex        =   85
         Top             =   1680
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   3
         Left            =   -70785
         TabIndex        =   84
         Top             =   2385
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   4
         Left            =   -70785
         TabIndex        =   83
         Top             =   3075
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   5
         Left            =   -70785
         TabIndex        =   82
         Top             =   3780
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   6
         Left            =   -70785
         TabIndex        =   81
         Top             =   4470
         Width           =   1410
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   8
         Left            =   -70785
         TabIndex        =   80
         Top             =   5865
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   1
         Left            =   -69135
         TabIndex        =   79
         Top             =   990
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   2
         Left            =   -69135
         TabIndex        =   78
         Top             =   1680
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   3
         Left            =   -69135
         TabIndex        =   77
         Top             =   2385
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   4
         Left            =   -69135
         TabIndex        =   76
         Top             =   3075
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   5
         Left            =   -69135
         TabIndex        =   75
         Top             =   3780
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   6
         Left            =   -69135
         TabIndex        =   74
         Top             =   4470
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   8
         Left            =   -69135
         TabIndex        =   73
         Top             =   5865
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   1
         Left            =   -74820
         TabIndex        =   72
         Top             =   990
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   2
         Left            =   -74820
         TabIndex        =   71
         Top             =   1725
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   6
         Left            =   -74820
         TabIndex        =   70
         Top             =   4470
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   3
         Left            =   -74820
         TabIndex        =   69
         Top             =   2385
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   5
         Left            =   -74820
         TabIndex        =   68
         Top             =   3780
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   8
         Left            =   -74820
         TabIndex        =   67
         Top             =   5865
         Width           =   735
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   4
         Left            =   -74820
         TabIndex        =   66
         Top             =   3075
         Width           =   735
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Propina"
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
         Index           =   20
         Left            =   -68385
         TabIndex        =   65
         Top             =   585
         Width           =   660
      End
      Begin VB.Label txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   7
         Left            =   -70785
         TabIndex        =   64
         Top             =   5160
         Width           =   1410
      End
      Begin VB.Label txtPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   285
         Index           =   7
         Left            =   -69135
         TabIndex        =   63
         Top             =   5160
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
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
         Index           =   7
         Left            =   -74820
         TabIndex        =   62
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label txtTTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -70785
         TabIndex        =   61
         Top             =   6615
         Width           =   1410
      End
      Begin VB.Label txtTPropina 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -69135
         TabIndex        =   60
         Top             =   6615
         Width           =   1410
      End
      Begin VB.Label lblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Total Tarjeta :"
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
         Index           =   0
         Left            =   -74820
         TabIndex        =   59
         Top             =   6615
         Width           =   1230
      End
      Begin VB.Label txtTotalIngresoE 
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
         Left            =   5895
         TabIndex        =   42
         Top             =   5580
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Inicio de Caja :"
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
         Index           =   0
         Left            =   1380
         TabIndex        =   41
         Top             =   1305
         Width           =   1305
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Cheques / Depósitos :"
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
         Left            =   270
         TabIndex        =   40
         Top             =   3300
         Width           =   2415
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total de Tarjetas :"
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
         Left            =   1095
         TabIndex        =   39
         Top             =   2610
         Width           =   1590
      End
      Begin VB.Label txtChequeE 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5895
         TabIndex        =   38
         Top             =   3300
         Width           =   1410
      End
      Begin VB.Label txtTotalTarjetaPropina 
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
         ForeColor       =   &H00FF00FF&
         Height          =   285
         Left            =   5895
         TabIndex        =   37
         Top             =   2610
         Width           =   1410
      End
      Begin VB.Label txtInicioE 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5895
         TabIndex        =   36
         Top             =   1260
         Width           =   1410
      End
      Begin VB.Label txtEfectivoE 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5895
         TabIndex        =   35
         Top             =   1845
         Width           =   1410
      End
      Begin VB.Label txtME 
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
         Height          =   285
         Left            =   5895
         TabIndex        =   34
         Top             =   765
         Width           =   690
      End
      Begin VB.Label txtTotalOtroE 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5895
         TabIndex        =   33
         Top             =   3975
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Otros Tipos Pago :"
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
         Index           =   27
         Left            =   585
         TabIndex        =   32
         Top             =   3975
         Width           =   2100
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Puntos :"
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
         Index           =   31
         Left            =   1470
         TabIndex        =   31
         Top             =   4650
         Width           =   1215
      End
      Begin VB.Label txtPuntoE 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5895
         TabIndex        =   30
         Top             =   4650
         Width           =   1410
      End
      Begin VB.Label txtTotalOtroN 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   4260
         TabIndex        =   29
         Top             =   3975
         Width           =   1410
      End
      Begin VB.Label txtTotalTarjeta 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   4260
         TabIndex        =   28
         Top             =   2610
         Width           =   1410
      End
      Begin VB.Label txtMN 
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
         Height          =   285
         Left            =   4260
         TabIndex        =   27
         Top             =   765
         Width           =   690
      End
      Begin VB.Label txtChequeN 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   4260
         TabIndex        =   26
         Top             =   3300
         Width           =   1410
      End
      Begin VB.Label txtEfectivoN 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   4260
         TabIndex        =   25
         Top             =   1845
         Width           =   1410
      End
      Begin VB.Label txtInicioN 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   4260
         TabIndex        =   24
         Top             =   1260
         Width           =   1410
      End
      Begin VB.Label txtPuntoN 
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   4260
         TabIndex        =   23
         Top             =   4650
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Ingresos :"
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
         Index           =   23
         Left            =   1335
         TabIndex        =   22
         Top             =   5580
         Width           =   1350
      End
      Begin VB.Label txtTotalIngresoN 
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
         Left            =   4275
         TabIndex        =   21
         Top             =   5580
         Width           =   1410
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Total Efectivo :"
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
         Index           =   32
         Left            =   1350
         TabIndex        =   20
         Top             =   1845
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "( Cobrado + Recibos - Egresos )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   12
         Left            =   585
         TabIndex        =   19
         Top             =   2070
         Width           =   1980
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "( Cobrado + Recibos )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   13
         Left            =   1215
         TabIndex        =   18
         Top             =   2835
         Width           =   1350
      End
   End
   Begin VB.Label txtUsuario 
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
      Height          =   285
      Left            =   825
      TabIndex        =   3
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label TxtFecha 
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
      Height          =   285
      Left            =   7755
      TabIndex        =   2
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Usuario :"
      Height          =   195
      Index           =   19
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   630
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      Height          =   195
      Index           =   18
      Left            =   7125
      TabIndex        =   0
      Top             =   90
      Width           =   540
   End
End
Attribute VB_Name = "frmLiquidacionMozo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSumas As Recordset
Dim RsTarjetas As Recordset
Dim RsTarjeta As Recordset
Dim RsOtro As Recordset
Dim RsOtros As Recordset
Dim RsPago As Recordset
Dim RsEgreso As Recordset
Dim RsIngreso As Recordset
Dim nCambio As Double
Dim xFecha As Date
Dim nInicioN As Double
Dim nInicioE As Double
Dim nEfectivoN As Double
Dim nEfectivoE As Double
Dim nTotalTarjeta As Double
Dim nTotalTarjetaPropina As Double
Dim nTarjeta(8) As Double
Dim nTarjetaPropina(8) As Double
Dim nChequeN As Double
Dim nChequeE As Double
Dim nTotalOtroN As Double
Dim nTotalOtroE As Double
Dim nOtroN(10) As Double
Dim nOtroE(10) As Double
Dim nPuntoN As Double
Dim nPuntoE As Double
Dim nDolar As Double
Dim tTarjeta(8) As String
Dim tOtro(10) As String

Dim nRetiroN As Double
Dim nRetiroE As Double

Dim nReciboIngresoN As Double
Dim nReciboIngresoE As Double
Dim nReciboAnticipoN As Double
Dim nReciboAnticipoE As Double
Dim nReciboEgresoN As Double
Dim nReciboEgresoE As Double
Dim nReciboTarjeta As Double
Dim nCuentaCobrar As Double
Dim nCuentaCorriente As Double
Dim nTotalEfectivoN As Double
Dim nTotalEfectivoE As Double
Dim nTotalIngresoN As Double
Dim nTotalIngresoE As Double
Dim nSaldoN As Double
Dim nSaldoE As Double
Dim nEgresoN As Double
Dim nEgresoE As Double
Dim nIngresoN As Double
Dim nIngresoE As Double

Dim nFinalN As Double
Dim nFinalE As Double

Dim sMozoUsuario As String
Dim i As Integer

Private Sub cmdMovimiento_Click(Index As Integer)
  Select Case Index
         Case Is = 0 ' Efectivo MN
              sTipo = ""
              frmNumPad.Show vbModal
              nEfectivoN = IIf(wEnter = True, sDescrip, nEfectivoN)
              txtEfectivoN.Caption = Format(nEfectivoN, "###,###,###,##0.00")
              CalcularTotales
  
         Case Is = 1 ' Efectivo ME
              sTipo = ""
              frmNumPad.Show vbModal
              nEfectivoE = IIf(wEnter = True, sDescrip, nEfectivoE)
              txtEfectivoE.Caption = Format(nEfectivoE, "###,###,###,##0.00")
              CalcularTotales
                               
         Case Is = 2 ' Tarjetas de Credito
              SSTab.Tab = 1
                                             
         Case Is = 4 ' Otros Pagos
              SSTab.Tab = 2
                                                                                          
         Case Is = 6 ' Cheque MN
              sTipo = ""
              frmNumPad.Show vbModal
              nChequeN = IIf(wEnter = True, sDescrip, nChequeN)
              txtChequeN.Caption = Format(nChequeN, "###,###,###,##0.00")
              CalcularTotales

         Case Is = 7 ' Cheque MN
              sTipo = ""
              frmNumPad.Show vbModal
              nChequeE = IIf(wEnter = True, sDescrip, nChequeE)
              txtChequeE.Caption = Format(nChequeE, "###,###,###,##0.00")
              CalcularTotales
              
         Case Is = 8 ' puntos MN
              sTipo = ""
              frmNumPad.Show vbModal
              nPuntoN = IIf(wEnter = True, sDescrip, nPuntoN)
              txtPuntoN.Caption = Format(nPuntoN, "###,###,###,##0.00")
              CalcularTotales

  End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
  Dim nCorrela As String
  Select Case Index
         Case Is = 0 ' Cierre
              If lObligaCierre Then
                 If (Supervisor("11") = False) Then
                    MsgBox "Clave No Permitida", vbExclamation, sMensaje
                    Exit Sub
                 End If
              End If
              
              If MsgBox("Seguro de Cerrar el Turno " & sTurno & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                 Exit Sub
              End If
              Dim sCajon As String
              sCajon = Calcular("select tCajon as Codigo from MTURNOCAJERO where tTurno='" & sTurno & "' and tMozo='" & sMozo & "' and lCierre=0", Cn)
              Cn.Execute "update TTABLA set nBoton=0 where tTabla='CAJON' and tCodigo='" & sCajon & "'"
              
              Isql = "Update MTURNOCAJERO Set " & _
                     "lCierre = 1, " & _
                     "fFechaFin = getdate(), " & _
                     "nMontoEN =" & nEfectivoN & ", " & _
                     "nMontoEE =" & nEfectivoE & ", " & _
                     "nMontoCN =" & nChequeN & ", " & _
                     "nMontoCE =" & nChequeE & ", " & _
                     "nMontoPN =" & nPuntoN & ", " & _
                     "nMontoPE =" & nPuntoE & ", " & _
                     "nTarjeta1 =" & nTarjeta(1) & ", nTarjeta2 =" & nTarjeta(2) & ", nTarjeta3 =" & nTarjeta(3) & ", nTarjeta4 =" & nTarjeta(4) & ", nTarjeta5 =" & nTarjeta(5) & ", nTarjeta6 =" & nTarjeta(6) & ", nTarjeta7 =" & nTarjeta(7) & ", nTarjeta8 =" & nTarjeta(8) & ", " & _
                     "nPropina1 =" & nTarjetaPropina(1) & ", nPropina2 =" & nTarjetaPropina(2) & ", nPropina3 =" & nTarjetaPropina(3) & ", nPropina4 =" & nTarjetaPropina(4) & ", nPropina5 =" & nTarjetaPropina(5) & ", nPropina6 =" & nTarjetaPropina(6) & ", nPropina7 =" & nTarjetaPropina(7) & ", nPropina8 =" & nTarjetaPropina(8) & ", " & _
                     "nOtroN1 =" & nOtroN(1) & ", nOtroN2 =" & nOtroN(2) & ", nOtroN3 =" & nOtroN(3) & ", nOtroN4 =" & nOtroN(4) & ", nOtroN5 =" & nOtroN(5) & ", nOtroN6 =" & nOtroN(6) & ", nOtroN7 =" & nOtroN(7) & ", nOtroN8 =" & nOtroN(8) & ", nOtroN9 =" & nOtroN(9) & ", nOtroN10 =" & nOtroN(10) & ", " & _
                     "nOtroE1 =" & nOtroE(1) & ", nOtroE2 =" & nOtroE(2) & ", nOtroE3 =" & nOtroE(3) & ", nOtroE4 =" & nOtroE(4) & ", nOtroE5 =" & nOtroE(5) & ", nOtroE6 =" & nOtroE(6) & ", nOtroE7 =" & nOtroE(7) & ", nOtroE8 =" & nOtroE(8) & ", nOtroE9 =" & nOtroE(9) & ", nOtroE10 =" & nOtroE(10) & ", " & _
                     "tObservacion='" & txtObservacion.Text & "' " & _
                     "where tTurno = '" & sTurno & "' and tMozo='" & sMozo & "' and lCierre=0"
              Cn.Execute Isql
              MsgBox "Turno " & sTurno & " cerrado satisfactoriamente", vbInformation, sMensaje
              Unload Me
         
         Case Is = 1 ' Cancelar
              Unload Me
         
         Case Is = 2 ' Impresion
              ImprimeLiquidacion
                      
         Case Is = 4 ' Imprime Liquidacion
               With frmRepLiquidacionTicket
                    .cmdBusca(0).Enabled = False
                    .txtTurno = sTurno
                    .chkUsuario.Enabled = False
                    .chkTurno.Enabled = False
                    .chkNoCortesia.Enabled = False
                    .chkCortesia.Enabled = False
                    .chkCajero.Enabled = False
                    .chkCajero.Value = 0
                    .cboCajero.Enabled = False
                    .cboCajero.Text = txtUsuario.Caption
                    .Show vbModal
               End With
              
         Case Is = 5 ' Imprime Paloteo
               With frmRepPaloteoTicket
                    .chkTurno.Value = 0
                    .cmdBusca(1).Enabled = False
                    .txtTurno = sTurno
                    .sTurno = sTurno
                    .chkTurno.Enabled = False
                    .chkCajero.Enabled = False
                    .chkCajero.Value = 0
                    .cboCajero.Enabled = False
                    .cboCajero.Text = txtUsuario.Caption
                    .Show vbModal
               End With
               
         Case Is = 7 ' Imprime Paloteo
              SSTab.Tab = 3
         
  End Select
End Sub

Private Sub cmdOtroE_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nOtroE(Index) = IIf(wEnter = True, sDescrip, nOtroE(Index))
    txtOtroE(Index).Caption = Format(nOtroE(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub cmdOtroN_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nOtroN(Index) = IIf(wEnter = True, sDescrip, nOtroN(Index))
    txtOtroN(Index).Caption = Format(nOtroN(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub cmdPropina_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nTarjetaPropina(Index) = IIf(wEnter = True, sDescrip, nTarjetaPropina(Index))
    txtPropina(Index).Caption = Format(nTarjetaPropina(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub cmdTarjeta_Click(Index As Integer)
    sTipo = ""
    frmNumPad.Show vbModal
    nTarjeta(Index) = IIf(wEnter = True, sDescrip, nTarjeta(Index))
    txtTarjeta(Index).Caption = Format(nTarjeta(Index), "###,###,###,##0.00")
    CalcularTotales
End Sub

Private Sub Form_Load()
   Centrar Me
   Screen.MousePointer = vbHourglass
   sMozoUsuario = Calcular("select Descripcion as Codigo FROM vMOZO where Codigo = '" & sMozo & "'", Cn)
   
   If Calcular("select count(tCodigoPedido) as codigo from vPedidoDetalle where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tCaja = '" & sCaja & "')", Cn) > 0 Then
      MsgBox "Tienes Pedidos Abiertos", vbExclamation, sMensaje
   End If
              
   If Calcular("select count(tCodigoPedido) as codigo from vPedidoDetalle where tEstadoItem ='N' and tCodigoPedido in (select tCodigoPedido from MPEDIDO where tEstadoPedido='01' and tTurno = 'MOZO')", Cn) > 0 Then
      MsgBox "Tienes Pedidos Abiertos del Punto de Ventas de Mozos", vbExclamation, sMensaje
   End If

   txtUsuario.Caption = sMozoUsuario
   TxtFecha.Caption = Date
   Me.Caption = "Cierre Turno " & sTurno
   SSTab.Tab = 0
   
   If lCierre Then
      cmdOpcion(4).Enabled = False
      cmdOpcion(5).Enabled = False
      lblReciboN.Visible = False
      lblReciboE.Visible = False
   End If
        
   'Descripciones MN y ME
   txtMN.Caption = sMonN
   lblMontoN.Caption = sMonN
   cmdMovimiento(0).Caption = "Ventas Efectivo " & sMonN
   cmdMovimiento(4).Caption = "Otros Tipos Pago "
   cmdMovimiento(6).Caption = "Cheques " & sMonN
   cmdMovimiento(8).Caption = "Puntos " & sMonN
   
   If sMonE <> "" And sMonN <> sMonE Then
      txtME.Caption = sMonE
      lblMontoE.Caption = sMonE
      cmdMovimiento(1).Caption = "Ventas Efectivo " & sMonE
      cmdMovimiento(7).Caption = "Cheques " & sMonE
   Else
      txtME.Visible = False
      lblMontoE.Visible = False
      txtInicioE.Visible = False
      txtEfectivoE.Visible = False
      txtEfectivoE.Visible = False
      txtTotalOtroE.Visible = False
      txtChequeE.Visible = False
      txtPuntoE.Visible = False
      txtTotalIngresoE.Visible = False
      cmdMovimiento(1).Visible = False
      cmdMovimiento(7).Visible = False
      cmdMovimiento(11).Visible = False
      
      cmdOtroE(1).Visible = False
      cmdOtroE(2).Visible = False
      cmdOtroE(3).Visible = False
      cmdOtroE(4).Visible = False
      cmdOtroE(5).Visible = False
      cmdOtroE(6).Visible = False
      cmdOtroE(7).Visible = False
      cmdOtroE(8).Visible = False
      cmdOtroE(9).Visible = False
      cmdOtroE(10).Visible = False
      
      txtOtroE(1).Visible = False
      txtOtroE(2).Visible = False
      txtOtroE(3).Visible = False
      txtOtroE(4).Visible = False
      txtOtroE(5).Visible = False
      txtOtroE(6).Visible = False
      txtOtroE(7).Visible = False
      txtOtroE(8).Visible = False
      txtOtroE(9).Visible = False
      txtOtroE(10).Visible = False
   End If
                        
   'Fecha del Turno
   xFecha = Calcular("select fInicial as codigo from MTURNO where tTurno='" & sTurno & "'", Cn)
      
   'Tipo de Cambio del turno
   nCambio = Calcular("select nVenta as Codigo from TTIPOCAMBIO where fFecha='" & Format(xFecha, "yyyy/mm/dd") & "'", Cn)
   
   'Saldo Inicial
   Set RsSumas = Lib.OpenRecordset("select * from MTURNO where tTurno ='" & sTurno & "'", Cn)
   If RsSumas.RecordCount = 0 Then
      nInicioN = 0
      nInicioE = 0
   Else
      nInicioN = RsSumas!nMontoIN
      nInicioE = RsSumas!nMontoIE
   End If
   
   LlenaTarjeta
   LlenaOtro
      
   Asignar
   nTotalIngresoN = nEfectivoN + nTotalTarjeta + nChequeN + nTotalOtroN + nPuntoN + nEgresoN
   nTotalIngresoE = nEfectivoE + nChequeE + nTotalOtroE + nEgresoE
   nTotalEfectivoN = nEfectivoN
   nTotalEfectivoE = nEfectivoE
   nFinalN = nTotalEfectivoN - nSaldoN
   nFinalE = nTotalEfectivoE - nSaldoE
   
   'Llena Pantalla
   lblReciboN.Caption = Format(nEgresoN, "###,###,##0.00")
   lblReciboE.Caption = Format(nEgresoE, "###,###,##0.00")
   
   txtInicioN.Caption = Format(nInicioN, "###,###,##0.00")
   txtInicioE.Caption = Format(nInicioE, "###,###,##0.00")
   txtEfectivoN.Caption = Format(nEfectivoN, "###,###,##0.00")
   txtEfectivoE.Caption = Format(nEfectivoE, "###,###,##0.00")
   txtTotalTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
   txtTotalTarjetaPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
   txtChequeN.Caption = Format(nChequeN, "###,###,##0.00")
   txtChequeE.Caption = Format(nChequeE, "###,###,##0.00")
   txtTotalOtroN.Caption = Format(nTotalOtroN, "###,###,##0.00")
   txtTotalOtroE.Caption = Format(nTotalOtroE, "###,###,##0.00")
   txtPuntoN.Caption = Format(nPuntoN, "###,###,##0.00")
   txtPuntoE.Caption = Format(nPuntoE, "###,###,##0.00")
   txtTotalIngresoN.Caption = Format(nTotalIngresoN, "###,###,##0.00")
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmLiquidacionDetalle = Nothing
End Sub

Public Sub LlenaTarjeta()
   Set RsTarjetas = Lib.OpenRecordset("select tCodigoTarjeta, tDetallado, tResumido from tTarjetaCredito", Cn)
   For i = 1 To 8
       RsTarjetas.MoveFirst
       RsTarjetas.Find "tcodigoTarjeta = '0" & Trim(Str(i)) & "'"
       If RsTarjetas.EOF Then
          lblTarjeta(i).Visible = False
          cmdTarjeta(i).Visible = False
          cmdPropina(i).Visible = False
          txtTarjeta(i).Visible = False
          txtPropina(i).Visible = False
       Else
          lblTarjeta(i).Caption = RsTarjetas!tDetallado
          cmdTarjeta(i).Caption = RsTarjetas!tResumido
          cmdPropina(i).Caption = RsTarjetas!tResumido
          tTarjeta(i) = RsTarjetas!tCodigoTarjeta
          If Not lCierre Then
             cmdTarjeta(i).Enabled = False
             cmdPropina(i).Enabled = False
          End If
       End If
   Next i
End Sub

Public Sub CalcularTotales()
   'Calcula Tarjeta
   nTotalTarjeta = 0
   nTotalTarjetaPropina = 0
   nTotalOtroN = 0
   nTotalOtroE = 0
   For i = 1 To 8
       nTotalTarjeta = nTotalTarjeta + nTarjeta(i)
       nTotalTarjetaPropina = nTotalTarjetaPropina + nTarjetaPropina(i)
   Next i
   
   For i = 1 To 10
       nTotalOtroN = nTotalOtroN + nOtroN(i)
       nTotalOtroE = nTotalOtroE + nOtroE(i)
   Next i
   
   txtTTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
   txtTPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
   txtTotalTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
   txtTotalTarjetaPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
   
   txtTotalOtroN.Caption = Format(nTotalOtroN, "###,###,##0.00")
   txtTotalOtroE.Caption = Format(nTotalOtroE, "###,###,##0.00")
'   txtOtroN(9).Caption = Format(nTotalOtroN, "###,###,##0.00")
'   txtOtroE(9).Caption = Format(nTotalOtroE, "###,###,##0.00")
      
   'Total de ingreso
   nTotalIngresoN = nEfectivoN + nTotalTarjeta + nChequeN + nTotalOtroN + nPuntoN + nEgresoN
   nTotalIngresoE = nEfectivoE + nChequeE + nTotalOtroE + nEgresoE
   
   'Total de Efectivo
   nTotalEfectivoN = nEfectivoN
   nTotalEfectivoE = nEfectivoE
   
   nFinalN = nTotalEfectivoN - nRetiroN
   nFinalE = nTotalEfectivoE - nRetiroE

   'Saldo en Efectivo
   txtTotalIngresoN.Caption = Format(nTotalIngresoN, "###,###,##0.00")
   txtTotalIngresoE.Caption = Format(nTotalIngresoE, "###,###,##0.00")
  ' txtTotalEfectivoN.Caption = Format(nTotalEfectivoN, "###,###,##0.00")
  ' txtTotalEfectivoE.Caption = Format(nTotalEfectivoE, "###,###,##0.00")
  ' txtSaldoN.Caption = Format(nFinalN, "###,###,##0.00")
  ' txtSaldoE.Caption = Format(nFinalE, "###,###,##0.00")
End Sub

Public Sub ImprimeLiquidacion()
   Dim sTitulo1 As String
   Dim sTitulo2 As String
   Dim sTitulo3 As String
   
   Screen.MousePointer = vbHourglass
   'Configura la impresora la impresion Font
   Imprimir (sPreCuenta)
   Printer.FontName = sFont
   Printer.FontBold = False
   
   'Cabecera
   ImprimeXCentro "Liquidación de Cajero", 40
   ImprimeXCentro sRazonSocial, 40
   Printer.Print ""
      
   sTitulo1 = ""
   sTitulo2 = ""
   sTitulo1 = "Turno   : " & sTurno
   sTitulo2 = "Fecha   : " & Format(Now, "dd MMMM yyyy HH:mm") & " Hrs "
                
   Printer.Print sTitulo1
   Printer.Print sTitulo2
   Printer.Print ""
   
   Printer.Print String(40, "-")
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "CONCEPTO                  " & sMonN & "        " & sMonE
   Else
      Printer.Print "CONCEPTO                  " & sMonN
   End If
   
   Printer.Print String(40, "-")
               
   'Fondo de Caja
   Printer.Print
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Fondo de Caja   : " & Right(String(11, " ") & Format(nInicioN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nInicioE, "###,##0.00"), 11)
   Else
      Printer.Print "Fondo de Caja   : " & Right(String(11, " ") & Format(nInicioN, "###,##0.00"), 11)
   End If
      
   'Efectivo
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "       Efectivo : " & Right(String(11, " ") & Format(nEfectivoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nEfectivoE, "###,##0.00"), 11)
   Else
      Printer.Print "       Efectivo : " & Right(String(11, " ") & Format(nEfectivoN, "###,##0.00"), 11)
   End If
   
   'Total de Tarjetas
   Printer.Print ""
   Printer.Print "Tarjetas Credito: " & Right(String(11, " ") & Format(nTotalTarjeta, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalTarjetaPropina, "###,##0.00"), 11)
   For i = 1 To 8
       If lblTarjeta(i).Visible = True Then
          Printer.Print " - " & Mid(cmdTarjeta(i).Caption & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nTarjeta(i), "###,##0.00"), 11) & Right(String(11, " ") & Format(nTarjetaPropina(i), "###,##0.00"), 11)
       End If
   Next i

   'Total de Cheques
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Cheque/Deposito : " & Right(String(11, " ") & Format(nChequeN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nChequeE, "###,##0.00"), 11)
   Else
      Printer.Print "Cheque/Deposito : " & Right(String(11, " ") & Format(nChequeN, "###,##0.00"), 11)
   End If

   'Otros Tipos de Pago
   Printer.Print ""
   Printer.Print "Otros Tipos Pago: " & Right(String(11, " ") & Format(nTotalOtroN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalOtroE, "###,##0.00"), 11)
   For i = 1 To 10
       If sMonE <> "" And sMonN <> sMonE Then
          If lblOtro(i).Visible = True Then
             Printer.Print " - " & Mid(lblOtro(i).Caption & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nOtroN(i), "###,##0.00"), 11) & Right(String(11, " ") & Format(nOtroE(i), "###,##0.00"), 11)
          End If
       Else
          If lblOtro(i).Visible = True Then
             Printer.Print " - " & Mid(lblOtro(i).Caption & String(15, " "), 1, 15) & Right(String(11, " ") & Format(nOtroN(i), "###,##0.00"), 11)
          End If
       End If
   Next i
 
   'Puntos
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Puntos          : " & Right(String(11, " ") & Format(nPuntoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nPuntoE, "###,##0.00"), 11)
   Else
      Printer.Print "Puntos          : " & Right(String(11, " ") & Format(nPuntoN, "###,##0.00"), 11)
   End If

   'Total Ingreso en Caja
   Printer.Print ""
   Printer.Print String(40, "-")
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Total Ingreso   : " & Right(String(11, " ") & Format(nTotalIngresoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalIngresoE, "###,##0.00"), 11)
   Else
      Printer.Print "Total Ingreso   : " & Right(String(11, " ") & Format(nTotalIngresoN, "###,##0.00"), 11)
   End If

   'Total Efectivo en Caja
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Total Efectivo  : " & Right(String(11, " ") & Format(nTotalEfectivoN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalEfectivoE, "###,##0.00"), 11)
   Else
      Printer.Print "Total Efectivo  : " & Right(String(11, " ") & Format(nTotalEfectivoN, "###,##0.00"), 11)
   End If

   'Retiro de Caja
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Retiro Efectivo : " & Right(String(11, " ") & Format(nRetiroN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nRetiroE, "###,##0.00"), 11)
   Else
      Printer.Print "Retiro Efectivo : " & Right(String(11, " ") & Format(nRetiroN, "###,##0.00"), 11)
   End If

   'Saldo de Caja
   Printer.Print ""
   If sMonE <> "" And sMonN <> sMonE Then
      Printer.Print "Saldo en Caja   : " & Right(String(11, " ") & Format(nTotalEfectivoN - nRetiroN, "###,##0.00"), 11) & Right(String(11, " ") & Format(nTotalEfectivoE - nRetiroE, "###,##0.00"), 11)
   Else
      Printer.Print "Saldo en Caja   : " & Right(String(11, " ") & Format(nTotalEfectivoN - nRetiroN, "###,##0.00"), 11)
   End If

   Printer.Print " "
   Printer.Print String(40, "-")
   Printer.Print " "
   Printer.Print " "
   Printer.Print " "
   Printer.Print " "
   Printer.Print String(25, " ") & "---------------"
   Printer.Print String(25, " ") & String((15 - Len(RTrim(sUsuario))) / 2, " ") & sUsuario
   Printer.Print " "
   Printer.Print String(40, "-")
   
   Dim sDoc As String
   Dim sTipoDoc As String
   Dim nCantidadDoc As Integer
   
   For i = 1 To Calcular("select count(tCodigo) as codigo From TTABLA Where TTABLA = 'TIPODOCUMENTO'", Cn)
       sTipoDoc = Mid("00", 1, 2 - Len(Trim(Str(i)))) & Trim(Str(i))
       sDoc = Calcular("Select tDetallado As Codigo from TTABLA Where TTABLA = 'TIPODOCUMENTO' AND tCODIGO = '" & sTipoDoc & "'", Cn)
       nCantidadDoc = Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn)
       If sDoc <> "0" And nCantidadDoc > 0 Then
          Printer.Print " "
          Printer.Print sDoc
          Printer.Print "Del     : " & Format(Calcular("select min(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Al      : " & Format(Calcular("select max(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Emitido : " & Format(Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tEstadoDocumento<>'04' and tTurno='" & sTurno & "'", Cn), "##,##0")
          Printer.Print "Anulado : " & Format(Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tTipoDocumento='" & sTipoDoc & "' and tEstadoDocumento='04' and tTurno='" & sTurno & "'", Cn), "##,##0")
       End If
   Next i
   Printer.Print ""
   If Len(Trim(txtObservacion.Text)) > 0 Then
      Printer.Print String(40, "-")
      ImprimeXLinea txtObservacion.Text, 40, 0
   End If
   Printer.Print String(40, "-")
   Printer.Print "Caja    : " & sCaja
   Printer.Print " "
   Printer.EndDoc
   Screen.MousePointer = vbDefault
End Sub

Public Sub LlenaOtro()
   Set RsOtros = Lib.OpenRecordset("select * from vTipoCancelacion where lActivo=1", Cn)
   If RsOtros.RecordCount = 0 Then
      cmdMovimiento(4).Enabled = False
      Exit Sub
   End If
   RsOtros.MoveFirst
   
   For i = 1 To 10
       If RsOtros.EOF Then
          lblOtro(i).Visible = False
          cmdOtroN(i).Visible = False
          cmdOtroE(i).Visible = False
          txtOtroN(i).Visible = False
          txtOtroE(i).Visible = False
       Else
          lblOtro(i).Caption = RsOtros!Descripcion
          cmdOtroN(i).Caption = RsOtros!tResumido
          cmdOtroE(i).Caption = RsOtros!tResumido
          tOtro(i) = RsOtros!Codigo
          If Not lCierre Then
             cmdOtroN(i).Enabled = False
             cmdOtroE(i).Enabled = False
          End If
          RsOtros.MoveNext
       End If
   Next i
      
End Sub

Public Sub Asignar()
   nDolar = 0
   nEfectivoN = 0
   nEfectivoE = 0
   nTotalTarjeta = 0
   nTotalTarjetaPropina = 0
   nChequeN = 0
   nChequeE = 0
   nTotalOtroN = 0
   nTotalOtroE = 0
   nPuntoN = 0
   nPuntoE = 0
   
   If Not lCierre Then
'      'Recibos de Egreso
'      nEgresoN = 0
'      nEgresoE = 0
'      Isql = "select tMoneda, sum(nMonto) as nMonto From megreso where tEstadoDocumento='01' and tTurno='" & sTurno & "' Group by tMoneda"
'      Set RsEgreso = Lib.OpenRecordset(Isql, Cn)
'      If RsEgreso.RecordCount > 0 Then
'         Do While Not RsEgreso.EOF
'            If RsEgreso!tMoneda = "01" Then
'               nEgresoN = IIf(IsNull(RsEgreso!nMonto), 0, RsEgreso!nMonto)
'            Else
'               nEgresoE = IIf(IsNull(RsEgreso!nMonto), 0, RsEgreso!nMonto)
'            End If
'            RsEgreso.MoveNext
'         Loop
'      End If

'      'Recibos de Ingreso
'      nIngresoN = 0
'      nIngresoE = 0
'      Isql = "select tMoneda, sum(nMonto) as nMonto From mingreso where tEstadoDocumento='01' and tTurno='" & sTurno & "' Group by tMoneda"
'      Set RsIngreso = Lib.OpenRecordset(Isql, Cn)
'      If RsIngreso.RecordCount > 0 Then
'         Do While Not RsIngreso.EOF
'            If RsIngreso!tMoneda = "01" Then
'               nIngresoN = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
'            Else
'               nIngresoE = IIf(IsNull(RsIngreso!nMonto), 0, RsIngreso!nMonto)
'            End If
'            RsIngreso.MoveNext
'         Loop
'      End If
'
      Isql = "select tTipoPago, tMoneda, tTarjeta, tOtroTipoPago, sum(nMonto) as nMonto, sum(nPropina) as nPropina, sum(nDolar) as nDolar from dpagodocumento " & _
             "where tTurno='" & sTurno & "' and tUsuario='" & Mid(sMozoUsuario, 1, 15) & "' group by tTipoPago, tMoneda, tTarjeta, tOtroTipoPago"
      Set RsPago = Lib.OpenRecordset(Isql, Cn)
      If RsPago.RecordCount > 0 Then
         cmdMovimiento(0).Enabled = False
         cmdMovimiento(1).Enabled = False
         cmdMovimiento(6).Enabled = False
         cmdMovimiento(7).Enabled = False
         cmdMovimiento(8).Enabled = False
         RsPago.MoveFirst
         Do While Not RsPago.EOF
            Select Case RsPago!tTipoPago
                   Case Is = "01"  'Efectivo
                        If RsPago!tMoneda = "01" Then
                           nEfectivoN = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        Else
                           nEfectivoE = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                           nDolar = IIf(IsNull(RsPago!nDolar), 0, RsPago!nDolar)
                        End If
                        
                   Case Is = "02"  'TC
                        For i = 1 To 8
                            If tTarjeta(i) = IIf(IsNull(RsPago!tTarjeta), "", RsPago!tTarjeta) Then
                               nTarjeta(i) = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                               nTarjetaPropina(i) = IIf(IsNull(RsPago!nPropina), 0, RsPago!nPropina)
                               txtTarjeta(i).Caption = Format(nTarjeta(i), "###,###,##0.00")
                               txtPropina(i).Caption = Format(nTarjetaPropina(i), "###,###,##0.00")
                               nTotalTarjeta = nTotalTarjeta + nTarjeta(i)
                               nTotalTarjetaPropina = nTotalTarjetaPropina + nTarjetaPropina(i)
                               txtTTarjeta.Caption = Format(nTotalTarjeta, "###,###,##0.00")
                               txtTPropina.Caption = Format(nTotalTarjetaPropina, "###,###,##0.00")
                               Exit For
                            End If
                        Next i
                   
                   Case Is = "03"  'Cheque
                        If RsPago!tMoneda = "01" Then
                           nChequeN = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        Else
                           nChequeE = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        End If
                                   
                   Case Is = "04"  'Otros Pago
                        For i = 1 To 10
                            If tOtro(i) = IIf(IsNull(RsPago!tOtroTipoPago), "", RsPago!tOtroTipoPago) Then
                               If RsPago!tMoneda = "01" Then
                                  nOtroN(i) = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                                  nTotalOtroN = nTotalOtroN + nOtroN(i)
                               Else
                                  nOtroE(i) = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                                  nTotalOtroE = nTotalOtroE + nOtroE(i)
                               End If
                               txtOtroN(i).Caption = Format(nOtroN(i), "###,###,##0.00")
                               txtOtroE(i).Caption = Format(nOtroE(i), "###,###,##0.00")
                               txtOtroN(11).Caption = Format(nTotalOtroN, "###,###,##0.00")
                               txtOtroE(11).Caption = Format(nTotalOtroE, "###,###,##0.00")
                               Exit For
                            End If
                        Next i
                        
                   Case Is = "05"  'Puntos
                        If RsPago!tMoneda = "01" Then
                           nPuntoN = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        Else
                           nPuntoE = IIf(IsNull(RsPago!nMonto), 0, RsPago!nMonto)
                        End If
                   
            End Select
            RsPago.MoveNext
         Loop
         
         nEfectivoN = nEfectivoN + nIngresoN - nEgresoN - ((nDolar - nEfectivoE) * nCambio)
         nEfectivoE = nDolar + nIngresoE - nEgresoE
      End If
   End If
End Sub

