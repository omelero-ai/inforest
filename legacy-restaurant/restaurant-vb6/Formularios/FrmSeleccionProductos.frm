VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmSeleccionProductos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5640
   ClientLeft      =   5775
   ClientTop       =   5085
   ClientWidth     =   8400
   Icon            =   "FrmSeleccionProductos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8400
   Begin VB.CommandButton Adicionar 
      Appearance      =   0  'Flat
      Caption         =   "Historial"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Index           =   4
      Left            =   3270
      Picture         =   "FrmSeleccionProductos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Muestra los últimos datos de compra"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CheckBox ChkExclusivo 
      Caption         =   "Sólo de este proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4530
      TabIndex        =   20
      Top             =   3285
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton Adicionar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Index           =   3
      Left            =   2190
      Picture         =   "FrmSeleccionProductos.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cancela los datos ingresados"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Adicionar 
      Appearance      =   0  'Flat
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Index           =   0
      Left            =   1110
      Picture         =   "FrmSeleccionProductos.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Graba el artículo como item"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CheckBox chkRecargo 
      Alignment       =   1  'Right Justify
      Caption         =   "Recargo"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8865
      TabIndex        =   38
      Top             =   2565
      Width           =   1350
   End
   Begin VB.TextBox TxtRecargo 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8910
      TabIndex        =   37
      Text            =   "0"
      Top             =   2070
      Width           =   660
   End
   Begin VB.CommandButton CmdPRecargo 
      Caption         =   "%"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Aplica el porcentaje a recargar"
      Top             =   2100
      Width           =   255
   End
   Begin VB.TextBox TxtPRecargo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   230
      Left            =   9450
      TabIndex        =   35
      Text            =   "0"
      Top             =   2130
      Width           =   540
   End
   Begin VB.CommandButton CmdMRecargo 
      Caption         =   "M"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10335
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Aplica el monto a recargar"
      Top             =   2100
      Width           =   255
   End
   Begin VB.CommandButton Adicionar 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Index           =   1
      Left            =   7275
      Picture         =   "FrmSeleccionProductos.frx":1488
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sale sin grabar datos"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Adicionar 
      Appearance      =   0  'Flat
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Index           =   2
      Left            =   30
      Picture         =   "FrmSeleccionProductos.frx":157A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Adiciona un nuevo registro"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame FramAdicionaDocumento 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   15
      TabIndex        =   25
      Top             =   -75
      Width           =   8355
      Begin VB.TextBox TxtPorcentajeLeyAD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7155
         TabIndex        =   69
         Text            =   "0"
         Top             =   1275
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox TxtPrecioCompra 
         Enabled         =   0   'False
         Height          =   330
         Left            =   7065
         TabIndex        =   68
         Top             =   330
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtIRBP_ICE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0EADB&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   320
         Left            =   5505
         TabIndex        =   65
         Text            =   "0.000"
         ToolTipText     =   "No afecto a IVA"
         Top             =   1065
         Width           =   1020
      End
      Begin VB.TextBox TxtRet3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10455
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "0"
         Top             =   915
         Width           =   660
      End
      Begin VB.CheckBox ChkRet3 
         Alignment       =   1  'Right Justify
         Caption         =   "Varios"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9060
         TabIndex        =   62
         Top             =   915
         Width           =   1350
      End
      Begin VB.TextBox TxtICE 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7170
         TabIndex        =   58
         Text            =   "0"
         Top             =   1725
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox ChkICE 
         Alignment       =   1  'Right Justify
         Caption         =   "I.C.E."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5760
         TabIndex        =   57
         Top             =   1710
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CheckBox chkIUE_IT 
         Alignment       =   1  'Right Justify
         Caption         =   "Asumir I.U.E. e I.T."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6585
         TabIndex        =   56
         Top             =   2940
         Width           =   1650
      End
      Begin VB.CheckBox ChkRet2 
         Alignment       =   1  'Right Justify
         Caption         =   "Varios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5760
         TabIndex        =   55
         Top             =   2565
         Width           =   1350
      End
      Begin VB.TextBox TxtRet2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "0"
         Top             =   2565
         Width           =   660
      End
      Begin VB.TextBox TxtRet1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "0"
         Top             =   2145
         Width           =   645
      End
      Begin VB.CheckBox ChkRet1 
         Alignment       =   1  'Right Justify
         Caption         =   "Varios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5760
         TabIndex        =   52
         Top             =   2130
         Width           =   1350
      End
      Begin VB.CommandButton cmdFind 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5580
         Picture         =   "FrmSeleccionProductos.frx":1AAC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda de Búsqueda/Adición"
         Top             =   690
         UseMaskColor    =   -1  'True
         Width           =   380
      End
      Begin VB.TextBox txtArticulo 
         BackColor       =   &H00F0EADB&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   705
         Width           =   4320
      End
      Begin VB.Frame fraCuenta 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   7335
         TabIndex        =   49
         Top             =   645
         Visible         =   0   'False
         Width           =   3825
         Begin VB.TextBox txtCC 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3075
            MaxLength       =   4
            TabIndex        =   3
            Top             =   90
            Width           =   675
         End
         Begin VB.TextBox txtCuenta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   750
            MaxLength       =   15
            TabIndex        =   2
            Top             =   90
            Width           =   1605
         End
         Begin VB.Label LCantidad 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C. Costo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   2400
            TabIndex        =   51
            Top             =   135
            Width           =   645
         End
         Begin VB.Label LCantidad 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   135
            Width           =   555
         End
      End
      Begin VB.CheckBox ChkImp3 
         Alignment       =   1  'Right Justify
         Caption         =   "Varios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3210
         TabIndex        =   10
         Top             =   2520
         Width           =   1350
      End
      Begin VB.CommandButton CmdMDescuento 
         Caption         =   "M"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6165
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Aplica el monto a descontar"
         Top             =   2955
         Width           =   255
      End
      Begin VB.TextBox TxtPDescuento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   230
         Left            =   5280
         TabIndex        =   27
         Text            =   "0"
         Top             =   2970
         Width           =   540
      End
      Begin VB.CommandButton CmdPDescuento 
         Caption         =   "%"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5865
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Aplica el porcentaje a descontar"
         Top             =   2955
         Width           =   255
      End
      Begin VB.CheckBox ChkDescuento 
         Alignment       =   1  'Right Justify
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3210
         TabIndex        =   12
         Top             =   2925
         Width           =   1350
      End
      Begin VB.TextBox TxtDescuento 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   4620
         TabIndex        =   13
         Text            =   "0"
         Top             =   2925
         Width           =   660
      End
      Begin VB.TextBox TxtPrecioUnitario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1665
         TabIndex        =   23
         Text            =   "0.000"
         Top             =   2520
         Width           =   1410
      End
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0EADB&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "0.000"
         Top             =   2100
         Width           =   1410
      End
      Begin VB.TextBox TxtCantidadPedida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0EADB&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   1665
         TabIndex        =   4
         Text            =   "0.000"
         Top             =   1695
         Width           =   1410
      End
      Begin VB.TextBox TxtCodigo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   22
         Top             =   345
         Width           =   1155
      End
      Begin VB.TextBox TxtImp3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   2520
         Width           =   660
      End
      Begin VB.TextBox TxtImp2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   2100
         Width           =   645
      End
      Begin VB.TextBox TxtImp1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   1695
         Width           =   645
      End
      Begin VB.CheckBox ChkImp1 
         Alignment       =   1  'Right Justify
         Caption         =   "I.G.V."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3210
         TabIndex        =   6
         Top             =   1680
         Width           =   1350
      End
      Begin VB.CheckBox ChkImp2 
         Alignment       =   1  'Right Justify
         Caption         =   "Servicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3210
         TabIndex        =   8
         Top             =   2085
         Width           =   1350
      End
      Begin MSDataListLib.DataCombo CboUnidad 
         Height          =   330
         Left            =   1200
         TabIndex        =   21
         Top             =   1065
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         BackColor       =   14737632
         ForeColor       =   -2147483647
         ListField       =   "Descripcion"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LPrecioUnitario 
         AutoSize        =   -1  'True
         Caption         =   "Otros Cargos No Afectos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   2
         Left            =   3210
         TabIndex        =   67
         Top             =   1110
         Width           =   1845
      End
      Begin VB.Label LblMoneda 
         Alignment       =   1  'Right Justify
         Caption         =   "US$"
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
         Height          =   195
         Index           =   4
         Left            =   5070
         TabIndex        =   66
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label LblRet3 
         Caption         =   "%"
         Height          =   195
         Left            =   11145
         TabIndex        =   64
         Top             =   990
         Width           =   120
      End
      Begin VB.Label LblMoneda 
         Alignment       =   1  'Right Justify
         Caption         =   "US$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7830
         TabIndex        =   61
         Top             =   1770
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblRet2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   60
         Top             =   2640
         Width           =   180
      End
      Begin VB.Label LblRet1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   59
         Top             =   2235
         Width           =   180
      End
      Begin VB.Label LblPorcionable 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Porcionable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5985
         TabIndex        =   42
         Top             =   735
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LUnidadEntrada 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unidad Kardex:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   45
         TabIndex        =   48
         Top             =   1110
         Width           =   1110
      End
      Begin VB.Label LblMoneda 
         Alignment       =   1  'Right Justify
         Caption         =   "US$"
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
         Height          =   195
         Index           =   1
         Left            =   1215
         TabIndex        =   45
         Top             =   2580
         Width           =   375
      End
      Begin VB.Label LblMoneda 
         Alignment       =   1  'Right Justify
         Caption         =   "US$"
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
         Height          =   195
         Index           =   0
         Left            =   1215
         TabIndex        =   44
         Top             =   2190
         Width           =   375
      End
      Begin VB.Label LblImp3 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5310
         TabIndex        =   41
         Top             =   2595
         Width           =   180
      End
      Begin VB.Label LblImp2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5310
         TabIndex        =   40
         Top             =   2175
         Width           =   180
      End
      Begin VB.Label LblImp1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5310
         TabIndex        =   39
         Top             =   1740
         Width           =   180
      End
      Begin VB.Label LStockActual 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Stock Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   195
         TabIndex        =   33
         Top             =   2955
         Width           =   960
      End
      Begin VB.Label LPrecioUnitario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Precio Unitario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   32
         Top             =   2565
         Width           =   1080
      End
      Begin VB.Label LCantidad 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label LTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   765
         TabIndex        =   30
         Top             =   2175
         Width           =   390
      End
      Begin VB.Label LArticulo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   555
         TabIndex        =   29
         Top             =   735
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   615
         TabIndex        =   28
         Top             =   390
         Width           =   540
      End
      Begin VB.Label LblStockActuales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.000"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1665
         TabIndex        =   24
         Top             =   2925
         Width           =   1410
      End
      Begin VB.Label LblDescripcion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Artículo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   4770
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6030
         TabIndex        =   43
         Top             =   765
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin MSComCtl2.DTPicker dtpVencimiento 
      Height          =   315
      Left            =   8775
      TabIndex        =   46
      ToolTipText     =   "Fecha de Vencimiento del Artículo"
      Top             =   270
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   15788763
      CalendarTitleBackColor=   -2147483635
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   54657027
      CurrentDate     =   38201
   End
   Begin TrueOleDBGrid80.TDBGrid grdGrilla 
      Height          =   1920
      Left            =   30
      TabIndex        =   71
      Top             =   3720
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   3387
      _LayoutType     =   4
      _RowHeight      =   19
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
      Caption         =   "Ultimos Datos de Compra del Artículo"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
      _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(9)   =   ":id=4,.fontname=Arial"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=0"
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
      _StyleDefs(41)  =   ":id=33,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=33,.fontname=Arial"
      _StyleDefs(43)  =   "Named:id=34:Heading"
      _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(47)  =   ":id=34,.fontname=Arial"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
      _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(53)  =   ":id=36,.fontname=Arial"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
      _StyleDefs(57)  =   ":id=37,.fontname=Arial"
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
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Height          =   1800
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   3175
      _LayoutType     =   4
      _RowHeight      =   19
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
      Caption         =   "Ultimos Datos de Compra del Artículo"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.locked=0"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.alignment=2,.bgcolor=&H8000000A&,.fgcolor=&H0&"
      _StyleDefs(8)   =   ":id=4,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(9)   =   ":id=4,.fontname=Arial"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bgcolor=&H80000000&,.borderSize=1,.bold=0"
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
      _StyleDefs(41)  =   ":id=33,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=33,.fontname=Arial"
      _StyleDefs(43)  =   "Named:id=34:Heading"
      _StyleDefs(44)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&HC0C0C0&"
      _StyleDefs(45)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1,.locked=0,.borderSize=1,.bold=-1"
      _StyleDefs(46)  =   ":id=34,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(47)  =   ":id=34,.fontname=Arial"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&HE7FAB6&,.fgcolor=&H0&,.borderColor=&H808000&"
      _StyleDefs(52)  =   ":id=36,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(53)  =   ":id=36,.fontname=Arial"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(56)  =   ":id=37,.strikethrough=0,.charset=0"
      _StyleDefs(57)  =   ":id=37,.fontname=Arial"
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
   Begin VB.Label LblVencimiento 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8595
      TabIndex        =   47
      ToolTipText     =   "Vencimiento del Artículo"
      Top             =   330
      Width           =   870
   End
End
Attribute VB_Name = "FrmSeleccionProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RsBusca        As Recordset
Public lCriterio      As Boolean
Dim RsMax             As Recordset
Dim rsDes             As Recordset
Dim RsMedida          As Recordset
Dim RsUnidad          As Recordset
Dim RsBGrilla         As Recordset
Dim RsProveedorPrecio As Recordset
Dim sCodi As String, sUnid As String, sNeto As String, sPuni As String, sTota As String, sPorcentajes As String, sPorcentajes2 As String
Dim ntImp             As String
Dim ntImp2             As String
Dim IMP1, Imp2, Imp3  As String
Dim sSubTot           As Double
Dim sDocumento        As String
Dim sPVen             As String
Public sTotal         As String
Public sCant          As String
Dim sPrec             As String
Dim iRow              As Integer
Public tTipoIngreso   As String
Dim dPrecioUnitario   As Double
Dim sBsc              As String
Dim lContabilidad     As Boolean
Dim lCodigoBarra      As Boolean
Dim Inicio            As Boolean
Dim Total_Ant         As Double
Dim PrecioUnitario_Ant As Double
Dim lDescontado As Boolean
Dim clsAlmac          As clsAlmacen
Public tEstadoDocu As String
Dim sAtotal As Double
Dim lAuditoria As Boolean
Dim nOtrosCargosInafecto As Double
Dim Flagloctem As Boolean
Public lPrecioCompra As Boolean
Public lAgregaItem As Boolean

Dim SwTmp As Boolean
Dim lDesc As String
Dim swPunto As Boolean
Dim swEntrar As Boolean
Dim nVPrecio As Double
Dim simp1, simp2, simp3 As Double
Dim sUsuarioAutoriza As String

Private Sub Adicionar_Click(Index As Integer)
    Select Case Index
        Case Is = 0 'Grabar
            'Control Mensajes
            If Len(Trim(TxtCodigo)) < 7 Then
                If Len(Trim(txtArticulo)) > 1 Then
                    MsgBox "Articulo no existe", vbInformation, sMensaje
                    'txtArticulo.SetFocus
                    Exit Sub
                End If
            End If
            If Trim(TxtCantidadPedida) = "" Then TxtCantidadPedida = "0"
            If Trim(TxtTotal) = "" Then TxtTotal = "0"
            If Len(Trim(TxtCodigo)) = 0 And Len(Trim(txtArticulo)) = 0 Then MsgBox "Ingrese el artículo...": Exit Sub
            If Len(Trim(TxtCodigo)) = 0 Then MsgBox "Ingrese el artículo...": txtArticulo.SetFocus: Exit Sub
            If CDbl(TxtCantidadPedida) = 0 Then MsgBox "Ingrese la cantidad...": TxtCantidadPedida.SetFocus: Exit Sub
            If CDbl(TxtTotal) = 0 Then MsgBox "Ingrese el total...": TxtTotal.SetFocus: Exit Sub
            If Not IsNumeric(TxtCantidadPedida) Then MsgBox "Ingrese la cantidad...", vbInformation, sMensaje: TxtCantidadPedida.SetFocus: TxtCantidadPedida = "": Exit Sub

            InsertarData

            'Adicionar Items
            Adicionar(0).Enabled = False
            Adicionar(3).Enabled = False
            Adicionar(2).Enabled = True
            Adicionar(2).SetFocus
            
        Case Is = 1  'Salir
            Unload Me
        
        Case Is = 2  'Nuevo
            lAgregaItem = True
            Me.chkIUE_IT.value = 0
            If Me.Height = 5985 Then Adicionar_Click (4)
            PantallaD
            Adicionar(0).Enabled = True
            Adicionar(3).Enabled = True
            Adicionar(2).Enabled = False
            ActivaD True
            TxtCodigo.Enabled = True
            'CmdCreaArticulo.Enabled = True
            'TxtCodigoBarra = ""
            'txtArticulo.Enabled = True
            'Modificado 21/10/2005 DETP
            'TxtCodigoBarra.Enabled = True
'            If TxtCodigoBarra.Visible Then
'                TxtCodigoBarra.SetFocus
'            Else
'                txtArticulo.SetFocus
'            End If
            If ChkRet1.Visible = True Then
                ChkRet1.value = 0
                ChkRet2.value = 0
            End If
        
            TxtDescuento.Text = "0"
            TxtPDescuento.Text = "0"
            lDescontado = False
            Me.ChkDescuento.value = 0
            
        Case Is = 3 'Cancelar
            lDescontado = False
            PantallaD
'            TxtCodigoBarra.Enabled = True
'            If TxtCodigoBarra.Visible Then
'                TxtCodigoBarra.SetFocus
'            End If
            
        Case Is = 4 'Historial del artículo
'            LlenaArticulo
'            Dim X As Integer
'            If Me.Height = 4150 Then
'                If frmDocumentoCompra.TxtProveedor = "" Then
'                   ChkExclusivo.value = 0
'                   ChkExclusivo.Enabled = False
'                Else
'                   ChkExclusivo.Enabled = True
'                End If
'                ChkExclusivo.Visible = True
'                For X = 4150 To 5985
'                    Me.Height = X
'                    Centrar Me
'                Next
'                Me.Height = 5985
'            Else
'                ChkExclusivo.Visible = False
'                For X = 5985 To 4150 Step -1
'                    Me.Height = X
'                    Centrar Me
'                Next
'                Me.Height = 4150
'            End If
    End Select
End Sub

Sub ActivaD(ByVal Valor As Boolean)
    LblStockActuales.Enabled = Valor
    TxtCantidadPedida.Enabled = Valor
End Sub

Sub MuestraImpuestos()
    With frmDocumentoCompra.RsDocumentoD
        If nPorcentaje1 <= 0 Then
            ChkImp1.Visible = False
            TxtImp1.Visible = False
        End If
        If nPorcentaje2 <= 0 Then
            ChkImp2.Visible = False
            TxtImp2.Visible = False
        End If
        If nPorcentaje3 <= 0 Then
            ChkImp3.Visible = False
            TxtImp3.Visible = False
        End If
        If nPorcentaje1 > 0 Then
            ChkImp1 = IIf(!nImpuesto1 = True, 1, 0)
            TxtImp1 = IIf(IsNull(nPorcentaje1), "", nPorcentaje1)
            TxtImp1.Visible = True: ChkImp1.Visible = True
        End If
        If nPorcentaje2 > 0 Then
            ChkImp2 = IIf(!nImpuesto2 = True, 1, 0)
            TxtImp2 = IIf(IsNull(nPorcentaje2), "", nPorcentaje2)
            TxtImp2.Visible = True: ChkImp2.Visible = True
        End If
        If nPorcentaje3 > 0 Then
            ChkImp3 = IIf(!nImpuesto3 = True, 1, 0)
            TxtImp3 = IIf(IsNull(nPorcentaje3), "", nPorcentaje3)
            TxtImp3.Visible = True: ChkImp3.Visible = True
        End If
    End With
End Sub

Sub SubAsignar()

'******************************************* Asigna desde Tabla
    With frmDocumentoCompra.RsDocumentoD
        TxtCodigo = IIf(IsNull(!tCodigoProducto), "", !tCodigoProducto)
        'TxtCodigoBarra = IIf(IsNull(!CBarra), "", !CBarra)
        txtArticulo = IIf(IsNull(!Producto), "", !Producto)
        TxtCantidadPedida = IIf(IsNull(!nCantidad), "", Format(!nCantidad, "###,###,##0.000"))
        TxtTotal = IIf(IsNull(!nNeto), "", Format(!nNeto, "###,###,##0.000"))
 
        'Jesus 151113
        TxtPrecioCompra.Text = Calcular("select IsNull(nPrecioCosto,0) as Codigo from tproducto where tcodigoproducto='" & TxtCodigo & "'", CnAlmacen)
        'Fin
        
'        If nRetencion1 <> 0 Then
'            If !lIncluidoRetencion1 = True Then
'                ChkRet1.value = 1
'            End If
'        End If
        
'        If nRetencion2 <> 0 Then
'            If !lIncluidoRetencion2 = True Then
'                ChkRet2.value = 1
'            End If
'        End If
        
        Select Case pais
            Case "000", "004", "005"
                TxtTotal = IIf(IsNull(!nNeto), "", Format(!nNeto, "###,###,##0.000"))
            Case "001"
 
                If !lAsumirIUE_IT = True Then
                    Me.chkIUE_IT.value = 1
                Else
                    Me.chkIUE_IT.value = 0
                End If
                If !nDescuento <> 0 Or IsNull(IsNull(!nDescuento)) = False Then
                    TxtTotal = Format(CDbl(IIf(IsNull(!nPVenta), 0, !nPVenta)) - CDbl(IIf(IsNull(!nDescuento), 0, !nDescuento)), "###,###,##0.000")
                Else
                    TxtTotal = IIf(IsNull(!nPVenta), "", Format(!nPVenta, "###,###,##0.000"))
                End If
            Case "002"
                TxtTotal = IIf(IsNull(!nNeto), "", Format(!nNeto, "###,###,##0.000"))
        End Select
        
        dPrecioUnitario = IIf(IsNull(!nPrecio), "", !nPrecio)
        TxtPrecioUnitario = Format(dPrecioUnitario, "###,###,##0.000")
        
        nOtrosCargosInafecto = IIf(IsNull(!nOtrosCargosInafecto), 0, !nOtrosCargosInafecto)
        txtIRBP_ICE = Format(nOtrosCargosInafecto, "###,###,##0.000")
        
        LblStockActuales = IIf(IsNull(!nStockActual), "", Format(!nStockActual, "###,###,##0.000"))
        CboUnidad.BoundText = IIf(IsNull(!tUnidadEntrada), "", !tUnidadEntrada)
        txtCuenta.Text = IIf(IsNull(!TCUENTACONTABLE), "", !TCUENTACONTABLE)
        txtCC.Text = IIf(IsNull(!tCentroCosto), "", !tCentroCosto)
        
        If !nImpuesto1 <> 0 Then
            TxtImp1.Visible = True
            ChkImp1.Visible = True
            ChkImp1.value = 1
            TxtImp1 = IIf(IsNull(!nPorcentaje1), "", !nPorcentaje1)
        Else
            If !nPorcentaje1 > 0 Then
                ChkImp1.Visible = True
                TxtImp1.Visible = True
            End If
        End If
        If !nImpuesto2 <> 0 Then
            TxtImp2.Visible = True
            ChkImp2.Visible = True
            ChkImp2.value = 1
            TxtImp2 = IIf(IsNull(!nPorcentaje2), "", !nPorcentaje2)
        Else
            If !nPorcentaje2 > 0 Then
                ChkImp2.Visible = True
                TxtImp2.Visible = True
            End If
        End If
        If !nImpuesto3 <> 0 Then
            TxtImp3.Visible = True
            ChkImp3.Visible = True
            ChkImp3.value = 1
            TxtImp3 = IIf(IsNull(!nPorcentaje3), "", !nPorcentaje3)
        Else
            If !nPorcentaje3 > 0 Then
                ChkImp3.Visible = True
                TxtImp3.Visible = True
            End If
        End If
        
        '************************************************************** Ricky
        If !nDescuento <> 0 Then
            TxtDescuento = IIf(IsNull(!nDescuento), "", !nDescuento)
            TxtPDescuento = Format((!nDescuento * 100) / !nPVenta, "##0.000")
            ChkDescuento.Enabled = True
            ChkDescuento.value = 1
        Else
            TxtDescuento = "0"
            TxtPDescuento = "0"
            ChkDescuento.Enabled = True
        End If
    End With
End Sub

Sub AsignaTmp()

'******************************************* Asigna desde Temporal
    With frmDocumentoCompra.RsDtemporal
        TxtCodigo = IIf(IsNull(!tCodigoProducto), "", !tCodigoProducto)
        'TxtCodigoBarra = IIf(IsNull(!CBarra), "", !CBarra)
        txtArticulo = IIf(IsNull(!Producto), "", !Producto)
        TxtCantidadPedida = IIf(IsNull(!nCantidad), "", Format(!nCantidad, "###,###,##0.000"))
'        If nRetencion1 <> 0 Then
'            If !lIncluidoRetencion1 = True Then
'                ChkRet1.value = 1
'            End If
'        End If
        
'        If nRetencion2 <> 0 Then
'            If !lIncluidoRetencion2 = True Then
'                ChkRet2.value = 1
'            End If
'        End If

        Select Case pais
            Case "000", "004", "005"
                TxtTotal = IIf(IsNull(!nNeto), "", Format(!nNeto, "###,###,##0.000"))
'                Total_Ant = TxtTotal.Text - borrado clr
            Case "001"
                If !lAsumirIUE_IT = True Then
                    Me.chkIUE_IT.value = 1
                Else
                    Me.chkIUE_IT.value = 0
                End If
                TxtTotal = IIf(IsNull(!nNeto), "", Format(!nNeto, "###,###,##0.000")) + (!nImpuesto1 + !nImpuesto2 + !nImpuesto3)
            Case "002"
                TxtTotal = IIf(IsNull(!nNeto), "", Format(!nNeto, "###,###,##0.000"))
        End Select
        txtCuenta = IIf(IsNull(!TCUENTACONTABLE), "", !TCUENTACONTABLE)
        txtCC = IIf(IsNull(!tCentroCosto), "", !tCentroCosto)
        
        dPrecioUnitario = IIf(IsNull(!nPrecio), "", !nPrecio)
        TxtPrecioUnitario = Format(dPrecioUnitario, "###,###,##0.000")
        
        nOtrosCargosInafecto = IIf(IsNull(!nOtrosCargosInafecto), "", !nOtrosCargosInafecto)
        txtIRBP_ICE = Format(nOtrosCargosInafecto, "###,###,##0.000")
        
        LblStockActuales = IIf(IsNull(!nStockActual), "", Format(!nStockActual, "###,###,##0.000"))
        CboUnidad.BoundText = IIf(IsNull(!tUnidadEntrada), "", !tUnidadEntrada)
        
        If nPorcentaje1 <> 0 Then
            TxtImp1.Visible = True
            ChkImp1.Visible = True
            TxtImp1.Text = IIf(IsNull(!nPorcentaje1), "", !nPorcentaje1)
            If !lIncluidoImpuesto1 = True Then
                ChkImp1.value = 1
            End If
        Else
            ChkImp1.Visible = False
            TxtImp1.Visible = False
        End If
    
        If nPorcentaje2 <> 0 Then
            TxtImp2.Visible = True
            ChkImp2.Visible = True
            TxtImp2.Text = IIf(IsNull(!nPorcentaje2), "", !nPorcentaje2)
            If !lIncluidoImpuesto2 = True Then
                ChkImp2.value = 1
            End If
        Else
            ChkImp2.Visible = False
            TxtImp2.Visible = False
        End If
    
        If nPorcentaje3 <> 0 Then
            TxtImp3.Visible = True
            ChkImp3.Visible = True
            TxtImp3.Text = IIf(IsNull(!nPorcentaje3), "", !nPorcentaje3)
            If !lIncluidoImpuesto3 = True Then
                ChkImp3.value = 1
            End If
        Else
            ChkImp3.Visible = False
            TxtImp3.Visible = False
        End If
    
        '************************************************************** Ricky
        If !nDescuento <> 0 Then
            TxtDescuento = IIf(IsNull(!nDescuento), "", !nDescuento)
            TxtPDescuento = Format((!nDescuento * 100) / !nPVenta, "##0.000")
            ChkDescuento.Enabled = True
            ChkDescuento.value = 1
        Else
            TxtDescuento = "0"
            TxtPDescuento = "0"
            ChkDescuento.Enabled = True
        End If
    End With
End Sub

Private Sub ChkDescuento_Click()
    
    If ChkDescuento.value = 1 Then
        If val(TxtTotal.Text) = 0 Then ChkDescuento.value = 0: Exit Sub
        TxtDescuento.Enabled = True
        CmdPDescuento.Enabled = True
        CmdMDescuento.Enabled = True
        If TxtDescuento.Text > 0 Then
            lDescontado = True
        End If
    Else
        TxtDescuento.Enabled = False
        CmdPDescuento.Enabled = False
        CmdMDescuento.Enabled = False
        lDescontado = False
        TxtTotal = 0
    End If
      
    If SwTmp Then
        If lAgregaItem Then
            If ChkDescuento.value = 0 Then
                If frmDocumentoCompra.grdGrilla.VisibleRows > 0 Then
                    sPVen = frmDocumentoCompra.RsDtemporal!nPVenta
                    If TxtTotal < CDbl(sPVen) And TxtTotal > 0 Then
                        TxtTotal = Format(CDbl(TxtTotal) + CDbl(TxtDescuento), "###,###,##0.000")
                    End If
                Else
                    If TxtTotal < sPVen And TxtTotal > "0" Then
                        TxtTotal = Format(CDbl(TxtTotal) + CDbl(TxtDescuento), "###,###,##0.000")
                    End If
                End If
                TxtDescuento = "0"
                TxtPDescuento = "0"
                calculo
            Else
                If TxtDescuento = "0" Then
                    TxtDescuento.Enabled = True
                    TxtDescuento.SetFocus
                    TxtDescuento.SelStart = 0
                    TxtDescuento.SelLength = Len(TxtDescuento)
                    lDesc = "I"
                End If
            End If
        Else
            If ChkDescuento.value = 0 Then
                If frmDocumentoCompra.grdGrilla.VisibleRows > 0 Then
                    sPVen = frmDocumentoCompra.RsDtemporal!nPVenta
                    If TxtTotal < CDbl(sPVen) And TxtTotal > 0 Then
                        TxtTotal = Format(CDbl(TxtTotal) + CDbl(TxtDescuento), "###,###,##0.000")
                    End If
                Else
                    If TxtTotal < sPVen And TxtTotal > 0 Then
                        TxtTotal = Format(CDbl(TxtTotal) + CDbl(TxtDescuento), "###,###,##0.000")
                    End If
                End If
                TxtDescuento = "0"
                TxtPDescuento = "0"
                calculo
            Else
                If TxtDescuento = "0" Then
                    TxtDescuento.Enabled = True
                    TxtDescuento.SetFocus
                    TxtDescuento.SelStart = 0
                    TxtDescuento.SelLength = Len(TxtDescuento)
                    lDesc = "I"
                End If
            End If
        End If
    Else
        If lAgregaItem Then
            If ChkDescuento.value = 0 Then
                If Trim(TxtTotal) = "" Then TxtTotal = 0
                If frmDocumentoCompra.grdGrilla.VisibleRows > 0 Then
                    sPVen = frmDocumentoCompra.RsDocumentoD!nPVenta
                    If TxtTotal < sPVen And TxtTotal > 0 Then
                        TxtTotal = Format(TxtTotal + TxtDescuento, "###,###,##0.000")
                    End If
                Else
                    If TxtTotal < sPVen And TxtTotal > 0 Then
                        TxtTotal = Format(CDbl(TxtTotal) + CDbl(TxtDescuento), "###,###,##0.000")
                        'TxtTotal = CDbl(TxtTotal) + CDbl(TxtDescuento)
                    End If
                End If
                TxtDescuento = "0"
                TxtPDescuento = "0"
                calculo
                'lDesc = ""
            Else
                If TxtDescuento = "0" Then
                    TxtDescuento.Enabled = True
                    TxtDescuento.SetFocus
                    TxtDescuento.SelStart = 0
                    TxtDescuento.SelLength = Len(TxtDescuento)
                    lDesc = "I"
                End If
            End If
        Else
            If ChkDescuento.value = 0 Then
                If frmDocumentoCompra.grdGrilla.VisibleRows > 0 Then
                    sPVen = frmDocumentoCompra.RsDocumentoD!nPVenta
                    If TxtTotal < CDbl(sPVen) And TxtTotal > 0 Then
                        TxtTotal = Format(CDbl(TxtTotal) + CDbl(TxtDescuento), "###,###,##0.000")
                    End If
                Else
                    If TxtTotal < sPVen And TxtTotal > 0 Then
                        TxtTotal = Format(CDbl(TxtTotal) + CDbl(TxtDescuento), "###,###,##0.000")
                    End If
                End If
                TxtDescuento = "0"
                TxtPDescuento = "0"
                calculo
            Else
                If TxtDescuento = "0" Then
                    TxtDescuento.Enabled = True
                    TxtDescuento.SetFocus
                    TxtDescuento.SelStart = 0
                    TxtDescuento.SelLength = Len(TxtDescuento)
                    lDesc = "I"
                End If
            End If
         End If
      End If
End Sub

Private Sub ChkExclusivo_Click()
    LlenaArticulo
End Sub

Private Sub ChkImp1_Click()
    Select Case pais
        Case "001"
            TxtImp1.Enabled = Not TxtImp1.Enabled
            calculo
        Case "000", "002", "004", "005"
            TxtImp1.Enabled = Not TxtImp1.Enabled
    End Select
End Sub

Private Sub ChkImp2_Click()
    Select Case pais
        Case "000", "004", "005"
            TxtImp2.Enabled = Not TxtImp2.Enabled
        Case "001"
            TxtImp2.Enabled = Not TxtImp2.Enabled
            calculo
        Case "002"
            TxtImp2.Enabled = Not TxtImp2.Enabled
    End Select
End Sub

Private Sub chkIUE_IT_Click()
    If TxtTotal.Text = "" Then TxtTotal.Text = 0
    If ChkRet1.value = 1 And ChkRet2.value = 1 Then
        If chkIUE_IT.value = 1 Then
            TxtTotal.Text = Format((CDbl(TxtTotal.Text) * 100) / (100 - (CDbl(TxtRet1.Text) + CDbl(TxtRet2.Text))), "###,###,##0.000")
                calculo
        Else
             TxtTotal.Text = Format((CDbl(TxtTotal.Text) * (100 - (CDbl(TxtRet1.Text) + CDbl(TxtRet2.Text))) / 100), "###,###,##0.000")
            calculo
        End If
    Else
          chkIUE_IT.value = False
    End If
End Sub

Private Sub ChkRet1_Click()
            TxtRet1.Enabled = Not TxtRet1.Enabled
End Sub

Private Sub ChkRet2_Click()
            TxtRet2.Enabled = Not TxtRet2.Enabled
End Sub

Private Sub ChkRet3_Click()
    TxtRet3.Enabled = Not TxtRet3.Enabled
End Sub

Private Sub ChkImp3_Click()
    Select Case pais
        Case "000", "004", "005"
            TxtImp3.Enabled = Not TxtImp3.Enabled
        Case "001"
            TxtImp3.Enabled = Not TxtImp3.Enabled
            calculo
        Case "002"
            TxtImp3.Enabled = Not TxtImp3.Enabled
    End Select
End Sub

Private Sub chkRecargo_Click()
    TxtRecargo.Enabled = Not TxtRecargo.Enabled
    CmdPRecargo.Enabled = Not CmdPRecargo.Enabled
    CmdMRecargo.Enabled = Not CmdMRecargo.Enabled
    ChkDescuento.Enabled = Not ChkDescuento.Enabled
    If chkRecargo.value = 0 And TxtTotal <= frmDocumentoCompra.RsDtemporal!nPVenta Then
        TxtTotal = Format(CDbl(TxtTotal) - CDbl(TxtRecargo), "#,###,##0.000")
        calculo
    End If
End Sub

Private Sub cmdFind_Click()
    Isql = "sp_Ayuda_Producto '" & frmDocumentoCompra.CboTipoIngreso.BoundText & "'"

   ' REVISAR
    'sBuscar = "Articulo"
    frmBusca.lEntra = True
    frmBusca.nPredeterm = 1
    frmBusca.TipoOperacion = "AlmacenArticulo"
    frmBusca.txtCriterio = Trim(UCase(txtArticulo))
    frmBusca.txtCriterio.SelStart = Len(Trim(frmBusca.txtCriterio))
    lCriterio = False

    Call ConfGrilla(3, frmBusca.grdGrilla, "codigo", 2, "Codigo", 1000, 0, 0, "", _
                        "Descripcion", 2, "Descripcion", 4100, 0, 0, "", _
                        "UnidadMedida", 2, "UnidadMedida", 2400, 0, 0, "")

    frmBusca.Show vbModal, FrmSeleccionProductos
    Adicionar(4).Enabled = True
    VerificaFVence
    If lContabilidad And VerificaTipoIngreso(frmDocumentoCompra.CboTipoIngreso.BoundText) Then
        txtCuenta.SetFocus
    Else
        TxtCantidadPedida.SetFocus
    End If
    If TxtCodigo.Text <> "" Then
        If Calcular("Select lCompraDocumento as Codigo from tParametro", CnAlmacen) = True Then
            VentanaCompra
        End If
    End If
    Total_Ant = 0
    PrecioUnitario_Ant = 0
End Sub

Private Sub CmdMDescuento_Click()
    If lDescontado = True Then Exit Sub
    TxtPDescuento = Format((TxtDescuento.Text * 100) / TxtTotal, "##0.000")
    TxtTotal = TxtTotal - TxtDescuento
    calculo
    lDescontado = True
End Sub

Private Sub CmdMRecargo_Click()
    TxtPRecargo = Format((TxtRecargo.Text * 100) / TxtTotal, "##0.000")
    calculo
End Sub

Private Sub CmdPDescuento_Click()
    If lDescontado = True Then Exit Sub
    TxtPDescuento = Format(TxtDescuento.Text, "##0.000")
    TxtDescuento = Format(TxtTotal * (TxtDescuento.Text / 100), "##0.000")
    TxtTotal = TxtTotal - TxtDescuento
    calculo
    lDescontado = True
End Sub

Private Sub CmdPRecargo_Click()
    TxtPRecargo = Format(TxtRecargo.Text, "##0.000")
    TxtRecargo = Format(TxtTotal * (TxtRecargo.Text / 100), "##0.000")
    TxtTotal = TxtTotal + TxtRecargo
    calculo
End Sub

Private Sub CmdCreaArticulo_Click()
    'Set RsParametro = New ADODB.Recordset
    'Jesus 15/03/2017
    'If PermisoCentralizado("10104000") Then Exit Sub
    'Fin 15/03/2017
    
'    If AccesoFormulario("FrmArticulo") Then
'        CreaArticulo "Artículo", "", txtArticulo
'        If CboUnidad.Enabled Then CboUnidad.BoundText = sCodigo
'    Else
'        MsgBox "Usuario sin Acceso", vbExclamation, sMensaje
'    End If
End Sub

Sub LlenaArticulo()
    If Calcular("Select lCompraDocumento as Codigo from tParametro", CnAlmacen) = False Then Exit Sub
    If frmDocumentoCompra.TxtProveedor = "" Or ChkExclusivo.value = 0 Then
        Isql = "sp_LlenaArticuloDocumento '" & TxtCodigo & "', '', '1'"
    ElseIf ChkExclusivo.value = 1 Then
        Isql = "sp_LlenaArticuloDocumento '" & TxtCodigo & "', '" & frmDocumentoCompra.TxtProveedor & "', '2'"
    End If
    Set RsProveedorPrecio = Lib.OpenRecordset(Isql, CnAlmacen)
    Set grdGrilla.DataSource = RsProveedorPrecio
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    'lAgregaItem = True
    PantallaD
    
    'Jesus 15/11/06
    lPrecioCompra = Calcular("select IsNull(lPrecioCompra,0) as Codigo from vTipoDocumento where Codigo='" & frmDocumentoCompra.sxTipoDocumento & "'", CnAlmacen)
    'Fin
    
    Select Case pais
        Case "000", "004", "002", "003", "005"
            Me.Height = 4150
            Me.Width = 7185
            Adicionar(1).Left = 5990
            grdGrilla.Width = 7065
            
            Centrar Me
    
            'Deshabilita Impuestos
            
            txtIRBP_ICE.Visible = True
            LPrecioUnitario(2).Visible = True
            LblMoneda(4).Visible = True
            
            TxtImp1.Visible = False
            TxtImp2.Visible = False
            TxtImp3.Visible = False
            ChkImp1.Visible = False
            ChkImp2.Visible = False
            ChkImp3.Visible = False
            chkIUE_IT.Visible = False
            
            'Muestra Impuestos
             
            ChkImp1.ToolTipText = "'" & Trim(sImpuesto1) & "' afecto o no al valor neto"
            ChkImp2.ToolTipText = "'" & Trim(sImpuesto2) & "' afecto o no al valor neto"
            ChkImp3.ToolTipText = "'" & Trim(sImpuesto3) & "' afecto o no al valor neto"
            ChkDescuento.ToolTipText = "Descuento por aplicar/aplicado al valor neto"
             
            ChkImp1.Caption = sImpuesto1
            ChkImp2.Caption = sImpuesto2
            ChkImp3.Caption = sImpuesto3
            
            'Muestra si estan los Porcentajes
             
            TxtImp1 = nPorcentaje1
            TxtImp2 = nPorcentaje2
            TxtImp3 = nPorcentaje3
            
            If nPorcentaje1 = 0 Then
                ChkImp1.Visible = False
                TxtImp1.Visible = False
                LblImp1.Visible = False
            Else
                ChkImp1.Visible = True
                TxtImp1.Visible = True
                LblImp1.Visible = True
            End If
             
            If nPorcentaje2 = 0 Then
                ChkImp2.Visible = False
                TxtImp2.Visible = False
                LblImp2.Visible = False
            Else
                ChkImp2.Visible = True
                TxtImp2.Visible = True
                LblImp2.Visible = True
            End If
             
            If nPorcentaje3 = 0 Then
                ChkImp3.Visible = False
                TxtImp3.Visible = False
                LblImp3.Visible = False
            Else
                ChkImp3.Visible = True
                TxtImp3.Visible = True
                LblImp3.Visible = True
            End If
            
            ChkRet1.Visible = False
            TxtRet1.Visible = False
            LblRet1.Visible = False
            ChkRet2.Visible = False
            TxtRet2.Visible = False
            LblRet2.Visible = False
            ChkRet3.Visible = False
            TxtRet3.Visible = False
            LblRet3.Visible = False
            
            If pais = "005" Then Me.TxtImp1.Locked = False
            
        Case "001"
            Me.Height = 4150
            Me.Width = 8520
'            ChkICE.Visible = True
'            TxtICE.Visible = True
'            LblMoneda(2).Visible = True
'            chkIUE_IT.Visible = True
            Adicionar(1).Left = 7275
            grdGrilla.Width = 8355
            
            Centrar Me
    
            txtIRBP_ICE.Visible = False
            LPrecioUnitario(2).Visible = False
            LblMoneda(4).Visible = False
            txtIRBP_ICE.Text = "0.000"
            'Deshabilita Impuestos
             
            TxtImp1.Visible = False
            TxtImp2.Visible = False
            TxtImp3.Visible = False
            ChkImp1.Visible = False
            ChkImp2.Visible = False
            ChkImp3.Visible = False
            chkIUE_IT.Visible = True
            
            'Muestra Impuestos
             
            ChkImp1.ToolTipText = "'" & Trim(sImpuesto1) & "' afecto o no al valor neto"
            ChkImp2.ToolTipText = "'" & Trim(sImpuesto2) & "' afecto o no al valor neto"
            ChkImp3.ToolTipText = "'" & Trim(sImpuesto3) & "' afecto o no al valor neto"
            ChkDescuento.ToolTipText = "Descuento por aplicar/aplicado al valor neto"
             
            ChkImp1.Caption = sImpuesto1
            ChkImp2.Caption = sImpuesto2
            ChkImp3.Caption = sImpuesto3
            
            'Muestra si estan los Porcentajes
             
            TxtImp1 = nPorcentaje1
            TxtImp2 = nPorcentaje2
            TxtImp3 = nPorcentaje3
                
'            ChkRet1.Caption = sRetencion1
'            ChkRet2.Caption = sRetencion2
'            ChkRet3.Caption = sRetencion3
'            TxtRet1 = nRetencion1
'            TxtRet2 = nRetencion2
'            TxtRet3 = nRetencion3
            
            If nPorcentaje1 = 0 Then
                ChkImp1.Visible = False
                TxtImp1.Visible = False
                LblImp1.Visible = False
            Else
                ChkImp1.Visible = True
                TxtImp1.Visible = True
                LblImp1.Visible = True
            End If
             
            If nPorcentaje2 = 0 Then
                ChkImp2.Visible = False
                TxtImp2.Visible = False
                LblImp2.Visible = False
            Else
                ChkImp2.Visible = True
                TxtImp2.Visible = True
                LblImp2.Visible = True
            End If
             
            If nPorcentaje3 = 0 Then
                ChkImp3.Visible = False
                TxtImp3.Visible = False
                LblImp3.Visible = False
            Else
                ChkImp3.Visible = True
                TxtImp3.Visible = True
                LblImp3.Visible = True
            End If
            
'            If nRetencion1 = 0 Then
'                ChkRet1.Visible = False
'                TxtRet1.Visible = False
'                LblRet1.Visible = False
'            Else
'                ChkRet1.Visible = True
'                TxtRet1.Visible = True
'                LblRet1.Visible = True
'            End If
'
'            If nRetencion2 = 0 Then
'                ChkRet2.Visible = False
'                TxtRet2.Visible = False
'                LblRet2.Visible = False
'            Else
'                ChkRet2.Visible = True
'                TxtRet2.Visible = True
'                LblRet2.Visible = True
'            End If
'
'            If nRetencion3 = 0 Then
'                ChkRet3.Visible = False
'                TxtRet3.Visible = False
'                LblRet3.Visible = False
'            Else
'                ChkRet3.Visible = True
'                TxtRet3.Visible = True
'                LblRet3.Visible = True
'            End If
'
    End Select
    
    'Linea Agregada por Luis
    lCodigoBarra = Calcular("select top 1 lCodigoBarra as Codigo from tParametro", CnAlmacen)
    'TxtCodigoBarra.Visible = lCodigoBarra
    'Label3.Visible = lCodigoBarra
     
    lContabilidad = Calcular("select top 1 lContabilidad as codigo from tParametro", CnAlmacen)
    'Nuevo
'    If lContabilidad And VerificaTipoIngreso(frmDocumentoCompra.CboTipoIngreso.BoundText) Then
'        fraCuenta.Visible = True
'    End If
    TxtDescuento = 0

'    ChkDescuento.Enabled = False
    SwTmp = Sw
    If Not lAgregaItem Then
        If SwTmp Then              ' en Temporal
            If frmDocumentoCompra.RsDtemporal.RecordCount > 0 Then         'Modifica item nuevo
                iRow = frmDocumentoCompra.RsDtemporal.AbsolutePosition
                frmDocumentoCompra.RsDtemporal.Requery
                If frmDocumentoCompra.RsDtemporal.RecordCount > 0 Then frmDocumentoCompra.RsDtemporal.AbsolutePosition = IIf(iRow = -1, 1, iRow)
                If frmDocumentoCompra.RsDtemporal!nDescuento > 0 Then      'Tiene descuento
                    'ChkDescuento.Value = 1
                    If lDesc = "G" Then
                        ChkDescuento.Enabled = False
                    Else
                        ChkDescuento.Enabled = True
                        TxtDescuento = frmDocumentoCompra.RsDtemporal!nDescuento
                        TxtPDescuento = frmDocumentoCompra.RsDtemporal!nDescuento * 100 / frmDocumentoCompra.RsDtemporal!nPVenta
                        ChkDescuento.value = 1
                    End If
                Else
                    TxtDescuento = 0
                    ChkDescuento.Enabled = True
                End If
            Else
                ChkDescuento.Enabled = True
            End If
         Else
            If frmDocumentoCompra.RsDocumentoD.RecordCount > 0 Then        'Modifica item antiguo
                iRow = frmDocumentoCompra.RsDocumentoD.AbsolutePosition
                frmDocumentoCompra.RsDocumentoD.Requery
                If frmDocumentoCompra.RsDocumentoD.RecordCount > 0 Then frmDocumentoCompra.RsDocumentoD.AbsolutePosition = IIf(iRow = -1, 1, iRow)
                If frmDocumentoCompra.RsDocumentoD!nDescuento > 0 Then     'Tiene descuento
                    If lDesc = "G" Then
                        ChkDescuento.Enabled = False
                    Else
                        ChkDescuento.Enabled = True
                        If lDesc = "" Then
                            TxtDescuento = 0
                            TxtPDescuento = 0
                        Else
                            TxtDescuento = frmDocumentoCompra.RsDocumentoD!nDescuento
                            TxtPDescuento = frmDocumentoCompra.RsDocumentoD!nDescuento * 100 / frmDocumentoCompra.RsDocumentoD!nPVenta
                            ChkDescuento.value = 1
                        End If
                    End If
                Else
                    TxtDescuento = 0
                    ChkDescuento.Enabled = True
                End If
            Else
                ChkDescuento.Enabled = True
            End If
         End If
     Else
        lDescontado = False
        ChkDescuento.value = 0
     End If
     
     Flagloctem = IIf(Calcular("select count(*) as codigo from trelacionOrdenCompraDocumento where tcorrelativodoc='" & frmDocumentoCompra.LblCorrelativo.Caption & "'", CnAlmacen) > 0, True, False)
     'Nuevo
    'PantallaD   'Borra la Pantalla
     Llena_Combos
     LlenaArticulo
     
     If lDesc = "G" Then
        ChkDescuento.Enabled = False
     Else
        ChkDescuento.Enabled = True
     End If
     If Trim(txtArticulo) = "" Then
        Adicionar(4).Enabled = False
     Else
        Adicionar(4).Enabled = True
     End If
     
     ConfGrilla 5, grdGrilla, "Proveedor", 2, "tRazonComercial", 3200, 0, 0, "", _
                              "Mon.", 2, "ValorMoneda", 500, 2, 0, "", _
                              "Precio", 1, "nPrecio", 800, 1, 0, "#,##0.000", _
                              "Cantidad", 1, "nCantidad", 800, 1, 0, "#,##0.000", _
                              "Fecha", 2, "FEmision", 1200, 2, 0, "dd/mm/yyyy"
     VerificaFVence
     Screen.MousePointer = vbDefault
     swPunto = False
     
     'If ErpOF_Activo Or ErpSG_Activo Or ErpSB_Activo Then CmdCreaArticulo.Enabled = False
End Sub

Sub Llena_Combos()
    Isql = "Select tUnidadEntrada As Codigo,UnidadMedida As Descripcion From vComboProductoImpuestos Order By 2"
    Llena_Combo Isql, CboUnidad, RsUnidad, "Codigo", "Descripcion"
End Sub

Sub Liberar()
    Set rsDes = Nothing
    Set RsMedida = Nothing
    Set RsUnidad = Nothing
    Set RsBusca = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Liberar
    Isql = ""
End Sub

Private Sub TxtArticulo_Change()
    If swEntrar Then
        If Len(Trim(UCase(txtArticulo.Text))) = 1 Then
            Isql = "sp_Ayuda_Producto '" & frmDocumentoCompra.CboTipoIngreso.BoundText & "'"

            'sBuscar = "Articulo"
            frmBusca.lEntra = True
            frmBusca.txtCriterio = Trim(UCase(txtArticulo))
            frmBusca.txtCriterio.SelStart = Len(Trim(frmBusca.txtCriterio))
            lCriterio = False
            'lDondeAgrega = False
            frmBusca.Show vbModal, FrmSeleccionProductos
            Adicionar(4).Enabled = True
            VerificaFVence
            If lContabilidad And VerificaTipoIngreso(frmDocumentoCompra.CboTipoIngreso.BoundText) Then
                txtCuenta.SetFocus
            Else
                TxtCantidadPedida.SetFocus
            End If
            If TxtCodigo.Text <> "" Then
                If Calcular("Select lCompraDocumento as Codigo from tParametro", CnAlmacen) = True Then
                    VentanaCompra
                End If
            End If
        End If
    End If
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If lContabilidad Then
          txtCuenta.SetFocus
       Else
          TxtCantidadPedida.SetFocus
       End If
    End If
End Sub

Private Sub TxtCantidadPedida_GotFocus()
    TxtCantidadPedida.SelStart = 0
    TxtCantidadPedida.SelLength = Len(TxtCantidadPedida)
    swPunto = IIf(TxtCantidadPedida.Text = "", False, True)
End Sub

Private Sub TxtCantidadPedida_KeyPress(KeyAscii As Integer)
    swPunto = ChrBuscaPunto(TxtCantidadPedida.Text) 'If Trim(TxtCantidad) = "" Then swPunto = False
    If Trim(TxtCantidadPedida) = "" Then swPunto = False
    If KeyAscii = 13 Then
       TxtTotal.SetFocus
    Else
       PictureNumero KeyAscii      'Para evaluar si lo digitado es un numero
    End If
End Sub

Private Sub TxtCantidadPedida_LostFocus()
    
    If lPrecioCompra Then
        TxtTotal = Format(IIf(TxtCantidadPedida = "", 0, TxtCantidadPedida) * IIf(TxtPrecioCompra = "", 0, TxtPrecioCompra), "###,###,##0.000")
    End If
    
    If Trim(TxtCantidadPedida) <> "" Then
       If TxtTotal.Text <> "" Then
          calculo
       End If
       TxtCantidadPedida = Format(TxtCantidadPedida, "###,###,##0.000")
       TxtTotal = Format(TxtTotal, "###,###,##0.000")
       Prc_Mensaje
    Else
       Adicionar(0).Enabled = True
    End If
    swPunto = False
End Sub

Sub Prc_Mensaje()
    If IsNumeric(TxtCantidadPedida) Then
        If Trim(TxtCantidadPedida) = "" Then TxtCantidadPedida = 0
        If Trim(CDbl(TxtCantidadPedida)) = 0 Then
           MsgBox "Ingrese la cantidad...", vbInformation, sMensaje
           TxtCantidadPedida = ""
           TxtTotal = ""
        End If
    End If
End Sub

Private Sub txtCC_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       PictureNumero KeyAscii
    Else
       TxtCantidadPedida.SetFocus
    End If
End Sub

Private Sub TxtCodigo_Change()
    If swEntrar Then
        If Len(Trim(TxtCodigo)) Then
           frmBusca.lEntra = True
        Else
           Exit Sub
        End If
        If frmBusca.lEntra Then
           If Len(Trim(UCase(TxtCodigo))) = 1 Then
              Isql = "Select Distinct tCodigoProducto As Codigo,Producto As Descripcion,tCodigoBarra as CBarra,tUnidadEntrada,UnidadMedida,nPrecioCosto,nStockActual,nImp1,nImp2,nImp3,tTipoProducto From vComboProductoImpuestos Where tTipoProducto='" & frmDocumentoCompra.CboTipoIngreso.BoundText & "' Order By Producto"
              'sBuscar = "CodigoI"
              frmBusca.lEntra = True
              frmBusca.txtCriterio = Trim(UCase(TxtCodigo))
              frmBusca.txtCriterio.SelStart = Len(Trim(frmBusca.txtCriterio))
              lCriterio = False
              frmBusca.Show vbModal, FrmSeleccionProductos
              TxtCantidadPedida.SetFocus
           End If
        End If
        LlenaArticulo
    End If
End Sub

'Private Sub TxtCodigoBarra_Change()
'    If Trim(TxtCodigoBarra) = "" Then
'       LblCaptura.Visible = True
'    Else
'       LblCaptura.Visible = False
'    End If
'    If swEntrar Then
'        If Len(Trim(TxtCodigoBarra)) Then
'           frmBusca.lEntra = True
'        Else
'           Exit Sub
'        End If
'        If frmBusca.lEntra Then
'           If Len(Trim(UCase(TxtCodigoBarra))) = 1 Then
'              Isql = "Select Distinct tCodigoProducto As Codigo,Producto As Descripcion,tCodigoBarra as CBarra,tUnidadEntrada,UnidadMedida,nPrecioCosto,nStockActual,nImp1,nImp2,nImp3,tTipoProducto From vComboProductoImpuestos Where tTipoProducto='" & frmDocumentoCompra.CboTipoIngreso.BoundText & "' Order By Producto"
'              'sBuscar = "CodigoB"
'              frmBusca.lEntra = True
'              frmBusca.txtCriterio = Trim(UCase(TxtCodigoBarra))
'              frmBusca.txtCriterio.SelStart = Len(Trim(frmBusca.txtCriterio))
'              lCriterio = False
'              frmBusca.Show vbModal, FrmSeleccionProductos
'              TxtCantidadPedida.SetFocus
'           End If
'        End If
'        LlenaArticulo
'    End If
'End Sub

'Private Sub TxtCodigoBarra_GotFocus()
'    If Trim(TxtCodigoBarra) = "" Then
'       LblCaptura.Visible = True
'    Else
'       LblCaptura.Visible = False
'    End If
'End Sub

'Private Sub TxtCodigoBarra_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       If Trim(TxtCodigoBarra) = "" Then
'          SendKeys "{TAB}"
'          Exit Sub
'       End If
'       If Trim(txtArticulo) = "" Then
'          VerificaCBarra
'       End If
'       If Trim(TxtCantidadPedida) = "" Then
'          If Trim(txtArticulo) = "" Then
'             'txtArticulo.SetFocus
'          Else
'             TxtCantidadPedida.SetFocus
'          End If
'       Else
'          TxtCantidadPedida.SetFocus
'       End If
'    End If
'End Sub
'
'Private Sub TxtCodigoBarra_LostFocus()
'    LblCaptura.Visible = False
'End Sub

Private Sub TxtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       PictureNumero KeyAscii
    Else
       txtCC.SetFocus
    End If
End Sub


Private Sub TxtDescuento_LostFocus()
    If Trim(TxtDescuento) = "" Then TxtDescuento = 0
    If TxtDescuento < 0 Then TxtDescuento = 0
    swPunto = False
End Sub

Private Sub TxtICE_GotFocus()
    TxtICE.SelStart = 0
    TxtICE.SelLength = Len(TxtICE)
End Sub

Private Sub txtICE_KeyPress(KeyAscii As Integer)
    If Trim(TxtICE) = "" Then swPunto = False
    If KeyAscii <> 13 Then PictureNumero KeyAscii
    If KeyAscii = 13 Then TxtCantidadPedida.SetFocus
End Sub

'Private Sub TxtImp1_KeyPress(KeyAscii As Integer)
'    'If Trim(TxtImp1) = "" Then swPunto = False
'    'If KeyAscii <> 13 Then PictureNumero KeyAscii
'End Sub

Private Sub TxtImp2_KeyPress(KeyAscii As Integer)
    If Trim(TxtImp2) = "" Then swPunto = False
    If KeyAscii <> 13 Then PictureNumero KeyAscii
End Sub

Private Sub TxtImp3_KeyPress(KeyAscii As Integer)
    If Trim(TxtImp3) = "" Then swPunto = False
    If KeyAscii <> 13 Then PictureNumero KeyAscii
End Sub

Private Sub TxtDescuento_KeyPress(KeyAscii As Integer)
    If Trim(TxtDescuento) = "" Then swPunto = False
    If KeyAscii <> 13 Then PictureNumero KeyAscii
End Sub

Private Sub txtIRBP_ICE_GotFocus()
    If txtIRBP_ICE.Text = "0" Or txtIRBP_ICE.Text = "0.000" Then
       txtIRBP_ICE.Text = ""
    End If
    txtIRBP_ICE.SelStart = 0
    txtIRBP_ICE.SelLength = Len(txtIRBP_ICE)
    swPunto = IIf(txtIRBP_ICE.Text = "", False, True)
End Sub

Private Sub txtIRBP_ICE_KeyPress(KeyAscii As Integer)
    swPunto = ChrBuscaPunto(txtIRBP_ICE.Text)    'If Trim(txtPrecioUnitario) = "" Then swPunto = False
    If KeyAscii = 13 Then
       Adicionar(0).SetFocus
    Else
       PictureNumero KeyAscii
    End If
End Sub

Private Sub txtIRBP_ICE_LostFocus()
    If TxtTotal.Text = "" Then Exit Sub
'    If txtPrecioNeto.Text = "" Then Exit Sub
'    If val(TxtCantidadPedida) = 0 Then
'        TxtPrecioUnitario.Text = "0.000"
'    Else
''        nPrecio = CDbl((val(txtIRBP_ICE.Text) + val(txtPrecioNeto.Text)) / val(txtCantidad.Text))
'         nPrecio = CDbl((val(txtIRBP_ICE.Text) + val(nSubTotal) - nOtrosCargosInafecto) / val(txtCantidad.Text))
'    End If
    
'    txtPrecioNeto.Text = Format(CDbl(val(txtIRBP_ICE.Text) + val(nSubTotal) - nOtrosCargosInafecto), "###,###0.000")
'
'    nTotal = Format(CDbl(val(txtIRBP_ICE.Text) + val(nTotal) - nOtrosCargosInafecto), "###,###0.000")
'
'    txtIRBP_ICE.Text = Format(txtIRBP_ICE.Text, "###,###0.000")
''    txtPrecioUnitario.Text
'    TxtPrecioUnitario.Text = Format(nPrecio, "###,###0.000")
'    txtPrecioNeto.Text = Format(val(txtIRBP_ICE.Text) + val(nSubTotal) - nOtrosCargosInafecto, "###,###0.000")
'    lblTotal.Caption = Format(nTotal, "###,###0.000")
'

    
    calculo
    nOtrosCargosInafecto = CDbl(IIf(IsNumeric(txtIRBP_ICE.Text), txtIRBP_ICE.Text, 0))
End Sub

Private Sub TxtPrecioUnitario_Change()
    If CDbl(nVPrecio) > 0 Then
       Dim dUPrecio As Double
       Dim dValor   As Double
       Dim dMinimo  As Double
       Dim dMaximo  As Double
       dUPrecio = Calcular("Select nPrecioCosto As Codigo From tProducto Where tCodigoProducto='" & TxtCodigo & "'", CnAlmacen)
       If dUPrecio > 0 Then
          dValor = dUPrecio * CDbl(nVPrecio) / 100
          dMinimo = dUPrecio - dValor
          dMaximo = dUPrecio * (1 + (CDbl(nVPrecio) / 100))
          If Trim(TxtPrecioUnitario) = "" Then TxtPrecioUnitario = 0
          'If CDbl(TxtPrecioUnitario) > dMaximo Then
          If dPrecioUnitario > dMaximo Then
             MsgBox "El precio unitario obtenido está por encima" & vbCrLf & "del margen considerado como fluctuación" & vbCrLf & "en relación al precio de compra actual...", vbInformation, sMensaje
          End If
          'If CDbl(TxtPrecioUnitario) < dMinimo Then
          If dPrecioUnitario < dMinimo Then
             MsgBox "El precio unitario obtenido está por debajo" & vbCrLf & "del margen considerado como fluctuación" & vbCrLf & "en relación al precio de compra actual...", vbInformation, sMensaje
          End If
       End If
    End If
End Sub

Private Sub TxtTotal_GotFocus()
    If TxtTotal.Text = "0" Or TxtTotal.Text = "0.000" Then
       TxtTotal.Text = ""
    End If
    TxtTotal.SelStart = 0
    TxtTotal.SelLength = Len(TxtTotal)
    swPunto = IIf(TxtTotal.Text = "", False, True)
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
    'If Trim(TxtTotal) = "" Then swPunto = False
    swPunto = ChrBuscaPunto(TxtTotal.Text)
    If KeyAscii = 13 Then
       If Trim(TxtTotal) = "" Then TxtTotal = "0"
       If Trim(TxtDescuento) = "" Then TxtDescuento = "0"
       If Trim(TxtRecargo) = "" Then TxtRecargo = "0"
       
       If Trim(sPVen) = "" Then sPVen = CDbl(TxtTotal) + CDbl(TxtDescuento) - CDbl(TxtRecargo)     'Ricky
       Adicionar(0).Enabled = True
       Adicionar(0).SetFocus
    Else
       PictureNumero KeyAscii
    End If
End Sub

Private Sub TxtTotal_LostFocus()
    If Trim(TxtTotal) = "" Then TxtTotal = 0
    If CDbl(TxtTotal) <= 0 Then Exit Sub
    If Trim(sPVen) = "" Then sPVen = CDbl(TxtTotal) + CDbl(TxtDescuento) - CDbl(TxtRecargo)   'Ricky
    
    
    calculo
    Adicionar(0).Enabled = True
    Adicionar(0).SetFocus
    swPunto = False
End Sub

Sub calculo()
    Select Case pais
        Case "000", "004", "005"
            If Trim(TxtTotal) = "" Then TxtTotal = 0
            'If TxtTotal = 0 Then TxtTotal = Format(dPrecioUnitario * sCant, "###,###,##0.000")
            If Trim(txtIRBP_ICE.Text) = "" Then txtIRBP_ICE = "0.000"
            If Trim(TxtCantidadPedida) = "" Then TxtCantidadPedida = 0
            sTotal = CDbl(TxtTotal) + CDbl(txtIRBP_ICE.Text) - nOtrosCargosInafecto
            sCant = TxtCantidadPedida
            
            
            If Trim(sTotal) <> "" And Trim(sCant) <> "" Then
                If Trim(sTotal) = "" Or Trim(sTotal) = "0" Or Trim(sCant) = "" Or Trim(sCant) = "0" Then Exit Sub
                dPrecioUnitario = CDbl(sTotal) / CDbl(sCant)
                TxtPrecioUnitario = Format(dPrecioUnitario, "###,###,##0.000")
            Else
               Exit Sub
            End If
            TxtTotal = Format(sTotal, "###,###,##0.000")
            Adicionar(0).Enabled = True
            Adicionar(0).SetFocus
        Case "001"
            If Trim(TxtTotal) = "" Then TxtTotal = 0
            If Trim(TxtCantidadPedida) = "" Then TxtCantidadPedida = 0
            sTotal = TxtTotal
            sCant = TxtCantidadPedida
           
            If Trim(sTotal) <> "" And Trim(sCant) <> "" Then
               If Trim(sTotal) = "" Or Trim(sTotal) = "0" Or Trim(sCant) = "" Or Trim(sCant) = "0" Then Exit Sub
               dPrecioUnitario = CDbl(sTotal) / CDbl(sCant)
               TxtPrecioUnitario = Format(dPrecioUnitario, "###,###,##0.000")
            Else
               Exit Sub
            End If
            TxtTotal = Format(sTotal, "###,###,##0.000")
            Adicionar(0).Enabled = True
        Case "002", "003"
            If Trim(TxtTotal) = "" Then TxtTotal = 0
            If Trim(TxtCantidadPedida) = "" Then TxtCantidadPedida = 0
            If Trim(txtIRBP_ICE) = "" Then txtIRBP_ICE = "0.000"
            sTotal = CDbl(TxtTotal) + CDbl(txtIRBP_ICE.Text) - nOtrosCargosInafecto
            sCant = TxtCantidadPedida
            
            If Trim(sTotal) <> "" And Trim(sCant) <> "" Then
               If Trim(sTotal) = "" Or Trim(sTotal) = "0" Or Trim(sCant) = "" Or Trim(sCant) = "0" Then Exit Sub
                    dPrecioUnitario = CDbl(sTotal) / CDbl(sCant)
                    TxtPrecioUnitario = Format(dPrecioUnitario, "###,###,##0.000")
            Else
               Exit Sub
            End If
            TxtTotal = Format(sTotal, "###,###,##0.000")
            Adicionar(0).Enabled = True
            Adicionar(0).SetFocus
    End Select
End Sub

Sub nPrecio()
    If Len(Trim(TxtTotal)) <> 0 Then
    If Not IsNumeric(TxtTotal) Then TxtTotal = "": Exit Sub
        If Trim(TxtTotal) = "" Then TxtTotal = 0
        If Trim(TxtCantidadPedida) = "" Then TxtCantidadPedida = 0
        CalPrecio CDbl(TxtTotal), TxtCantidadPedida
    End If
    If Len(TxtTotal) = 0 Then
        Adicionar(0).Enabled = False
    Else
        Adicionar(0).Enabled = True
    End If
End Sub

Sub CalPrecio(nMonto As Double, nCant As Double)
    dPrecioUnitario = CDbl(nMonto) / CDbl(nCant)
    TxtPrecioUnitario = Format(dPrecioUnitario, "###,###,##0.000")
End Sub

Sub PantallaD()
    nOtrosCargosInafecto = 0
    txtIRBP_ICE = ""
    TxtCodigo = ""
    'TxtCodigoBarra = ""
    txtArticulo = ""
    CboUnidad = ""
    TxtCantidadPedida = ""
    LblStockActuales = ""
    TxtTotal = ""
    
    dPrecioUnitario = 0
    TxtPrecioUnitario = ""
    
    ChkImp1.value = 0
    ChkImp2.value = 0
    ChkImp3.value = 0
    If lDesc = "G" Then ChkDescuento.Enabled = False
    TxtDescuento = 0
    TxtPDescuento = 0
    TxtRecargo = 0
    TxtPRecargo = 0
    LblVencimiento.Visible = False
    dtpVencimiento.Visible = False
    
    Dim sTmp As String
    sTmp = Calcular("Select tResumido As Codigo From vMoneda Where Codigo='" & frmDocumentoCompra.CboMoneda.BoundText & "'", CnAlmacen)
    If sTmp = "0" Then sTmp = ""
    LblMoneda(0) = sTmp
    LblMoneda(1) = LblMoneda(0)
    LblMoneda(4) = LblMoneda(0)
End Sub

Public Sub InsertarData()
    Dim Gi As Integer
    Dim sItem As Double
    Dim sPVU  As Double
    Dim AnCantidad As Double
    Dim AnPrecio As Double
    Dim AnPVenta As Double
    Dim AnDescuento As Double
    Dim AnRecargo As Double
    Dim AnImpuesto1 As Double
    Dim AnImpuesto2 As Double
    Dim AnImpuesto3 As Double
    With frmDocumentoCompra
        sPVU = CDbl(TxtTotal) + CDbl(TxtDescuento) - CDbl(TxtRecargo)
        If SwTmp Then
            If lAgregaItem Then
                Isql = "sp_CargaDetalleTemporalProducto 'T', '" & .LblCorrelativo & "', '" & TxtCodigo & "', '" & frmDocumentoCompra.sDtemporal & "'"
                Set RsBGrilla = Lib.OpenRecordset(Isql, CnAlmacen)
                If Not RsBGrilla.EOF Then: MsgBox "El artículo seleccionado ya existe...", vbExclamation, sMensaje: Exit Sub
                'Calculo
                Isql = "Select Max(Str(nItem)) As xItem From " & frmDocumentoCompra.sDtemporal & ""
                Set RsMax = Lib.OpenRecordset(Isql, CnAlmacen)
                If Not RsMax.EOF Then
                    sItem = IIf(IsNull(RsMax!xItem), 0, RsMax!xItem)
                End If
                sItem = sItem + 1
                sDocumento = Trim(.LblPDoc) & Trim(.txtSerie) & Trim(.txtCorrelativo)
                Screen.MousePointer = vbHourglass
                
                Select Case pais
                    Case "000", "002", "004", "003", "005"
                        If sTotal <> "" Then
                            IMP1 = TxtImp1
                            Imp2 = TxtImp2
                            Imp3 = TxtImp3
                            ntImp = IIf(ChkImp1.value = 1, CDbl(IMP1), 0) + IIf(ChkImp2.value = 1, CDbl(Imp2), 0) + IIf(ChkImp3.value = 1, CDbl(Imp3), 0)
                            sPorcentajes = CDbl(ntImp) / 100 + 1
                            sSubTot = CDbl(sTotal) - nOtrosCargosInafecto
                            
                            simp1 = IIf(ChkImp1.value = 1, (sSubTot * (CDbl(TxtImp1) / 100)), 0)
                            simp2 = IIf(ChkImp2.value = 1, (sSubTot * (CDbl(TxtImp2) / 100)), 0)
                            simp3 = IIf(ChkImp3.value = 1, (sSubTot * (CDbl(TxtImp3) / 100)), 0)
                        End If
                        
                        If Trim(TxtPrecioUnitario) = "" Then TxtPrecioUnitario = 0
                        Isql = "Insert Into " & frmDocumentoCompra.sDtemporal & _
                               " (tDocumento, tCodigoProveedor, nItem, tCodigoProducto, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                               " tUnidadEntrada, nCantidad, nPrecio, nPVenta, nNeto, nDescuento, nRecargo, nPorcentaje1, nPorcentaje2, nPorcentaje3, " & _
                               " nTotal, tTipoIngreso, lIncluidoImpuesto1, lIncluidoImpuesto2, lIncluidoImpuesto3, tCorrelativo, tEstadoDocumento, " & _
                               " tCuentacontable, tCentroCosto, tAreaOrigen, nRetencion1, nRetencion2, nRetencion3, nPRetencion1, nPRetencion2, nPRetencion3, lIncluidoRetencion1, lIncluidoRetencion2, lIncluidoRetencion3, nICE, nOtrosCargosInafecto)" & _
                               " Values( '" & sDocumento & "', " & "'" & .TxtProveedor & "', " & "" & sItem & ", " _
                             & "'" & TxtCodigo & "', " & "" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ", " _
                             & "" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ", " & "" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                             & "'" & CboUnidad.BoundText & "', " & "" & CDbl(TxtCantidadPedida) & ", " _
                             & "" & dPrecioUnitario & ", " & "" & sPVU & ", " & "" & CDbl(TxtTotal) & ", " _
                             & "" & CDbl(TxtDescuento) & ", " & "" & CDbl(TxtRecargo) & ", " _
                             & "" & IIf(simp1 > 0, IIf(pais = "005", IMP1, nPorcentaje1), 0) & ", " & "" & IIf(simp2 > 0, nPorcentaje2, 0) & ", " _
                             & "" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " & "" & CDbl(TxtTotal) + simp1 + simp2 + simp3 & ", " _
                             & " '" & tTipoIngreso & "', " & "" & IIf(simp1 > 0, 1, 0) & ", " _
                             & "" & IIf(simp2 > 0, 1, 0) & ", " & "" & IIf(simp3 > 0, 1, 0) & ",'" & .LblCorrelativo & "','01', " _
                             & " '" & Trim(txtCuenta.Text) & "', '" & Trim(txtCC.Text) & "','',0.000,0.000,0.000,0.000,0.000,0.000,0,0,0,0.000," & CDbl(IIf(IsNumeric(txtIRBP_ICE.Text), txtIRBP_ICE.Text, 0)) & ")"
                        CnAlmacen.Execute Isql
                        
                    Case "001"
                        If sTotal <> "" Then
                            IMP1 = TxtImp1
                            Imp2 = TxtImp2
                            Imp3 = TxtImp3
                            ntImp = IIf(ChkImp1.value = 1, CDbl(IMP1), 0) + IIf(ChkImp2.value = 1, CDbl(Imp2), 0) + IIf(ChkImp3.value = 1, CDbl(Imp3), 0)
                            sPorcentajes = CDbl(ntImp) / 100 + 1
                            sSubTot = CDbl(sTotal - nOtrosCargosInafecto)

                            simp1 = IIf(ChkImp1.value = 1, (sSubTot * (CDbl(TxtImp1) / 100)), 0)
                            simp2 = IIf(ChkImp2.value = 1, (sSubTot * (CDbl(TxtImp2) / 100)), 0)
                            simp3 = IIf(ChkImp3.value = 1, (sSubTot * (CDbl(TxtImp3) / 100)), 0)
                            'sRet1 = IIf(ChkRet1.value = 1, (sSubTot * (CDbl(TxtRet1) / 100)), 0)
                            'sRet2 = IIf(ChkRet2.value = 1, (sSubTot * (CDbl(TxtRet2) / 100)), 0)
                        End If
                        
                        If Trim(TxtPrecioUnitario) = "" Then TxtPrecioUnitario = 0
                        Isql = "Insert Into " & frmDocumentoCompra.sDtemporal & _
                               " (tDocumento, tCodigoProveedor, nItem, tCodigoProducto, nImpuesto1, nImpuesto2, nImpuesto3, " & _
                               " tUnidadEntrada, nCantidad, nPrecio, nPVenta, nNeto, nDescuento, nRecargo, nPorcentaje1, nPorcentaje2, nPorcentaje3, " & _
                               " nTotal, tTipoIngreso, lIncluidoImpuesto1, lIncluidoImpuesto2, lIncluidoImpuesto3, tCorrelativo, tEstadoDocumento, " & _
                               " tCuentacontable, tCentroCosto, tAreaOrigen,  nRetencion1, nRetencion2, nRetencion3, nPRetencion1, nPRetencion2, nPRetencion3, lIncluidoRetencion1, lIncluidoRetencion2, lIncluidoRetencion3, nICE, lAsumirIUE_IT, nOtrosCargosInafecto)" & _
                               " Values( '" & sDocumento & "', " & "'" & .TxtProveedor & "', " & "" & sItem & ", " _
                             & "'" & TxtCodigo & "', " & "" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ", " _
                             & "" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ", " & "" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                             & "'" & CboUnidad.BoundText & "', " & "" & CDbl(TxtCantidadPedida) & ", " _
                             & "" & dPrecioUnitario - ((ntImp * dPrecioUnitario) / 100) & ", " & "" & sPVU & ", " & "" & CDbl(TxtTotal) - (simp1 + simp2 + simp3) & ", " _
                             & "" & CDbl(TxtDescuento) & ", " & "" & CDbl(TxtRecargo) & ", " _
                             & "" & IIf(simp1 > 0, nPorcentaje1, 0) & ", " & "" & IIf(simp2 > 0, nPorcentaje2, 0) & ", " _
                             & "" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " & "" & CDbl(TxtTotal) - (simp1 + simp2 + simp3) & ", " _
                             & " '" & .CboTipoIngreso.BoundText & "', " & "" & IIf(simp1 > 0, 1, 0) & ", " _
                             & "" & IIf(simp2 > 0, 1, 0) & ", " & "" & IIf(simp3 > 0, 1, 0) & ",'" & .LblCorrelativo & "','01', " _
                             & " '" & Trim(txtCuenta.Text) & "','" & Trim(txtCC.Text) & "',''," & "" & "0.000" & "," & "" & "0.000" & "," & "" & "0.000" & "," & "" & 0 & "," & "" & 0 & "," & "" & 0 & "," & "" & 0 & "," & "" & 0 & "," & 0 & ",0," & IIf(chkIUE_IT.value = 1, 1, 0) & ",0)"

                        CnAlmacen.Execute Isql

                End Select
                .LblItems = .LblItems + 1
            Else
             
             '*********************************************************
            
                'Linea Agregada por Luis
                sTotal = CDbl(TxtTotal.Text)
                'Fin de Linea Agregada
                 
                If sTotal <> "" Then
                    IMP1 = TxtImp1: Imp2 = TxtImp2: Imp3 = TxtImp3
                    ntImp = IIf(ChkImp1.value = 1, CDbl(IMP1), 0) + IIf(ChkImp2.value = 1, CDbl(Imp2), 0) + IIf(ChkImp3.value = 1, CDbl(Imp3), 0)
                    sPorcentajes = ntImp / 100 + 1
                    sSubTot = sTotal - nOtrosCargosInafecto
                    simp1 = IIf(ChkImp1.value = 1, (sSubTot * (CDbl(TxtImp1) / 100)), 0)
                    simp2 = IIf(ChkImp2.value = 1, (sSubTot * (CDbl(TxtImp2) / 100)), 0)
                    simp3 = IIf(ChkImp3.value = 1, (sSubTot * (CDbl(TxtImp3) / 100)), 0)
                    
                    'sRet1 = IIf(ChkRet1.value = 1, (sSubTot * (CDbl(TxtRet1) / 100)), 0)
                    'sRet2 = IIf(ChkRet2.value = 1, (sSubTot * (CDbl(TxtRet2) / 100)), 0)
                                
                End If
                    
                '*********************************************************
                
                Select Case pais
                    Case "000", "002", "004", "003", "005"
                        Isql = "Update " & frmDocumentoCompra.sDtemporal & " Set" _
                             & " nPVenta =" & CDbl(sPVU) & ", " _
                             & " nTotal =" & CDbl(TxtTotal) + simp1 + simp2 + simp3 & ", " _
                             & " nImpuesto1=" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ", " _
                             & " nImpuesto2=" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ", " _
                             & " nImpuesto3=" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                             & " nPorcentaje1=" & IIf(simp1 > 0, IIf(pais = "005", IMP1, nPorcentaje1), 0) & ", " _
                             & " nPorcentaje2=" & IIf(simp2 > 0, nPorcentaje2, 0) & ", " _
                             & " nPorcentaje3=" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " _
                             & " nRetencion1=0.000, nRetencion2=0.000, nRetencion3=0.000, " _
                             & " nPRetencion1=0, nPRetencion2=0, nPRetencion3=0, " _
                             & " tUnidadEntrada='" & CboUnidad.BoundText & "', " _
                             & " nCantidad=" & CDbl(TxtCantidadPedida) & ", " _
                             & " nPrecio=" & dPrecioUnitario & ", " _
                             & " nNeto=" & CDbl(TxtCantidadPedida) * dPrecioUnitario & ", " _
                             & " nDescuento=" & CDbl(TxtDescuento) & ", " _
                             & " nRecargo=" & CDbl(TxtRecargo) & ", " _
                             & " tTipoIngreso='" & tTipoIngreso & "', " _
                             & " lIncluidoImpuesto1=" & IIf(simp1 > 0, 1, 0) & ", " _
                             & " lIncluidoImpuesto2=" & IIf(simp2 > 0, 1, 0) & ", " _
                             & " lIncluidoImpuesto3=" & IIf(simp3 > 0, 1, 0) & ", " _
                             & " lIncluidoRetencion1=0, lIncluidoRetencion2=0, lIncluidoRetencion3=0, nICE=0, " _
                             & " tCuentacontable='" & txtCuenta.Text & "', tCentroCosto='" & txtCC.Text & "', nOtrosCargosInafecto=" & CDbl(txtIRBP_ICE.Text) & " " _
                             & " Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "'"
                    Case "001"
                        Isql = "Update " & frmDocumentoCompra.sDtemporal & " Set" _
                             & " nPVenta =" & CDbl(sPVU) & ", " _
                             & " nTotal =" & CDbl(TxtTotal) - (simp1 + simp2 + simp3) & ", " _
                             & " nImpuesto1=" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ",  nImpuesto2=" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ",  nImpuesto3=" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                             & " nPorcentaje1=" & IIf(simp1 > 0, nPorcentaje1, 0) & ",  nPorcentaje2=" & IIf(simp2 > 0, nPorcentaje2, 0) & ",  nPorcentaje3=" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " _
                             & " nRetencion1=" & "0.000" & ", nRetencion2=" & "0.000" & ", nRetencion3=" & "0.000" & ", " _
                             & " nPRetencion1=" & 0 & ",  nPRetencion2=" & 0 & ",  nPRetencion3=" & 0 & ", " _
                             & " tUnidadEntrada='" & CboUnidad.BoundText & "', " _
                             & " nCantidad=" & CDbl(TxtCantidadPedida) & ", " _
                             & " nPrecio=" & dPrecioUnitario - ((ntImp * dPrecioUnitario) / 100) & ", " _
                             & " nNeto=" & CDbl(TxtCantidadPedida) * dPrecioUnitario - (simp1 + simp2 + simp3) & ", " _
                             & " nDescuento=" & CDbl(TxtDescuento) & ", " _
                             & " nRecargo=" & CDbl(TxtRecargo) & ", " _
                             & " tTipoIngreso='" & .CboTipoIngreso.BoundText & "', " _
                             & " lIncluidoImpuesto1=" & IIf(simp1 > 0, 1, 0) & ", lIncluidoImpuesto2=" & IIf(simp2 > 0, 1, 0) & ",  lIncluidoImpuesto3=" & IIf(simp3 > 0, 1, 0) & ", " _
                             & " lIncluidoRetencion1=" & 0 & ", " _
                             & " lIncluidoRetencion2=" & 0 & ", " _
                             & " lIncluidoRetencion3=" & 0 & ", " _
                             & " tCuentacontable='" & txtCuenta.Text & "', tCentroCosto='" & txtCC.Text & "', lAsumirIUE_IT=" & IIf(chkIUE_IT.value = 1, 1, 0) & " " _
                             & " Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "'"
                End Select
                CnAlmacen.Execute Isql
            End If
            iRow = frmDocumentoCompra.RsDtemporal.AbsolutePosition
            frmDocumentoCompra.RsDtemporal.Requery
            If frmDocumentoCompra.RsDtemporal.RecordCount > 0 Then frmDocumentoCompra.RsDtemporal.AbsolutePosition = IIf(iRow = -1, 1, iRow)
            Screen.MousePointer = vbDefault
            DoEvents
            frmDocumentoCompra.cmdOpcion(10).Enabled = True
'            If frmDocumentoCompra.LblSumDescuento > "0.000" Then
'                frmDocumentoCompra.cmdDescuento.Enabled = True
'            End If
            
            '****************** Verifica si ya no hay descuentos por item ********************
            If lDesc = "I" Then
                Isql = "Select Sum(nDescuento) As xDesc From " & frmDocumentoCompra.sDtemporal & " where tCorrelativo='" & .LblCorrelativo & "'"
                Set RsMax = Lib.OpenRecordset(Isql, CnAlmacen)
                If Not RsMax.EOF Then
                    If RsMax!xDesc = 0 Then lDesc = ""
                End If
            End If
            
        Else
            If lAgregaItem Then
                'Isql = "Select * From dDocumento Where tCodigoProducto='" & txtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "'"
                 Isql = "sp_CargaDetalleTemporalProducto 'N', '" & .LblCorrelativo & "', '" & TxtCodigo & "', ''"
                 Set RsBGrilla = Lib.OpenRecordset(Isql, CnAlmacen)
                 If Not RsBGrilla.EOF Then: MsgBox "El artículo ya existe...", vbExclamation, sMensaje: Exit Sub
                'Calculo
                 Isql = "Select Max(Str(nItem)) As xItem From DDocumento Where tCorrelativo='" & .LblCorrelativo & "'"
                 Set RsMax = Lib.OpenRecordset(Isql, CnAlmacen)
                 If Not RsMax.EOF Then
                    sItem = IIf(IsNull(RsMax!xItem), 0, RsMax!xItem)
                 End If
                 sItem = sItem + 1
                
                 sDocumento = Trim(.LblPDoc) & Trim(.txtSerie) & Trim(.txtCorrelativo)
                 Screen.MousePointer = vbHourglass
             
                If lAuditoria Then
                    lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "DDOCUMENTO", "INGRESO DE DOCUMENTOS - DETALLE", "01", IIf(sUsuario <> "", sUsuario, sUsuario), .LblCorrelativo, sItem, "tCorrelativo", "Correlativo", .LblCorrelativo, "nItem", "Item", sItem, _
                                                "nCantidad", "Cantidad", CDbl(TxtCantidadPedida), "nPrecio", "Precio", dPrecioUnitario, "nPVenta", "Precio de Venta", CDbl(sPVU), "nDescuento", "Monto de Descuento", CDbl(TxtDescuento), "nRecargo", "Monto de Recargo", CDbl(TxtRecargo), "nImpuesto1", "Monto Impuesto 1", CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")), "nImpuesto2", "Monto Impuesto 2", CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")), "nImpuesto3", "Monto Impuesto 3", CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")))
                                                
                    If lAuditoria = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
             
                '*********************************************************
                 Select Case pais
                    Case "000", "002", "004", "003", "005"
                        If sTotal <> "" Then
                           IMP1 = TxtImp1: Imp2 = TxtImp2: Imp3 = TxtImp3
                           ntImp = IIf(ChkImp1.value = 1, CDbl(IMP1), 0) + IIf(ChkImp2.value = 1, CDbl(Imp2), 0) + IIf(ChkImp3.value = 1, CDbl(Imp3), 0)
                           sPorcentajes = ntImp / 100 + 1
                           sSubTot = sTotal - nOtrosCargosInafecto
                           simp1 = IIf(ChkImp1.value = 1, (sSubTot * (CDbl(TxtImp1) / 100)), 0)
                           simp2 = IIf(ChkImp2.value = 1, (sSubTot * (CDbl(TxtImp2) / 100)), 0)
                           simp3 = IIf(ChkImp3.value = 1, (sSubTot * (CDbl(TxtImp3) / 100)), 0)
                        End If
                  
                '*********************************************************
              
                        Isql = "Insert Into dDocumento " & _
                               " (tDocumento,tCodigoProveedor,nItem,tCodigoProducto, nImpuesto1, nImpuesto2, nImpuesto3," & _
                               " tUnidadEntrada, nCantidad, nPrecio, nPVenta, nNeto, nDescuento, nRecargo, nPorcentaje1, nPorcentaje2, nPorcentaje3, " & _
                               " nTotal, tTipoIngreso, lIncluidoImpuesto1, lIncluidoImpuesto2, lIncluidoImpuesto3, tCorrelativo, tEstadoDocumento, " & _
                               " tCuentaContable, tCentroCosto, tAreaOrigen, nRetencion1, nRetencion2, nRetencion3, nPRetencion1, nPRetencion2, nPRetencion3, lIncluidoRetencion1, lIncluidoRetencion2, lIncluidoRetencion3, nICE,nOtrosCargosInafecto)" & _
                               " values( '" & sDocumento & "', " & "'" & .TxtProveedor.Text & "', " & "" & sItem & ", " _
                               & "'" & TxtCodigo & "', " & "" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ", " _
                               & "" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ", " & "" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                               & "'" & CboUnidad.BoundText & "', " & "" & CDbl(TxtCantidadPedida) & ", " _
                               & "" & dPrecioUnitario & ", " & "" & sPVU & ", " & "" & CDbl(TxtTotal) & ", " _
                               & "" & CDbl(TxtDescuento) & ", " & "" & CDbl(TxtRecargo) & ", " _
                               & "" & IIf(simp1 > 0, IIf(pais = "005", IMP1, nPorcentaje1), 0) & ", " & "" & IIf(simp2 > 0, nPorcentaje2, 0) & ", " _
                               & "" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " & "" & CDbl(TxtTotal) + simp1 + simp2 + simp3 & ", " _
                               & " '" & tTipoIngreso & "', " & "" & IIf(simp1 > 0, 1, 0) & ", " _
                               & "" & IIf(simp2 > 0, 1, 0) & ", " & "" & IIf(simp3 > 0, 1, 0) & ",'" & .LblCorrelativo & "','01', " _
                               & "'" & Trim(txtCuenta.Text) & "', '" & Trim(txtCC.Text) & "','',0.000,0.000,0.000,0.000,0.000,0.000,0,0,0,0.000, " & CDbl(IIf(IsNumeric(txtIRBP_ICE.Text), txtIRBP_ICE.Text, 0)) & " )"
                        CnAlmacen.Execute Isql
                    Case "001"
                        If sTotal <> "" Then
                           IMP1 = TxtImp1: Imp2 = TxtImp2: Imp3 = TxtImp3
                           ntImp = IIf(ChkImp1.value = 1, CDbl(IMP1), 0) + IIf(ChkImp2.value = 1, CDbl(Imp2), 0) + IIf(ChkImp3.value = 1, CDbl(Imp3), 0)
                           sPorcentajes = ntImp / 100 + 1
                           sSubTot = CDbl(sTotal) - CDbl(nOtrosCargosInafecto)
                           simp1 = IIf(ChkImp1.value = 1, (sSubTot * (CDbl(TxtImp1) / 100)), 0)
                           simp2 = IIf(ChkImp2.value = 1, (sSubTot * (CDbl(TxtImp2) / 100)), 0)
                           simp3 = IIf(ChkImp3.value = 1, (sSubTot * (CDbl(TxtImp3) / 100)), 0)
                           'sRet1 = IIf(ChkRet1.value = 1, (sSubTot * (CDbl(TxtRet1) / 100)), 0)
                           'sRet2 = IIf(ChkRet2.value = 1, (sSubTot * (CDbl(TxtRet2) / 100)), 0)
                        End If
                  
                '*********************************************************
              
                        Isql = "Insert Into dDocumento " & _
                               " (tDocumento,tCodigoProveedor,nItem,tCodigoProducto,nImpuesto1,nImpuesto2,nImpuesto3," & _
                               " tUnidadEntrada, nCantidad, nPrecio, nPVenta, nNeto, nDescuento, nRecargo, nPorcentaje1, nPorcentaje2, nPorcentaje3, " & _
                               " nTotal, tTipoIngreso, lIncluidoImpuesto1, lIncluidoImpuesto2, lIncluidoImpuesto3, tCorrelativo, tEstadoDocumento, " & _
                               " tCuentaContable, tCentroCosto, nRetencion1, nRetencion2, nRetencion3, nPRetencion1, nPRetencion2, nPRetencion3, lIncluidoRetencion1, lIncluidoRetencion2, lIncluidoRetencion3, nICE, lAsumirIUE_IT,nOtrosCargosInafecto)" & _
                               " values( '" & sDocumento & "', " & "'" & .TxtProveedor.Text & "', " & "" & sItem & ", " _
                               & "'" & TxtCodigo & "', " & "" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ", " _
                               & "" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ", " & "" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                               & "'" & CboUnidad.BoundText & "', " & "" & CDbl(TxtCantidadPedida) & ", " _
                               & "" & dPrecioUnitario - ((ntImp * dPrecioUnitario) / 100) & ", " & "" & sPVU & ", " & "" & CDbl(TxtTotal) - (simp1 + simp2 + simp3) & ", " _
                               & "" & CDbl(TxtDescuento) & ", " & "" & CDbl(TxtRecargo) & ", " _
                               & "" & IIf(simp1 > 0, nPorcentaje1, 0) & ", " & "" & IIf(simp2 > 0, nPorcentaje2, 0) & ", " _
                               & "" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " & "" & CDbl(TxtTotal) - (simp1 + simp2 + simp3) & ", " _
                               & " '" & .CboTipoIngreso.BoundText & "', " & "" & IIf(simp1 > 0, 1, 0) & ", " _
                               & "" & IIf(simp2 > 0, 1, 0) & ", " & "" & IIf(simp3 > 0, 1, 0) & ",'" & .LblCorrelativo & "','01', " _
                               & " '" & Trim(txtCuenta.Text) & "', '" & Trim(txtCC.Text) & "'," & "" & "0.000" & "," & "" & "0.000" & "," & "" & "0.000" & "," & "" & 0 & "," & "" & 0 & "," & "" & 0 & "," & "" & 0 & "," & "" & 0 & "," & "" & 0 & ",0," & IIf(chkIUE_IT.value = 1, 1, 0) & ",0)" '" & IIf(ChkICE.value = 1, CDbl(TxtICE.Text), 0#) & "
                               
                        CnAlmacen.Execute Isql
'                        .lCambioDetalle = True
                End Select
                
                'Registro de Modificaciones : Mangos
                '27-07-2016
                'cHistorial.FunInsertatHistorial "002", "001", "", .LblCorrelativo, "", "", "", "", sUsuario, sUsuario, "Se Registró el item " & sItem & " del documento " & Trim(sDocumento) & ""
                'Fin
                
                .LblItems = .LblItems + 1
            Else
                '*********************************************************
                
                 'Modificación CLR 15/12/2011 comprobar que no se cambie el valor de compra de documentos cancelados
                Select Case pais
                    Case "000", "002", "004", "003", "005"
                        If tEstadoDocu = "03" Then
                            sAtotal = Format(CDbl(Calcular("Select nneto as codigo from dDocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)), "###,###,##0.000")
                            If CDbl(TxtTotal.Text) <> sAtotal Then
                                MsgBox "A intentado cambiar el valor de un articulo de un documento ya cancelado," & vbCr & "elimine los pagos de cancelación antes de modificar el valor del documento.", vbExclamation + vbOKOnly, sMensaje
                                TxtTotal.Text = Format(sAtotal, "###,###,##0.000") 'sAtotal
                                Exit Sub
                            End If
                        End If
                    Case "001"
                        If tEstadoDocu = "03" Then
                            sAtotal = Format(CDbl(Calcular("Select npventa as codigo from dDocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)), "###,###,##0.000")
                            If CDbl(TxtTotal.Text) <> sAtotal Then
                                MsgBox "A intentado cambiar el valor de un articulo de un documento ya cancelado," & vbCr & "elimine los pagos de cancelación antes de modificar el valor del documento.", vbExclamation + vbOKOnly, sMensaje
                                TxtTotal.Text = Format(sAtotal, "###,###,##0.000") 'sAtotal
                                Exit Sub
                            End If
                        End If
                End Select
                'Modificación CLR 15/12/2011
                
                 'Linea Agregada por Luis
                 sTotal = CDbl(TxtTotal.Text)
                 'Fin de Linea Agregada
                 
                 If sTotal <> "" Then
                    IMP1 = TxtImp1: Imp2 = TxtImp2: Imp3 = TxtImp3
                    ntImp = IIf(ChkImp1.value = 1, CDbl(IMP1), 0) + IIf(ChkImp2.value = 1, CDbl(Imp2), 0) + IIf(ChkImp3.value = 1, CDbl(Imp3), 0)
                    sPorcentajes = ntImp / 100 + 1
                    sSubTot = CDbl(sTotal) - CDbl(nOtrosCargosInafecto)
                    simp1 = IIf(ChkImp1.value = 1, (sSubTot * (CDbl(TxtImp1) / 100)), 0)
                    simp2 = IIf(ChkImp2.value = 1, (sSubTot * (CDbl(TxtImp2) / 100)), 0)
                    simp3 = IIf(ChkImp3.value = 1, (sSubTot * (CDbl(TxtImp3) / 100)), 0)
                    
                    'sRet1 = IIf(ChkRet1.value = 1, (sSubTot * (CDbl(TxtRet1) / 100)), 0)
                    'sRet2 = IIf(ChkRet2.value = 1, (sSubTot * (CDbl(TxtRet2) / 100)), 0)
                    
                 End If
                  
                '*********************************************************
                
                                    
                      If Trim(TxtTotal) = "" Then TxtTotal = 0
                      If Trim(TxtCantidadPedida) = "" Then TxtCantidadPedida = 0
                      If Trim(TxtPrecioUnitario) = "" Then
                         dPrecioUnitario = 0
                         TxtPrecioUnitario = 0
                      End If
                      If Trim(TxtDescuento) = "" Then TxtDescuento = 0
                      If Trim(TxtRecargo) = "" Then TxtRecargo = 0
                      
                      AnCantidad = Calcular("Select nCantidad as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      AnPrecio = Calcular("Select nPrecio as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      AnPVenta = Calcular("Select nPVenta as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      AnDescuento = Calcular("Select nDescuento as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      AnRecargo = Calcular("Select nRecargo as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      AnImpuesto1 = Calcular("Select nImpuesto1 as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      AnImpuesto2 = Calcular("Select nImpuesto2 as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      AnImpuesto3 = Calcular("Select nImpuesto3 as Codigo from ddocumento Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & "", CnAlmacen)
                      
                Select Case pais
                    Case "000", "002", "004", "003", "005"
                            If lAuditoria Then
                                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "DDOCUMENTO", "INGRESO DE DOCUMENTOS - DETALLE", "02", IIf(sUsuario <> "", sUsuario, sUsuario), .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, "tCorrelativo", "Correlativo", .LblCorrelativo, "nItem", "Item", frmDocumentoCompra.RsDocumentoD!nItem, _
                                                            "nCantidad", "Cantidad", CDbl(TxtCantidadPedida), "nPrecio", "Precio", dPrecioUnitario, "nPVenta", "Precio de Venta", CDbl(sPVU), "nDescuento", "Monto de Descuento", CDbl(TxtDescuento), "nRecargo", "Monto de Recargo", CDbl(TxtRecargo), "nImpuesto1", "Monto Impuesto 1", CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")), "nImpuesto2", "Monto Impuesto 2", CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")), "nImpuesto3", "Monto Impuesto 3", CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")))
                                                            
                                If lAuditoria = False Then
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                End If
                            End If
                    Case "001"
                            If lAuditoria Then
                                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "DDOCUMENTO", "INGRESO DE DOCUMENTOS - DETALLE", "02", IIf(sUsuario <> "", sUsuario, sUsuario), .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, "tCorrelativo", "Correlativo", .LblCorrelativo, "nItem", "Item", frmDocumentoCompra.RsDocumentoD!nItem, _
                                                            "nCantidad", "Cantidad", CDbl(TxtCantidadPedida), "nPrecio", "Precio", dPrecioUnitario - ((ntImp * dPrecioUnitario) / 100), "nPVenta", "Precio de Venta", CDbl(sPVU), "nDescuento", "Monto de Descuento", CDbl(TxtDescuento), "nRecargo", "Monto de Recargo", CDbl(TxtRecargo), "nImpuesto1", "Monto Impuesto 1", CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")), "nImpuesto2", "Monto Impuesto 2", CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")), "nImpuesto3", "Monto Impuesto 3", CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")))
                                                            
                                If lAuditoria = False Then
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                End If
                            End If
                End Select
                      
                Select Case pais
                    Case "000", "002", "004", "003", "005"
                        Isql = "Update dDocumento Set" _
                             & " tCodigoProducto ='" & TxtCodigo & "', " _
                             & " nPVenta =" & CDbl(sPVU) & ", " _
                             & " nTotal =" & CDbl(TxtTotal) + simp1 + simp2 + simp3 & ", " _
                             & " nImpuesto1=" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ", " _
                             & " nImpuesto2=" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ", " _
                             & " nImpuesto3=" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                             & " nPorcentaje1=" & IIf(simp1 > 0, IIf(pais = "005", IMP1, nPorcentaje1), 0) & ", " _
                             & " nPorcentaje2=" & IIf(simp2 > 0, nPorcentaje2, 0) & ", " _
                             & " nPorcentaje3=" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " _
                             & " nRetencion1=0.000, nRetencion2=0.000, nRetencion3=0.000, " _
                             & " nPRetencion1=0, nPRetencion2=0, nPRetencion3=0, " _
                             & " tUnidadEntrada='" & CboUnidad.BoundText & "', " _
                             & " nCantidad=" & CDbl(TxtCantidadPedida) & ", " _
                             & " nPrecio=" & dPrecioUnitario & ", " _
                             & " nNeto=" & CDbl(TxtCantidadPedida) * dPrecioUnitario & ", " _
                             & " nDescuento=" & CDbl(TxtDescuento) & ", " _
                             & " nRecargo=" & CDbl(TxtRecargo) & ", " _
                             & " tTipoIngreso='" & tTipoIngreso & "', " _
                             & " lIncluidoImpuesto1=" & IIf(simp1 > 0, 1, 0) & ", " _
                             & " lIncluidoImpuesto2=" & IIf(simp2 > 0, 1, 0) & ", " _
                             & " lIncluidoImpuesto3=" & IIf(simp3 > 0, 1, 0) & ", " _
                             & " lIncluidoRetencion1=0, lIncluidoRetencion2=0, lIncluidoRetencion3=0, nICE=0, " _
                             & " tCuentaContable='" & txtCuenta.Text & "', tCentroCosto='" & txtCC.Text & "',nOtrosCargosInafecto=" & txtIRBP_ICE.Text & " " _
                             & " Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & ""
                    Case "001"
                        Isql = "Update dDocumento Set" _
                             & " tCodigoProducto ='" & TxtCodigo & "', " _
                             & " nPVenta =" & CDbl(sPVU) & ", " _
                             & " nTotal =" & CDbl(TxtTotal) - (simp1 + simp2 + simp3) & ", " _
                             & " nImpuesto1=" & CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) & ", " & " nImpuesto2=" & CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) & ", " & " nImpuesto3=" & CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) & ", " _
                             & " nPorcentaje1=" & IIf(simp1 > 0, nPorcentaje1, 0) & ", " & " nPorcentaje2=" & IIf(simp2 > 0, nPorcentaje2, 0) & ", " & " nPorcentaje3=" & IIf(simp3 > 0, nPorcentaje3, 0) & ", " _
                             & " nRetencion1=" & "0.000" & ", " & " nRetencion2=" & "0.000" & ", " & " nRetencion3=" & "0.000" & ", " _
                             & " nPRetencion1=" & 0 & ", " & " nPRetencion2=" & 0 & ", " & " nPRetencion3=" & 0 & ", " _
                             & " tUnidadEntrada='" & CboUnidad.BoundText & "', " _
                             & " nCantidad=" & CDbl(TxtCantidadPedida) & ", " _
                             & " nPrecio=" & dPrecioUnitario - ((ntImp * dPrecioUnitario) / 100) & ", " _
                             & " nNeto=" & CDbl(TxtCantidadPedida) * dPrecioUnitario - (simp1 + simp2 + simp3) & ", " _
                             & " nDescuento=" & CDbl(TxtDescuento) & ", " _
                             & " nRecargo=" & CDbl(TxtRecargo) & ", " _
                             & " tTipoIngreso='" & .CboTipoIngreso.BoundText & "', " _
                             & " lIncluidoImpuesto1=" & IIf(simp1 > 0, 1, 0) & ", " _
                             & " lIncluidoImpuesto2=" & IIf(simp2 > 0, 1, 0) & ", " _
                             & " lIncluidoImpuesto3=" & IIf(simp3 > 0, 1, 0) & ", " _
                             & " lIncluidoRetencion1=" & 0 & ", " _
                             & " lIncluidoRetencion2=" & 0 & ", " _
                             & " lIncluidoRetencion3=" & 0 & ", " _
                             & " tCuentaContable='" & txtCuenta.Text & "', tCentroCosto='" & txtCC.Text & "', lAsumirIUE_IT=" & IIf(chkIUE_IT.value = 1, 1, 0) & "" _
                             & " Where tCodigoProducto ='" & TxtCodigo & "' And tCorrelativo='" & .LblCorrelativo & "' And nItem = " & frmDocumentoCompra.RsDocumentoD!nItem & ""
                             
'                             .lCambioDetalle = True
                End Select
                CnAlmacen.Execute Isql
                
                'Registro de Modificaciones : Mangos
                '27-07-2016
                
                Select Case pais
                    Case "000", "004", "005"
'                        If AnPrecio <> dPrecioUnitario Then
'                            cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, "Precio", AnPrecio, dPrecioUnitario, sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                        End If
                    Case "001"
'                        If AnPrecio <> dPrecioUnitario - ((ntImp * dPrecioUnitario) / 100) Then
'                            cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, "Precio", AnPrecio, dPrecioUnitario - ((ntImp * dPrecioUnitario) / 100), sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                        End If
                End Select
'                If AnPVenta <> CDbl(sPVU) Then
'                    cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, "Precio Venta", AnPVenta, CDbl(sPVU), sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                End If
'                If AnDescuento <> CDbl(TxtDescuento) Then
'                    cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, "Descuento", AnDescuento, CDbl(TxtDescuento), sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                End If
'                If AnRecargo <> CDbl(TxtRecargo) Then
'                    cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, "Recargo", AnRecargo, CDbl(TxtRecargo), sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                End If
'                If AnImpuesto1 <> CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")) Then
'                    cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, sImpuesto1, AnImpuesto1, CDbl(IIf(ChkImp1.value = 1, simp1, "0.000")), sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                End If
'                If AnImpuesto2 <> CDbl(IIf(ChkImp2.value = 1, simp2, "0.000")) Then
'                    cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, sImpuesto2, AnImpuesto2, CDbl(IIf(ChkImp1.value = 1, simp2, "0.000")), sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                End If
'                If AnImpuesto3 <> CDbl(IIf(ChkImp3.value = 1, simp3, "0.000")) Then
'                    cHistorial.FunInsertatHistorial "002", "002", "", .LblCorrelativo, frmDocumentoCompra.RsDocumentoD!nItem, sImpuesto3, AnImpuesto3, CDbl(IIf(ChkImp1.value = 1, simp3, "0.000")), sUsuario, sUsuarioAutoriza, "Se Actualizó el item " & frmDocumentoCompra.RsDocumentoD!nItem & " del documento " & Trim(sDocumento) & ""
'                End If
                'Fin
                
            End If
            Dim xUltPos As String
            

            
            iRow = frmDocumentoCompra.RsDocumentoD.AbsolutePosition
            frmDocumentoCompra.RsDocumentoD.Requery
            If frmDocumentoCompra.RsDocumentoD.RecordCount > 0 Then frmDocumentoCompra.RsDocumentoD.AbsolutePosition = IIf(iRow = -1, 1, iRow)
            
            frmDocumentoCompra.CalculaTabla
            frmDocumentoCompra.Actualiza
            .lCambioDetalle = True
'            frmDocumentoCompra.lDocCambio = True
            
'            iRow = frmDocumento.RsCabecera.AbsolutePosition
'            frmDocumento.RsCabecera.Requery
'            If frmDocumento.RsCabecera.RecordCount > 0 Then frmDocumento.RsCabecera.AbsolutePosition = IIf(iRow = -1, 1, iRow)
                    
            Screen.MousePointer = vbDefault
            
        End If
            
        '******** Verifica si ya no hay descuentos por item ********
        If lDesc = "I" Then
           Isql = "Select Sum(nDescuento) As xDesc From dDocumento Where tCorrelativo='" & .LblCorrelativo & "'"
           Set RsMax = Lib.OpenRecordset(Isql, CnAlmacen)
           If Not RsMax.EOF Then
              If RsMax!xDesc = 0 Then lDesc = ""
           End If
        End If
        
        If .LblCorrelativo = "" Then .CmdProcesar.Enabled = False
    End With
End Sub

Sub ActualizaTabla()
    With frmDocumentoCompra
         Isql = "Update MDocumento Set" _
              & " nNeto=" & .LblSubTot & "," _
              & " nTotal=" & .lblTotal & "," _
              & " nImpuesto1=" & .LblTotImp1 & "," _
              & " nImpuesto2=" & .LblTotImp2 & "," _
              & " nImpuesto3=" & .LblTotImp3 & "," _
              & " nDescuento=" & .LblSumDescuento & ", lreplica=1 " _
              & " Where tCorrelativo='" & .LblCorrelativo & "'"
         CnAlmacen.Execute Isql
    End With
End Sub

Sub VerificaFVence()
    'Muestra o no la fecha de Vencimiento...
     
    Dim sFam           As String
    Dim bVence         As Boolean
    Dim bPorcionable   As Boolean
    
    bPorcionable = Calcular("Select lPorcionable As Codigo From tProducto Where tCodigoproducto='" & TxtCodigo & "'", CnAlmacen)
    sFam = Calcular("Select tCodigoSubFamilia As Codigo From tProducto Where tCodigoproducto='" & TxtCodigo & "'", CnAlmacen)
    bVence = Calcular("Select lVencimiento As Codigo From tSubFamilia Where tCodigoSubFamilia='" & sFam & "'", CnAlmacen)
    If bVence Then     'Pedir Fecha de Vencimiento a TODOS los artículos de esa SF
    Else               'Pedir Fecha de Vencimiento sólo a los artículos de esa SF que sean PERECIBLES
        Dim bPerece     As Boolean
        
        bPerece = Calcular("Select lPerecible As Codigo From tProducto Where tCodigoProducto='" & TxtCodigo & "'", CnAlmacen)
        If bPerece Then
        Else
            LblVencimiento.Visible = False
            dtpVencimiento.Visible = False
        End If
    End If
     
    If bPorcionable Then
        LblPorcionable.Visible = True
        Label2.Visible = True
    Else
        LblPorcionable.Visible = False
        Label2.Visible = False
    End If
End Sub

Function VerificaIngreso()
'Function VerificaIngreso(Campo As String, Valor As Integer)
    'TxtCodigoBarra = ""
    'TxtCodigoBarra.SetFocus
End Function

'Private Sub VerificaCBarra()
'    TxtCodigo = Calcular("Select tCodigoProducto As Codigo From tProducto Where tCodigoBarra='" & TxtCodigoBarra & "'", CnAlmacen)
'    If TxtCodigo = "0" Then TxtCodigo = ""
'    txtArticulo = Calcular("Select tDetallado As Codigo From tProducto Where tCodigoProducto='" & TxtCodigo & "'", CnAlmacen)
'    If txtArticulo = "0" Then txtArticulo = ""
'    CboUnidad.BoundText = Calcular("Select tUnidadEntrada As Codigo From tProducto Where tCodigoProducto='" & TxtCodigo & "'", CnAlmacen)
'End Sub

Function VerificaTipoIngreso(ByVal sTipoIngreso As String)
    Dim lTipoIngreso As Boolean
    lTipoIngreso = True
    Select Case sTipoIngreso
        Case Is = "A" 'Activo Fijo
            lTipoIngreso = False
        Case Is = "M" 'Mercaderia
            lTipoIngreso = False
    End Select
    VerificaTipoIngreso = lTipoIngreso
End Function

Sub VentanaCompra()
    ChkExclusivo.Visible = True
    ChkExclusivo.value = 0
'    LlenaArticulo
    Me.Height = 5985
    Centrar Me
End Sub
