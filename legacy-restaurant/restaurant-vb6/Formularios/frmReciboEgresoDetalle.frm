VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReciboEgresoDetalle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5115
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9810
   Icon            =   "frmReciboEgresoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9810
   Begin TabDlg.SSTab SSTab1 
      Height          =   4275
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   7541
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Egresos"
      TabPicture(0)   =   "frmReciboEgresoDetalle.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEstado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtRecibo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtTurno"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtFecha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdEgreso(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdEgreso(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdEgreso(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmReciboEgresoDetalle.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTipoComprobante"
      Tab(1).Control(1)=   "txtComprobante"
      Tab(1).Control(2)=   "txtRSocial"
      Tab(1).Control(3)=   "txtIdentificacion"
      Tab(1).Control(4)=   "cboTipoIdentificacion"
      Tab(1).Control(5)=   "cboCentroCosto"
      Tab(1).Control(6)=   "Label12"
      Tab(1).Control(7)=   "Label11"
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(9)=   "Label8"
      Tab(1).Control(10)=   "Label6"
      Tab(1).Control(11)=   "Label5"
      Tab(1).ControlCount=   12
      Begin VB.CommandButton cmdEgreso 
         BackColor       =   &H00F2FAED&
         Caption         =   "Egreso por compra de mercaderia."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton cmdEgreso 
         BackColor       =   &H00F2FAED&
         Caption         =   "Egreso por devolucón segun nota de credito."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdEgreso 
         BackColor       =   &H007EC642&
         Caption         =   "Egreso por tipo varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   840
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   1680
         TabIndex        =   33
         Top             =   840
         Width           =   7935
         Begin VB.TextBox txtObservacion 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   735
            Left            =   4440
            TabIndex        =   55
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Observación"
            Height          =   495
            Index           =   11
            Left            =   6600
            TabIndex        =   54
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Dolares"
            Height          =   615
            Index           =   10
            Left            =   6675
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   120
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Dolares"
            Height          =   615
            Index           =   9
            Left            =   5475
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   120
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "NC"
            Height          =   525
            Index           =   8
            Left            =   3555
            TabIndex        =   42
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox txtnotacredito 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
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
            Height          =   525
            Left            =   1275
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Dolares"
            Height          =   615
            Index           =   7
            Left            =   4275
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   120
            Width           =   1170
         End
         Begin VB.TextBox txtDescripcion 
            BorderStyle     =   0  'None
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
            Height          =   525
            Left            =   1275
            MaxLength       =   50
            TabIndex        =   39
            Text            =   " "
            Top             =   1200
            Width           =   5220
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
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
            Left            =   1275
            MaxLength       =   15
            TabIndex        =   38
            Text            =   " 0.00"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtAutorizacion 
            BorderStyle     =   0  'None
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
            Height          =   525
            Left            =   1275
            MaxLength       =   50
            TabIndex        =   37
            Text            =   " "
            Top             =   1800
            Width           =   5220
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Descripción"
            Height          =   525
            Index           =   4
            Left            =   6600
            Picture         =   "frmReciboEgresoDetalle.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1200
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Soles"
            Height          =   615
            Index           =   6
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   120
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Autorización"
            Height          =   525
            Index           =   5
            Left            =   6600
            Picture         =   "frmReciboEgresoDetalle.frx":0444
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1800
            Width           =   1170
         End
         Begin MSDataListLib.DataCombo cboTipoEgreso 
            Height          =   315
            Left            =   1275
            TabIndex        =   45
            Top             =   840
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            BackColor       =   16777215
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Doc.Referencia:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   2520
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Egreso :"
            Height          =   195
            Left            =   210
            TabIndex        =   49
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción :"
            Height          =   195
            Left            =   225
            TabIndex        =   48
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Autorización :"
            Height          =   195
            Left            =   195
            TabIndex        =   47
            Top             =   1800
            Width           =   960
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto :"
            Height          =   195
            Left            =   615
            TabIndex        =   46
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.TextBox txtTipoComprobante 
         Height          =   340
         Left            =   -72120
         TabIndex        =   24
         Top             =   2760
         Width           =   4500
      End
      Begin VB.TextBox txtComprobante 
         Height          =   340
         Left            =   -72120
         TabIndex        =   23
         Top             =   2280
         Width           =   4500
      End
      Begin VB.TextBox txtRSocial 
         Height          =   340
         Left            =   -72120
         TabIndex        =   22
         Top             =   1800
         Width           =   4500
      End
      Begin VB.TextBox txtIdentificacion 
         Height          =   340
         Left            =   -72120
         TabIndex        =   21
         Top             =   1320
         Width           =   4500
      End
      Begin VB.TextBox txtFecha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   3795
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   400
         Width           =   1635
      End
      Begin VB.TextBox txtTurno 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   6195
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   400
         Width           =   2175
      End
      Begin VB.TextBox txtRecibo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   400
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo cboTipoIdentificacion 
         Height          =   315
         Left            =   -72120
         TabIndex        =   31
         Top             =   850
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboCentroCosto 
         Height          =   315
         Left            =   -72120
         TabIndex        =   32
         Top             =   3250
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Centro Costo :"
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
         Left            =   -73620
         TabIndex        =   30
         Top             =   3300
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo Comprobante :"
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
         Left            =   -74070
         TabIndex        =   29
         Top             =   2805
         Width           =   1710
      End
      Begin VB.Label Label9 
         Caption         =   "N° Comprobante :"
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
         Left            =   -73905
         TabIndex        =   28
         Top             =   2325
         Width           =   1545
      End
      Begin VB.Label Label8 
         Caption         =   "Razon Social / Nombres :"
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
         Left            =   -74580
         TabIndex        =   27
         Top             =   1875
         Width           =   2235
      End
      Begin VB.Label Label6 
         Caption         =   "N° de Identificación :"
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
         Left            =   -74220
         TabIndex        =   26
         Top             =   1395
         Width           =   1860
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Identificación :"
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
         Left            =   -74385
         TabIndex        =   25
         Top             =   900
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Turno :"
         Height          =   195
         Left            =   5565
         TabIndex        =   20
         Top             =   400
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Recibo :"
         Height          =   195
         Left            =   555
         TabIndex        =   19
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   3195
         TabIndex        =   18
         Top             =   400
         Width           =   540
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Atendido"
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
         Left            =   8400
         TabIndex        =   17
         Top             =   480
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   800
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9750
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4320
      Width           =   9810
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8520
         Picture         =   "frmReciboEgresoDetalle.frx":0546
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7350
         Picture         =   "frmReciboEgresoDetalle.frx":0638
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5010
         Picture         =   "frmReciboEgresoDetalle.frx":073A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6180
         Picture         =   "frmReciboEgresoDetalle.frx":0C6C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4830
         TabIndex        =   1
         Top             =   60
         Width           =   4890
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Index           =   1
            Left            =   480
            MaskColor       =   &H00F2FAED&
            Picture         =   "frmReciboEgresoDetalle.frx":119E
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Index           =   2
            Left            =   960
            MaskColor       =   &H00F2FAED&
            Picture         =   "frmReciboEgresoDetalle.frx":16E0
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Index           =   0
            Left            =   0
            MaskColor       =   &H00F2FAED&
            Picture         =   "frmReciboEgresoDetalle.frx":1C22
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Index           =   5
            Left            =   4350
            MaskColor       =   &H00F2FAED&
            Picture         =   "frmReciboEgresoDetalle.frx":2164
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Index           =   4
            Left            =   3870
            MaskColor       =   &H00F2FAED&
            Picture         =   "frmReciboEgresoDetalle.frx":26A6
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Index           =   3
            Left            =   3390
            MaskColor       =   &H00F2FAED&
            Picture         =   "frmReciboEgresoDetalle.frx":2BE8
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   12
            Top             =   150
            Width           =   1845
         End
      End
   End
End
Attribute VB_Name = "frmReciboEgresoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sEstado As String
Public nMonto As Double
Dim RsEgreso As Recordset
Dim vwNC As Boolean
Dim rsTipoEgreso, rsTIdentificacion, rsCentroCosto As New ADODB.Recordset
Public sMonedaRecibo As String
Public sCompraOK As String
Dim ModoEgreso As String

Public Sub LlenaCombo()
        With cboTipoEgreso
             Isql = "select Codigo, tresumido from vTipoEgreso where lActivo=1 "
             Set rsTipoEgreso = Lib.OpenRecordset(Isql, Cn)
             Set .RowSource = rsTipoEgreso
                 .DataField = "tResumido"
                 .ListField = "tResumido"
                 .BoundColumn = "Codigo"
        End With
End Sub
Public Sub LlenaComboTipoIdentificacion()
    With cboTipoIdentificacion
         Isql = "select Codigo, tResumido from vTipoIdentidad where lActivo=1 "
         Set rsTIdentificacion = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsTIdentificacion
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
End Sub
Public Sub LlenaComboCentroCosto()
    With cboCentroCosto
         Isql = "select TCodigo, CentroCosto from TCENTROCOSTO where lActivo=1 "
         Set rsCentroCosto = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsCentroCosto
             .DataField = "CentroCosto"
             .ListField = "CentroCosto"
             .BoundColumn = "TCodigo"
    End With
End Sub

Sub Asignar()
    With frmReciboEgreso.RsCabecera
        'Cuadro de Texto
        txtRecibo = IIf(IsNull(!tRecibo), "", !tRecibo)
        txtFecha = IIf(IsNull(!fFecha), "", !fFecha)
        nMonto = IIf(IsNull(!nMonto), 0, !nMonto)
        txtMonto.Text = Format(nMonto, "###,###,##0.00")
        txtDescripcion = IIf(IsNull(!tDescripcion), "", !tDescripcion)
        txtAutorizacion = IIf(IsNull(!tAutoriza), "", !tAutoriza)
        txtObservacion = IIf(IsNull(!tObservacion), "", !tObservacion)
        txtTurno = IIf(IsNull(!tTurno), "", !tTurno)
        cboTipoEgreso.BoundText = IIf(IsNull(!ttipoegreso), "", !ttipoegreso)
        cmdOpcion(6).FontBold = IIf(IsNull(!tMONEDA), True, IIf(!tMONEDA = "01", True, False))
        cmdOpcion(7).FontBold = IIf(IsNull(!tMONEDA), False, IIf(!tMONEDA = "01", False, True))
        
        txtIdentificacion.Text = IIf(IsNull(!Identificacion), "", !Identificacion)
        txtRSocial.Text = IIf(IsNull(!razonsocial), "", !razonsocial)
        txtComprobante.Text = IIf(IsNull(!NComprobante), "", !NComprobante)
        txtTipoComprobante.Text = IIf(IsNull(!TipoComprobante), "", !TipoComprobante)
        
        cboTipoIdentificacion.BoundText = IIf(IsNull(!TipoIdentificacion), "", !TipoIdentificacion)
        cboCentroCosto.BoundText = IIf(IsNull(!CentroCosto), "", !CentroCosto)
                
        lblEstado.ForeColor = IIf(IsNull(!tEstadoDocumento), vbBlack, IIf(!tEstadoDocumento = "04", vbRed, vbBlack))
        sEstado = IIf(IsNull(!tEstadoDocumento), "01", !tEstadoDocumento)
        lblEstado.Caption = IIf(IsNull(!tEstadoDocumento), "", IIf(!tEstadoDocumento = "04", "ANULADO", "EMITIDO"))
        txtnotacredito.Text = IIf(IsNull(!dReferencia), "", !dReferencia)
        'If Periodo(!fFecha) Or sEstado = "04" Or sEstado = "02" Or sTurno <> !tTurno Then
        If Periodo(!fFecha) Or sEstado = "04" Or sEstado = "02" Then
           Bloquear True
        Else
           Bloquear False
        End If
        
         If sEstado = "01" Then
            cmdOpcion(1).Enabled = False
        End If
        
        If (!tModoEgreso) = "VARIOS" Then
            cmdEgreso_Click (1)
            
            cmdOpcion(7).Enabled = True
            cmdOpcion(6).Enabled = True
            cmdOpcion(9).Enabled = True
            cmdOpcion(10).Enabled = True
            
            cmdOpcion(4).Enabled = True
            cmdOpcion(5).Enabled = True
            
        ElseIf (!tModoEgreso) = "NOTACREDITO" Then
            cmdEgreso_Click (2)
            txtMonto.Enabled = False
            cmdOpcion(7).Enabled = False
            cmdOpcion(6).Enabled = False
            cmdOpcion(9).Enabled = False
            cmdOpcion(10).Enabled = False
        ElseIf (!tModoEgreso) = "MERCADERIA" Then
            cmdEgreso_Click (3)
            txtMonto.Enabled = False
            cmdOpcion(7).Enabled = False
            cmdOpcion(6).Enabled = False
            cmdOpcion(9).Enabled = False
            cmdOpcion(10).Enabled = False
            
            cmdOpcion(4).Enabled = False
            cmdOpcion(5).Enabled = False
        End If
        
        cmdEgreso(1).Enabled = False
        cmdEgreso(2).Enabled = False
        cmdEgreso(3).Enabled = False
        'NC
        cmdOpcion(8).Enabled = False
        
        'CESAR PERMISOS DE GRUPO
        Dim xGrupo As String
        Dim xRGrupoAcceso As Recordset
        xGrupo = Calcular("Select tGrupoUsuario As Codigo From TUSUARIO where tResumido = '" & sUsuario & "'", Cn)
            
           If xGrupo <> "00" Then
            Isql = "Select * From TGRUPOACCESO WHERE tGrupoUsuario = '" & xGrupo & "'"
            Set xRGrupoAcceso = Lib.OpenRecordset(Isql, Cn)
            '30204010 MODIFICAR
                xRGrupoAcceso.Filter = "tCodigoAcceso=30204010"
                If xRGrupoAcceso.RecordCount > 0 Then
                   cmdOpcion(1).Enabled = False
                Else
                   cmdOpcion(1).Enabled = True
                End If
            End If
    End With
    cmdTexto.Caption = "Registro " & frmReciboEgreso.RsCabecera.AbsolutePosition & " de " & frmReciboEgreso.RsCabecera.RecordCount
End Sub

Private Sub cmdEgreso_Click(Index As Integer)
    
    cmdEgreso(1).FontBold = False
    cmdEgreso(2).FontBold = False
    cmdEgreso(3).FontBold = False
    
    cmdEgreso(1).backColor = &HF2FAED
    cmdEgreso(2).backColor = &HF2FAED
    cmdEgreso(3).backColor = &HF2FAED
    ModoEgreso = ""
    If Index = 1 Then
        cmdEgreso(1).FontBold = True
        cmdEgreso(1).backColor = &H7EC642
        
        cmdOpcion(8).Enabled = False
        ModoEgreso = "VARIOS"
        vwNC = False
    ElseIf Index = 2 Then
        cmdEgreso(2).FontBold = True
        cmdEgreso(2).backColor = &H7EC642
        

        cmdOpcion(8).Enabled = True
        ModoEgreso = "NOTACREDITO"
        vwNC = True
    ElseIf Index = 3 Then
        'If cboTipoEgreso.Text = "" Then MsgBox "Seleccione un Tipo de Egreso", vbExclamation, sMensaje: Exit Sub
               
        sCompraOK = ""
        cmdOpcion(8).Enabled = False
    
        cmdEgreso(3).FontBold = True
        cmdEgreso(3).backColor = &H7EC642
        ModoEgreso = "MERCADERIA"
        vwNC = False
        
        If Sw Then
            'Dim frm As New frmDocumentoCompra
            frmDocumentoCompra.Show vbModal
            
            If sCompraOK = "OK" Then
                If sVar1 <> "" Then
                    txtAutorizacion.Text = "Aut: " & sVar1
                Else
                     txtAutorizacion.Text = "Aut: " & sUsuario
                End If
                cmdOpcion_Click 1
            End If
        End If
    End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmReciboEgreso.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmReciboEgreso.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmReciboEgreso.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmReciboEgreso.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmReciboEgreso.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmReciboEgreso.grdGrilla
    End Select
    Asignar

End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 'Agregar
                
               vwNC = False
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               
               'Cambia el Nombre del Primer Text
               txtTurno.Text = sTurno
               txtFecha.Text = Date
               nMonto = 0
               txtMonto.Text = "0.00"
               lblEstado.Caption = ""
               Bloquear False
               cboTipoEgreso.Text = ""
               cmdOpcion(1).Enabled = True

               cmdEgreso(1).Enabled = True
               cmdEgreso(2).Enabled = True
                
                If lAlmacen Then
                    cmdEgreso(3).Enabled = True
                End If
                
                cmdEgreso_Click (1)
          Case Is = 1 'Grabar
               'Chequea Datos
               Dim nCorrela As String
               Dim nPos As Integer
               If txtMonto.Text = "" Then MsgBox "Ingrese el Monto del Recibo", vbExclamation, sMensaje: txtMonto.SetFocus: Exit Sub
               If txtDescripcion.Text = "" Then MsgBox "Ingrese la Descripción del Recibo", vbExclamation, sMensaje: Exit Sub
               If ModoEgreso <> "MERCADERIA" Then
                If cboTipoEgreso.Text = "" Then MsgBox "Seleccione un Tipo de Egreso", vbExclamation, sMensaje: Exit Sub
               End If

               If vwNC Then
                If Trim(txtnotacredito.Text) = "" Then MsgBox "Seleccione una nota de credito", vbExclamation, sMensaje: Exit Sub
               End If
                
               If Trim(sMonedaRecibo) = "" Then MsgBox "Falta elegir una Moneda!!!": Exit Sub
               
               If Sw Then

                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tRecibo) as Codigo from MEGRESO where substring(tRecibo,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                  If IsNull(nCorrela) Or Mid(nCorrela, 1, 2) <> Mid(Trim(str(Year(Date))), 3, 2) Then
                      txtRecibo.Text = Mid(Trim(str(Year(Date))), 3, 2) & "00000001"
                  Else
                      txtRecibo.Text = Mid(Trim(str(Year(Date))), 3, 2) & Lib.Correlativo(Mid(nCorrela, 3, 8), 8)
                  End If
                  Sw = False
                   
                   If vwNC = False Then
                    'Cambiar el SQL
                    Isql = "insert into MEGRESO( " & _
                           "tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro,fDiaContable,ttipoEgreso,dReferencia,TipoIdentificacion,Identificacion,RazonSocial,NComprobante,TipoComprobante,CentroCosto,tModoEgreso, tObservacion) " & _
                           "values ('" & txtRecibo.Text & "', " & _
                                  " '" & sCaja & "', " & _
                                  " '" & txtTurno.Text & "', " & _
                                  "{fn curdate()}, " & _
                                  "'" & sMonedaRecibo & "', " & _
                                  "" & IIf(sMonedaRecibo = "04", nTC3, IIf(sMonedaRecibo = "03", nTC2, nTC)) & ", " & _
                                  nMonto & ", " & _
                                  " '" & txtDescripcion.Text & "', " & _
                                  " '" & txtAutorizacion.Text & "', " & _
                                  "'01', " & _
                                  "'" & sUsuario & "'," & _
                                  "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "','" & cboTipoEgreso.BoundText & "','" & txtnotacredito.Text & "', " & _
                                  " '" & cboTipoIdentificacion.BoundText & "','" & txtIdentificacion.Text & "','" & txtRSocial.Text & "','" & txtComprobante.Text & "','" & txtTipoComprobante.Text & "','" & cboCentroCosto.BoundText & "','" & ModoEgreso & "','" & txtObservacion.Text & "')"
                    Else
                     'Cambiar el SQL
                    Isql = "insert into MEGRESO( " & _
                           "tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro,fDiaContable,ttipoEgreso,dReferencia,TipoIdentificacion,Identificacion,RazonSocial,NComprobante,TipoComprobante,CentroCosto, tModoEgreso, tObservacion) " & _
                           "values ('" & txtRecibo.Text & "', " & _
                                  " '" & sCaja & "', " & _
                                  " '" & txtTurno.Text & "', " & _
                                  "{fn curdate()}, " & _
                                  "'01'" & ", " & _
                                  nTC & ", " & _
                                  nMonto & ", " & _
                                  " '" & txtDescripcion.Text & "', " & _
                                  " '" & txtAutorizacion.Text & "', " & _
                                  "'01', " & _
                                  "'" & sUsuario & "'," & _
                                  "getdate(),'" & Format(obtieneDiaContable, "yyyyMMdd") & "','','" & txtnotacredito.Text & "', " & _
                                  " '" & cboTipoIdentificacion.BoundText & "','" & txtIdentificacion.Text & "','" & txtRSocial.Text & "','" & txtComprobante.Text & "','" & txtTipoComprobante.Text & "','" & cboCentroCosto.BoundText & "','" & ModoEgreso & "','" & txtObservacion.Text & "')"
                    
                                  
                    End If
                    Cn.Execute Isql
                    
                    If vwNC Then
                        Cn.Execute "update mnotacredito set testadodocumento='02' where tNotaCredito='" & Trim(txtnotacredito.Text) & "'"
                    End If
                                              
                       frmReciboEgreso.RsCabecera.Requery
                       frmReciboEgreso.RsCabecera.Find "tRecibo ='" & txtRecibo.Text & "'"
                       'RsDetalle.MoveLast
                       MsgBox "Registro Guardado", vbInformation, sMensaje
                       ActivarBotones (True)
                       cmdTexto.Caption = "Registro " & IIf(frmReciboEgreso.RsCabecera.RecordCount = 0, 0, frmReciboEgreso.RsCabecera.AbsolutePosition) & " de " & frmReciboEgreso.RsCabecera.RecordCount
                       
                       
                       Isql = "select tRecibo, tTurno, fFecha, tMoneda, nMonto, tDescripcion, tAutoriza, tEstadoDocumento, tUsuario, fRegistro, Estado, Moneda, TipoEgreso,ttipoegreso, str(nMonto,10,2) as xMonto, convert(varchar, fFecha, 103) as xFecha, tObservacion as tObservacion  from vEgreso " & _
                               "where tRecibo = '" & txtRecibo.Text & "'"
                               
                       Set RsEgreso = Lib.OpenRecordset(Isql, Cn)
                                        
                       ImprimeReciboEgreso RsEgreso
                       Unload Me
               Else
               
                    'JESUS
                    If frmReciboEgreso.RsCabecera!tTurno = sTurno Then
                       'Password

                    Else
                       'Password
                       If Supervisor("06") = False Then
                          MsgBox "Clave no permitida", vbExclamation, sMensaje
                          Exit Sub
                       End If
                    End If
                    '----------Fin 24.10.2017---------
                    
                    
                  'Cambiar el SQL
                  If vwNC Then
                    Isql = "update MEGRESO set " & _
                           "tMoneda ='01', " & _
                           "nMonto =" & nMonto & ", " & _
                           "tDescripcion ='" & txtDescripcion.Text & "', " & _
                           "tTipoEgreso='" & cboTipoEgreso.BoundText & "', " & _
                           "tAutoriza ='" & txtAutorizacion.Text & "', lReplica=1 ," & _
                           "TipoIdentificacion ='" & cboTipoIdentificacion.BoundText & "', " & _
                           "Identificacion ='" & txtIdentificacion.Text & "', " & _
                           "RazonSocial ='" & txtRSocial.Text & "', " & _
                           "NComprobante ='" & txtComprobante.Text & "', " & _
                           "TipoComprobante ='" & txtTipoComprobante.Text & "', " & _
                           "CentroCosto ='" & cboCentroCosto.BoundText & "' " & _
                           "tObservacion='" & txtObservacion.Text & "'" & _
                           " where tRecibo = '" & txtRecibo & "'"
                         ',TipoIdentificacion,Identificacion,RazonSocial,NComprobante,TipoComprobante,CentroCosto
                    Else
                        Isql = "update MEGRESO set " & _
                                "tMoneda ='" & sMonedaRecibo & "', " & _
                                "nMonto =" & nMonto & ", " & _
                                "tDescripcion ='" & txtDescripcion.Text & "', " & _
                                "nTipoCambio=" & IIf(sMonedaRecibo = "04", nTC3, IIf(sMonedaRecibo = "03", nTC2, nTC)) & ", " & _
                                "tTipoEgreso='" & cboTipoEgreso.BoundText & "', " & _
                                "tAutoriza ='" & txtAutorizacion.Text & "', lReplica=1, " & _
                                "TipoIdentificacion ='" & cboTipoIdentificacion.BoundText & "', " & _
                                "Identificacion ='" & txtIdentificacion.Text & "', " & _
                                "RazonSocial ='" & txtRSocial.Text & "', " & _
                                "NComprobante ='" & txtComprobante.Text & "', " & _
                                "TipoComprobante ='" & txtTipoComprobante.Text & "', " & _
                                "CentroCosto ='" & cboCentroCosto.BoundText & "' " & _
                                "tObservacion='" & txtObservacion.Text & "'" & _
                                " where tRecibo = '" & txtRecibo & "'"
                    End If
                       
                   Cn.Execute Isql
                   
                   'Cambiar el Nombre del Formulario
                   nPos = frmReciboEgreso.RsCabecera.AbsolutePosition
                   frmReciboEgreso.RsCabecera.Requery
                   frmReciboEgreso.RsCabecera.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               'ImprimeReciboEgreso frmReciboEgreso.RsCabecera
          
          Case Is = 2 'Eliminar
               If frmReciboEgreso.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If frmReciboEgreso.RsCabecera!tTurno = sTurno Then
                  'Password
                  If Supervisor("07") = False Then
                     MsgBox "Clave no permitida", vbExclamation, sMensaje
                     Exit Sub
                  End If
               Else
                  'Password
                  If Supervisor("06") = False Then
                     MsgBox "Clave no permitida", vbExclamation, sMensaje
                     Exit Sub
                  End If
               End If
                              
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Recibo Nro." & txtRecibo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "update MEGRESO set tEstadoDocumento = '04' where tRecibo = '" & txtRecibo & "'"
                If ModoEgreso = "MERCADERIA" Then
                        Cn.Execute "update mnotacredito set testadodocumento='05' where tNotaCredito='" & Trim(txtnotacredito.Text) & "'"
                End If
               nPos = frmReciboEgreso.RsCabecera.AbsolutePosition
               frmReciboEgreso.RsCabecera.Requery
               frmReciboEgreso.RsCabecera.AbsolutePosition = nPos
               Asignar
               
          Case Is = 3 'Salir
               Unload Me
          
          Case Is = 4 'KB Descripcion
                frmKeyBoard.txtResultado.Text = txtDescripcion.Text
                frmKeyBoard.Show vbModal
                txtDescripcion.Text = IIf(wEnter, sDescrip, txtDescripcion.Text)
               
          Case Is = 5 'KB Autorizacion
                frmKeyBoard.txtResultado.Text = txtAutorizacion.Text
                frmKeyBoard.Show vbModal
                txtAutorizacion.Text = IIf(wEnter, sDescrip, txtAutorizacion.Text)
        
          Case Is = 11 'KB Observacion
                frmKeyBoard.txtResultado.Text = txtObservacion.Text
                frmKeyBoard.txtResultado.MaxLength = 0
                frmKeyBoard.Show vbModal
                txtObservacion.Text = IIf(wEnter, sDescrip, txtObservacion.Text)
          
          Case Is = 6 'Numpad Soles
               frmNumPad.Show vbModal
               nMonto = IIf(wEnter, val(sDescrip), nMonto)
               txtMonto.Text = Format(nMonto, "###,##0.00")
               cmdOpcion(6).FontBold = True
               cmdOpcion(7).FontBold = False
               cmdOpcion(9).FontBold = False
               cmdOpcion(10).FontBold = False
               sMonedaRecibo = "01"
               
          Case Is = 7 'Numpad Dolares
               frmNumPad.Show vbModal
               nMonto = IIf(wEnter, val(sDescrip), nMonto)
               txtMonto.Text = Format(nMonto, "###,##0.00")
               cmdOpcion(6).FontBold = False
               cmdOpcion(7).FontBold = True
               cmdOpcion(9).FontBold = False
               cmdOpcion(10).FontBold = False
               sMonedaRecibo = "02"
               
          Case Is = 9 'Numpad ME2
               frmNumPad.Show vbModal
               nMonto = IIf(wEnter, val(sDescrip), nMonto)
               txtMonto.Text = Format(nMonto, "###,##0.00")
               cmdOpcion(6).FontBold = False
               cmdOpcion(7).FontBold = False
               cmdOpcion(9).FontBold = True
               cmdOpcion(10).FontBold = False
               sMonedaRecibo = "03"
               
          Case Is = 10 'Numpad ME3
               frmNumPad.Show vbModal
               nMonto = IIf(wEnter, val(sDescrip), nMonto)
               txtMonto.Text = Format(nMonto, "###,##0.00")
               cmdOpcion(6).FontBold = False
               cmdOpcion(7).FontBold = False
               cmdOpcion(9).FontBold = False
               cmdOpcion(10).FontBold = True
               sMonedaRecibo = "04"
               
        
               
         Case Is = 8 'NOTACREDITO
          
               vwNC = True
               Isql = "select *, tNotaCredito as Codigo, nVenta as Descripcion from vNotaCredito where tEstadodocumento ='05'"
               Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Nota Credito", 2, "tNotaCredito", 1900, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1200, 0, 0, "dd/mm/yyyy", _
                                                               "Cliente", 2, "Cliente", 3900, 0, 0, "", _
                                                               "Monto", 2, "nVenta", 1200, 1, 0, "##,##0.00")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               If wEnter Then
                txtnotacredito.Text = sCodigo
                txtMonto.Text = Format(CDbl(sDescrip), "###,##0.00")
                nMonto = CDbl(sDescrip)
                txtMonto.Enabled = False
                cmdOpcion(7).Enabled = False
                cmdOpcion(6).Enabled = False
                cmdOpcion(9).Enabled = False
                cmdOpcion(10).Enabled = False

               Else
                'vwNC = False
                  Exit Sub
               
               End If
   End Select
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Centrar Me
   LlenaCombo
   LlenaComboTipoIdentificacion
   LlenaComboCentroCosto
   
   If sMonE = "" Or sMonN = sMonE Then
      cmdOpcion(6).Caption = sMonedaN
      cmdOpcion(7).Visible = False
   Else
      cmdOpcion(6).Caption = sMonedaN
      cmdOpcion(7).Caption = sMonedaE
   End If
   
   If pais = "003" Then
        If Trim(sMonE2) <> "" Then
            cmdOpcion(9).Caption = sMonedaE2
            cmdOpcion(9).Visible = True
        End If
        If Trim(sMonE3) <> "" Then
            cmdOpcion(10).Caption = sMonedaE3
            cmdOpcion(10).Visible = True
        End If
   End If
   
   cmdOpcion(6).FontBold = True
   sMonedaRecibo = "01"
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Recibos de Egreso "
'    fraDetalle.Caption = Me.Caption
    
   
        
    If Sw = True Then
        cmdOpcion_Click (0)
'       ActivarBotones (False)
'       Blanquear Me
'       txtTurno.Text = sTurno
'       txtFecha.Text = Date
'       txtMonto.Text = "0.00"
'       lblEstado.Caption = ""
'       cmdOpcion(1).Enabled = True
       
    Else
       Asignar
       ActivarBotones (True)
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmReciboEgreso.RsCabecera.RecordCount = 0, 0, frmReciboEgreso.RsCabecera.AbsolutePosition) & " de " & frmReciboEgreso.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault

   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmReciboEgresoDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
End Sub

Public Sub Bloquear(Activo As Boolean)
   cmdOpcion(1).Enabled = Not Activo
   cmdOpcion(2).Enabled = Not Activo
   cmdOpcion(4).Enabled = Not Activo
   cmdOpcion(5).Enabled = Not Activo
   cmdOpcion(6).Enabled = Not Activo
   cmdOpcion(7).Enabled = Not Activo
   cmdOpcion(9).Enabled = Not Activo
   cmdOpcion(10).Enabled = Not Activo
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
End Function
'diaContable

