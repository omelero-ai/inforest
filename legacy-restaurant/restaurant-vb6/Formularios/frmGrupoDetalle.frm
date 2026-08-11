VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGrupoDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   -570
   ClientTop       =   1875
   ClientWidth     =   10785
   Icon            =   "frmGrupoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10785
   Begin VB.Frame fraDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   1680
      TabIndex        =   32
      Top             =   30
      Width           =   9060
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   1920
         MaxLength       =   24
         TabIndex        =   2
         Text            =   " "
         Top             =   1020
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   630
         Width           =   5070
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   1440
         Width           =   885
      End
      Begin VB.Frame Frame1 
         Caption         =   " Botonera Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   7110
         TabIndex        =   46
         Top             =   195
         Width           =   1815
         Begin VB.CommandButton cmdNAG 
            BackColor       =   &H00C0C0C0&
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
            Left            =   1230
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   870
            Width           =   510
         End
         Begin VB.CommandButton cmdBotonG 
            BackColor       =   &H00C0C0C0&
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
            Left            =   690
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   870
            Width           =   510
         End
         Begin VB.CommandButton cmdBotonG 
            BackColor       =   &H00C0C0C0&
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
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   870
            Width           =   510
         End
         Begin VB.CommandButton cmdBotonG 
            BackColor       =   &H00C0C0C0&
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
            Left            =   1230
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   300
            Width           =   510
         End
         Begin VB.CommandButton cmdBotonG 
            BackColor       =   &H00C0C0C0&
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
            Left            =   690
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   300
            Width           =   510
         End
         Begin VB.CommandButton cmdBotonG 
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
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   300
            Width           =   510
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Boton :"
            Height          =   195
            Left            =   135
            TabIndex        =   54
            Top             =   1485
            Width           =   510
         End
         Begin VB.Label txtBotonG 
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
            Left            =   705
            TabIndex        =   53
            Top             =   1440
            Width           =   990
         End
      End
      Begin VB.Frame fraBotons 
         Caption         =   " Bot. SubGrupo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3120
         Left            =   7110
         TabIndex        =   35
         Top             =   2115
         Width           =   1845
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
            Left            =   652
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1620
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1620
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
            Left            =   1215
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   990
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
            Left            =   652
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   990
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   990
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
            TabIndex        =   38
            Top             =   330
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
            Left            =   652
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   330
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   330
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Boton :"
            Height          =   195
            Left            =   90
            TabIndex        =   45
            Top             =   2520
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
            Left            =   675
            TabIndex        =   44
            Top             =   2475
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdOpcionGrilla 
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
         Left            =   750
         Picture         =   "frmGrupoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5685
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcionGrilla 
         Caption         =   "Modificar"
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
         Left            =   1980
         Picture         =   "frmGrupoDetalle.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5685
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcionGrilla 
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
         Left            =   3210
         Picture         =   "frmGrupoDetalle.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5685
         Width           =   1170
      End
      Begin VB.Frame fraGrilla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   90
         TabIndex        =   33
         Top             =   1680
         Width           =   6975
         Begin VB.TextBox txtSubGrupoPatio 
            Height          =   285
            Left            =   1845
            MaxLength       =   30
            TabIndex        =   72
            Top             =   1320
            Width           =   4935
         End
         Begin VB.TextBox txtOrdenPatio 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6000
            MaxLength       =   2
            TabIndex        =   70
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.CheckBox chkDesplegaGrupo 
            Alignment       =   1  'Right Justify
            Caption         =   "Aparece expandido en Integración"
            Height          =   435
            Left            =   4680
            TabIndex        =   69
            Top             =   2040
            Width           =   2115
         End
         Begin VB.CheckBox chkPrecioTrans 
            Alignment       =   1  'Right Justify
            Caption         =   "Los Precios de venta seran los de Transferencia (Almacen) :"
            Height          =   315
            Left            =   2160
            TabIndex        =   68
            Top             =   3360
            Visible         =   0   'False
            Width           =   4605
         End
         Begin VB.TextBox txtCuenta 
            Height          =   285
            Left            =   1845
            MaxLength       =   15
            TabIndex        =   66
            Top             =   1680
            Width           =   2595
         End
         Begin VB.TextBox txtCriterio 
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
            Left            =   1845
            MaxLength       =   24
            TabIndex        =   14
            Text            =   " "
            Top             =   2700
            Width           =   5025
         End
         Begin VB.TextBox txtOrden 
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
            Left            =   1845
            MaxLength       =   24
            TabIndex        =   13
            Text            =   " "
            Top             =   2370
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcionGrilla 
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
            Index           =   3
            Left            =   4410
            Picture         =   "frmGrupoDetalle.frx":0B78
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   4000
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcionGrilla 
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
            Index           =   4
            Left            =   5640
            Picture         =   "frmGrupoDetalle.frx":10AA
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4000
            Width           =   1170
         End
         Begin VB.CheckBox chkImpuesto3 
            Alignment       =   1  'Right Justify
            Caption         =   "Impuesto 3 :"
            Height          =   195
            Left            =   4680
            TabIndex        =   17
            Top             =   3030
            Width           =   2085
         End
         Begin VB.CheckBox chkImpuesto2 
            Alignment       =   1  'Right Justify
            Caption         =   "Impuesto 2 :"
            Height          =   195
            Left            =   2280
            TabIndex        =   16
            Top             =   3030
            Width           =   2085
         End
         Begin VB.CheckBox chkImpuesto1 
            Alignment       =   1  'Right Justify
            Caption         =   "Impuesto 1 :"
            Height          =   195
            Left            =   840
            TabIndex        =   15
            Top             =   3030
            Width           =   1200
         End
         Begin VB.CheckBox chkArea 
            Alignment       =   1  'Right Justify
            Caption         =   "Imprime Area :"
            Height          =   195
            Left            =   3240
            TabIndex        =   12
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtDetalladoS 
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
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   9
            Text            =   " "
            Top             =   605
            Width           =   4890
         End
         Begin VB.TextBox txtResumidoS 
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
            Left            =   1845
            MaxLength       =   24
            TabIndex        =   10
            Text            =   " "
            Top             =   940
            Width           =   2595
         End
         Begin VB.TextBox txtSubGrupo 
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
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   270
            Width           =   1170
         End
         Begin VB.CheckBox chkActivoS 
            Alignment       =   1  'Right Justify
            Caption         =   "Activo :"
            Height          =   195
            Left            =   1155
            TabIndex        =   18
            Top             =   3350
            Width           =   885
         End
         Begin VB.Frame Frame3 
            Height          =   30
            Left            =   90
            TabIndex        =   34
            Top             =   3705
            Width           =   6765
         End
         Begin MSDataListLib.DataCombo cboArea 
            Height          =   315
            Left            =   1845
            TabIndex        =   11
            Top             =   2040
            Width           =   2610
            _ExtentX        =   4604
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
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Patio :"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   73
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Orden Integración:"
            Height          =   195
            Left            =   4680
            TabIndex        =   71
            Top             =   1720
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Contable (70) :"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Criterio de Agrupación :"
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   65
            Top             =   2745
            Width           =   1650
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Orden de Impresión :"
            Height          =   195
            Index           =   8
            Left            =   270
            TabIndex        =   64
            Top             =   2415
            Width           =   1470
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Index           =   3
            Left            =   1155
            TabIndex        =   63
            Top             =   315
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Resumida :"
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   62
            Top             =   985
            Width           =   1680
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción Detallada :"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   61
            Top             =   650
            Width           =   1650
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Area Producción :"
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   60
            Top             =   2040
            Width           =   1275
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   2805
         Left            =   120
         TabIndex        =   59
         Top             =   2340
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   4948
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1215
         TabIndex        =   57
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   1065
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   55
         Top             =   675
         Width           =   1650
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10725
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6585
      Width           =   10785
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   645
         Left            =   30
         ScaleHeight     =   585
         ScaleWidth      =   5820
         TabIndex        =   25
         Top             =   60
         Width           =   5880
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4380
            Picture         =   "frmGrupoDetalle.frx":11AC
            Style           =   1  'Graphical
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4860
            Picture         =   "frmGrupoDetalle.frx":16EE
            Style           =   1  'Graphical
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5340
            Picture         =   "frmGrupoDetalle.frx":1C30
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmGrupoDetalle.frx":2172
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmGrupoDetalle.frx":26B4
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmGrupoDetalle.frx":2BF6
            Style           =   1  'Graphical
            TabIndex        =   26
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
            Left            =   1500
            TabIndex        =   58
            Top             =   150
            Width           =   2775
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
         Left            =   7170
         Picture         =   "frmGrupoDetalle.frx":3138
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   6000
         Picture         =   "frmGrupoDetalle.frx":366A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   8340
         Picture         =   "frmGrupoDetalle.frx":3B9C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   9510
         Picture         =   "frmGrupoDetalle.frx":3C9E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   6390
      Left            =   30
      Picture         =   "frmGrupoDetalle.frx":3D90
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1620
   End
End
Attribute VB_Name = "frmGrupoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsArea   As Recordset
Dim RsbotonG As Recordset
Dim RsBoton  As Recordset
Dim RsGrilla As Recordset
Dim nCorrela As String
Dim wAgrega  As Boolean
Dim nOrden   As Integer

Sub LlenaCombos()
    With cboArea
         Isql = "Select * from vArea where lActivo = 1 order by Codigo"
         Set RsArea = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsArea
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub
Sub Asignar()
    With frmGrupo.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        
        'CESAR CUENTA CONTABLE quitar ya no pertenece a esa tabla
        'txtCuenta = IIf(IsNull(!tcuentaContable), "", !tcuentaContable)
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
    
    'Cambiar el Filtro
    RsGrilla.Filter = "tGrupo ='" & txtCodigo.Text & "'"
    BotoneraG
    Botonera
End Sub

'Private Sub chkPrecioTrans_Click()
'    If chkPrecioTrans.value = 1 Then
'        If Calcular("select COUNT(*) AS CODIGO from TPRODUCTO where tDescargo<>'D' AND TSUBGRUPO='" & Trim(txtSubGrupo.Text) & "'", Cn) > 0 Then
'            MsgBox "Hay Productos sin enlace Directo a Almacen!!"
'            chkPrecioTrans.value = False
'        End If
'    End If
'End Sub



Private Sub cmdBoton_Click(Index As Integer)
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
   End If
   cmdBoton(Index).backColor = vbRed
   cmdBoton(Index).Enabled = False
   txtBoton.Caption = Index
End Sub

Private Sub cmdBotonG_Click(Index As Integer)
   If val(txtBotonG) <> 0 Then
      cmdBotonG(val(txtBotonG)).backColor = vbButtonFace
      cmdBotonG(val(txtBotonG)).Enabled = True
   End If
   cmdBotonG(Index).backColor = vbRed
   cmdBotonG(Index).Enabled = False
   txtBotonG.Caption = Index
End Sub

Private Sub cmdNA_Click()
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
      txtBoton.Caption = "NA"
   End If
End Sub

Private Sub cmdNAG_Click()
   If val(txtBotonG) <> 0 Then
      cmdBotonG(val(txtBotonG)).backColor = vbButtonFace
      cmdBotonG(val(txtBotonG)).Enabled = True
      txtBotonG.Caption = "NA"
   End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmGrupo.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmGrupo.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmGrupo.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmGrupo.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmGrupo.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmGrupo.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmGrupo.RsCabecera.AbsolutePosition & " de " & frmGrupo.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.Value = 1
               chkArea.Value = 1
                
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
               BotoneraG
               Botonera
         
               cmdOpcionGrilla(0).Enabled = False
               cmdOpcionGrilla(1).Enabled = False
               cmdOpcionGrilla(2).Enabled = False
               RsGrilla.Filter = "Codigo ='" & txtCodigo.Text & "'"
               
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
          
          Case Is = 1 ' Grabar
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               'If Val(txtOrden.Text) = 0 Then MsgBox "Orden de impresion no puede ser con valor 0", vbExclamation, sMensaje: txtOrden.SetFocus: Exit Sub
               
               If Sw Then
                  'Verifica el Orden
                  'If Calcular("select nOrden as Codigo from tGrupo where nOrden='" & Val(txtOrden.Text) & "'", Cn) <> "0" Then
                  '   MsgBox "Numero de orden de impresion ya asignado verifique", vbExclamation, sMensaje
                  '   txtOrden.SetFocus
                  '   Exit Sub
                  'End If
               
                  'Asignar El Campo de Codificación
                  nCorrela = Calcular("select max(tCodigoGrupo) as Codigo from TGRUPO", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     txtCodigo = "01"
                  Else
                     txtCodigo = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  'CESAR CUENTA CONTABLE
                  Isql = "insert into TGRUPO( " & _
                         "tCodigoGrupo, tDetallado, tResumido, nBoton, lActivo) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                val(txtBotonG.Caption) & ", " & _
                                       chkActivo.Value & ") "
                                
           
                  Cn.Execute Isql
                  frmGrupo.RsCabecera.Requery
                  frmGrupo.RsCabecera.MoveLast
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmGrupo.RsCabecera.RecordCount = 0, 0, frmGrupo.RsCabecera.AbsolutePosition) & " de " & frmGrupo.RsCabecera.RecordCount
                  'CESAR CUENTA CONTABLE
               Else
                  'Cambiar el SQL
                  'CESAR CUENTA CONTABLE
                  Isql = "update TGRUPO set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nBoton =" & val(txtBotonG.Caption) & ", " & _
                         "lActivo =" & chkActivo.Value & _
                         ",lreplica=1  where tCodigoGrupo = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmGrupo.RsCabecera.Bookmark
                   frmGrupo.RsCabecera.Requery
                   If frmGrupo.RsCabecera.RecordCount = 0 Then
                      frmGrupo.RsCabecera.Filter = adFilterNone
                   End If
                   frmGrupo.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               
               Cn.Execute "update TGRUPO set lNuevoModificado=1 where tCodigoGrupo = '" & txtCodigo & "'"
               
               RsbotonG.Requery
               cmdOpcionGrilla(0).Enabled = True
               cmdOpcionGrilla(1).Enabled = True
               cmdOpcionGrilla(2).Enabled = True
                                                
          Case Is = 2 ' Eliminar
               If frmGrupo.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
            
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Grupo " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                                    
               If Calcular("select count(tCodigoGrupo) as Codigo From DPEDIDO where tCodigoGrupo='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este Grupo, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
                  
               'Cambia el Delete
               Cn.Execute "delete from TGRUPO where  tCodigoGrupo = '" & txtCodigo.Text & "'"
               Cn.Execute "delete from TSUBGRUPO where tCodigoGrupo ='" & txtCodigo.Text & "'"
               frmGrupo.RsCabecera.Requery
               
               If frmGrupo.RsCabecera.RecordCount <> 0 Then
                  frmGrupo.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmGrupo.RsCabecera.RecordCount = 0, 0, frmGrupo.RsCabecera.AbsolutePosition) & " de " & frmGrupo.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
                                   
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub cmdOpcionGrilla_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               'Cambiar los Controles
               With RsGrilla
                    'Cuadro de Texto
                    txtSubGrupo.Text = ""
                    txtDetalladoS.Text = ""
                    txtResumidoS.Text = ""
                    'CESAR CUENTA CONTABLE
                    txtCuenta.Text = ""
                    'data Como
                    cboArea.BoundText = ""
                    'Check Box
                    chkArea.Value = 1
                    chkActivoS.Value = 1
                    chkImpuesto1.Value = 1
                    chkImpuesto2.Value = 1
                    chkImpuesto3.Value = 1
                End With
                
                SubDetalle False
                wAgrega = True
                txtDetalladoS.SetFocus
          
          Case Is = 1 ' Modificar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               SubDetalle False
               wAgrega = False
               SubAsignar
               txtDetalladoS.SetFocus
               
          Case Is = 2 ' Eliminar
               If RsGrilla.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               SubAsignar
               If MsgBox("Seguro de Eliminar el Sub Grupo " & txtSubGrupo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If Calcular("select count(tCodigoSubGrupo) as Codigo From DPEDIDO where tCodigoSubGrupo='" & txtSubGrupo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este Subgrupo, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TSUBGRUPO where tCodigosubGrupo ='" & txtSubGrupo.Text & "'"
               frmGrupo.RsCabecera.Requery
               RsGrilla.Requery
               
               If RsGrilla.RecordCount <> 0 Then
                  RsGrilla.MoveLast
               End If
          
          Case Is = 3 ' Grabar
                
                If Trim(txtDetalladoS.Text) = "" Then
                    MsgBox "No ha ingresado un Nombre al Subgrupo!!", vbInformation
                    txtDetalladoS.SetFocus
                    Exit Sub
                End If

                
                If chkPrecioTrans.Value = 1 And lAlmacen Then ' consulta de transferencia de Precios de Venta Almacen.
                    If MsgBox("Se actualizaran los Precios de Venta de Los Productos con Enlace Directo Almacen!!! ¿Desea Continuar?", vbQuestion + vbYesNo, sMensaje) = vbNo Then
                        Exit Sub
                    End If
                End If

                  If Calcular("select count(*) as codigo from vsubgrupo where tOrdenPatio=" & IIf(Trim(Me.txtOrdenPatio.Text) = "", 0, Me.txtOrdenPatio.Text) & " and tordenPatio<>0 and Codigo<>'" & Me.txtSubGrupo.Text & "'", Cn) > 0 Then
                    MsgBox "El orden para patio de comidas ya esta asignado!!!", vbInformation
                    Exit Sub
                  End If
                  
               If wAgrega Then
                  'Cambiar el SQL Agregar Detalle

                  nCorrela = Mid(Calcular("select max(tCodigoSubGrupo) as Codigo from TSUBGRUPO where tCodigoGrupo ='" & txtCodigo.Text & "'", Cn), 3, 2)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     txtSubGrupo = Trim(txtCodigo.Text) & "01"
                  Else
                     txtSubGrupo = Trim(txtCodigo.Text) & Lib.Correlativo(nCorrela, 2)
                  End If

                  'CESAR CUENTA CONTABLE
                  Isql = "insert into TSUBGRUPO( " & _
                         "tCodigoGrupo, tCodigoSubGrupo, tDetallado, tResumido, tCuentaContable, tArea, lImprimeArea, lImpuesto1, lImpuesto2, lImpuesto3, nBoton, nOrden, tAgrupacion, lActivo, tUsuario, fRegistro,lTransVen, lDesplieguePatio, tOrdenPatio) " & _
                         "values (  '" & txtCodigo.Text & "', " & _
                                  " '" & txtSubGrupo.Text & "', " & _
                                  " '" & txtDetalladoS.Text & "', " & _
                                  " '" & txtResumidoS.Text & "', " & _
                                  " '" & txtCuenta.Text & "', " & _
                                  " '" & cboArea.BoundText & "', " & _
                                         chkArea.Value & ", " & _
                                         chkImpuesto1.Value & ", " & _
                                         chkImpuesto2.Value & ", " & _
                                         chkImpuesto3.Value & ", " & _
                                         val(txtBoton.Caption) & ", " & _
                                         val(txtOrden.Text) & ", " & _
                                  " '" & txtCriterio.Text & "', " & _
                                         chkActivoS.Value & ", " & _
                                  " '" & sUsuario & "', " & _
                                  " getdate() ," & _
                                  chkPrecioTrans.Value & "," & Me.chkDesplegaGrupo.Value & ", " & IIf(Trim(Me.txtOrdenPatio.Text) = "", 0, Me.txtOrdenPatio.Text) & ")"

                   Cn.Execute Isql
                   'Cambiar el Filtro de la Grilla
'                   RsGrilla.Filter = "tGrupo ='" & txtCodigo.Text & "'"
'                   RsGrilla.Requery
'                   RsGrilla.MoveLast
'                   RsBoton.Requery
                   MsgBox "Registro Agregado", vbInformation, sMensaje
                Else
                   ''Cambiar el SQL Modificar Detalle
                   'CESAR CUENTA CONTABLE
                   Isql = "update TSUBGRUPO set " & _
                          "tDetallado ='" & txtDetalladoS.Text & "', " & _
                          "tResumido ='" & txtResumidoS.Text & "', " & _
                          "tArea ='" & cboArea.BoundText & "', " & _
                          "lImprimeArea =" & chkArea.Value & ", " & _
                          "lImpuesto1 =" & chkImpuesto1.Value & ", " & _
                          "lImpuesto2 =" & chkImpuesto2.Value & ", " & _
                          "lImpuesto3 =" & chkImpuesto3.Value & ", " & _
                          "nBoton =" & val(txtBoton.Caption) & ", " & _
                          "tAgrupacion ='" & txtCriterio.Text & "', " & _
                          "nOrden =" & val(txtOrden.Text) & ", " & _
                          "lActivo =" & chkActivoS.Value & " " & _
                          ",lreplica=1, " & _
                          "tCuentaContable ='" & txtCuenta.Text & "' , lTransVen = " & chkPrecioTrans.Value & " , lDesplieguePatio=" & Me.chkDesplegaGrupo.Value & ", tOrdenPatio=" & IIf(Trim(Me.txtOrdenPatio.Text) = "", 0, Me.txtOrdenPatio.Text) & "  where tCodigoGrupo = '" & txtCodigo.Text & "' and tCodigoSubGrupo = '" & txtSubGrupo.Text & "' "

                    Cn.Execute Isql
'                    nPos = RsGrilla.AbsolutePosition
'                    RsBoton.Requery
'                    RsGrilla.Requery
'                    RsGrilla.AbsolutePosition = nPos
                    MsgBox "Registro Modificado", vbInformation, sMensaje
                End If
                
                Cn.Execute " update TSUBGRUPO set lNuevoModificado=1, tSubgrupoPatio='" & Me.txtSubGrupoPatio.Text & "' where tCodigoGrupo = '" & txtCodigo.Text & "' and tCodigoSubGrupo = '" & txtSubGrupo.Text & "' "
                
               
                
                nPos = RsGrilla.AbsolutePosition
                RsBoton.Requery
                RsGrilla.Requery
                 RsGrilla.Filter = "tGrupo ='" & txtCodigo.Text & "'"
                If nPos > 0 Then
                    RsGrilla.AbsolutePosition = nPos
                End If
                If chkPrecioTrans.Value = 1 And lAlmacen Then  ' actualiza los precios de Venta de almacen
                    Cn.Execute "exec Usp_ActualizarPreciosTransferenciaAlmacen '" & Trim(txtSubGrupo.Text) & "' , '" & sAlmacenMDB & "',1"
                End If
                
                SubDetalle True
                txtSubGrupo.Text = ""
                Botonera
          
          Case Is = 4 ' Cancelar
               SubDetalle True
               txtSubGrupo.Text = ""
               Botonera
               
   End Select

End Sub

Private Sub Form_Load()


    Screen.MousePointer = vbHourglass
    Centrar Me
    fraGrilla.Visible = False

    'Ingrese el Titulo
    Me.Caption = " Mantenimiento de Grupos "
    fraDetalle.Caption = Me.Caption
    
    If sImpuesto1 = "" Then
       chkImpuesto1.Visible = False
    Else
       chkImpuesto1.Caption = sImpuesto1 & " : "
       chkImpuesto1.Visible = True
    End If
    
    If sImpuesto2 = "" Then
       chkImpuesto2.Visible = False
    Else
       chkImpuesto2.Caption = sImpuesto2 & " : "
       chkImpuesto2.Visible = True
    End If
    
    If sImpuesto3 = "" Then
       chkImpuesto3.Visible = False
    Else
       chkImpuesto3.Caption = sImpuesto3 & " : "
       chkImpuesto3.Visible = True
    End If
        
    If lAlmacen Then
        chkPrecioTrans.Visible = True
        chkPrecioTrans.Value = 0
    Else
        chkPrecioTrans.Visible = False
        chkPrecioTrans.Value = 0
    End If
        
        
    'Ingrese el SubTitulo
    grdGrilla.Caption = " Mantenimiento de SubGrupos "
    fraGrilla.Caption = grdGrilla.Caption
    
    'Llena todos los Combos
    LlenaCombos
    
    'Botones
    Isql = "select tCodigoGrupo, nBoton from TGRUPO where nBoton > 0 order by tCodigoGrupo"
    Set RsbotonG = Lib.OpenRecordset(Isql, Cn)
    
    Isql = "select tCodigoGrupo, tCodigoSubGrupo, nBoton from TSUBGRUPO where nBoton > 0 order by tCodigoSubGrupo"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)
    
    'Ingresar la Vista de la Grilla
    ''CESAR CUENTA CONTABLE agregar al vsubgrupo el campo tCuentaContable
    Isql = "select * from vSubGrupo "
    Set RsGrilla = Lib.OpenRecordset(Isql, Cn)
    
    Call ConfGrilla(4, grdGrilla, "Código", 2, "Codigo", 800, 2, 0, "", _
                                  "Descripción", 2, "Descripcion", 2800, 0, 0, "", _
                                  "Area", 2, "Area", 1750, 0, 0, "", _
                                  "Imp.Area", 2, "lImprimeArea", 950, 2, 4, "")
    Set grdGrilla.DataSource = RsGrilla
    fraBotons.Enabled = False
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
       chkArea.Value = 1
       'Cambiar Filtro de la Grilla
       RsGrilla.Filter = "Codigo ='" & txtCodigo.Text & "'"
       BotoneraG
       Botonera
       cmdOpcion(0).Enabled = False
       cmdOpcion(2).Enabled = False
       cmdOpcionGrilla(0).Enabled = False
       cmdOpcionGrilla(1).Enabled = False
       cmdOpcionGrilla(2).Enabled = False
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & frmGrupo.RsCabecera.AbsolutePosition & " de " & frmGrupo.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmGrupoDetalle = Nothing
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


Private Sub grdGrilla_DblClick()
   cmdOpcionGrilla_Click (1)
End Sub

Public Sub SubDetalle(Activa As Boolean)
   fraGrilla.Visible = Not Activa
   fraBotons.Enabled = Not Activa
   ActivarBotones Activa
   
   cmdOpcion(1).Enabled = Activa
   cmdOpcion(3).Enabled = Activa
   
   cmdOpcionGrilla(0).Enabled = Activa
   cmdOpcionGrilla(1).Enabled = Activa
   cmdOpcionGrilla(2).Enabled = Activa
         
   'Controles de la cabecera
   txtDetallado.Enabled = Activa
   txtResumido.Enabled = Activa
   chkActivo.Enabled = Activa
   
   
End Sub

Sub SubAsignar()
    'Cambiar la Asignación del Detalle
    With RsGrilla
    
         'CESAR CUENTA CONTABLE
    
         'Cuadro de Texto
         txtSubGrupo.Text = IIf(IsNull(!codigo), "", !codigo)
         txtDetalladoS.Text = IIf(IsNull(!Descripcion), "", !Descripcion)
         txtResumidoS.Text = IIf(IsNull(!tResumido), "", !tResumido)
         txtCuenta.Text = IIf(IsNull(!tcuentaContable), "", !tcuentaContable)
         txtOrden = IIf(IsNull(!nOrden), 0, !nOrden)
         txtCriterio = IIf(IsNull(!tAgrupacion), "", !tAgrupacion)
         
         'Data Combo
         cboArea.BoundText = IIf(IsNull(!tArea), "", !tArea)
        
         'Check
         chkArea.Value = IIf(IsNull(!lImprimearea), 1, IIf(!lImprimearea, 1, 0))
         chkImpuesto1.Value = IIf(IsNull(!lImpuesto1), 1, IIf(!lImpuesto1, 1, 0))
         chkImpuesto2.Value = IIf(IsNull(!lImpuesto2), 1, IIf(!lImpuesto2, 1, 0))
         chkImpuesto3.Value = IIf(IsNull(!lImpuesto3), 1, IIf(!lImpuesto3, 1, 0))
         chkActivoS.Value = IIf(IsNull(!lActivo), 1, IIf(!lActivo, 1, 0))
         chkPrecioTrans.Value = IIf(IsNull(!lTransVen), 1, IIf(!lTransVen, 1, 0))
         Me.chkDesplegaGrupo.Value = IIf(IsNull(!lDesplieguePatio), 1, IIf(!lDesplieguePatio, 1, 0))
         Me.txtOrdenPatio.Text = IIf(IsNull(!tOrdenPatio), "0", !tOrdenPatio)
         Me.txtSubGrupoPatio.Text = IIf(IsNull(!tSubgrupopatio), "", !tSubgrupopatio)
         'chkExcluirPaloteo = IIf(IsNull(!lExcluyePaloteo), 0, IIf(!lExcluyePaloteo = True, 1, 0))
    End With
    Botonera
End Sub

Private Sub BotoneraG()
    Dim i As Integer
    txtBotonG.Caption = "NA"
    If RsbotonG.RecordCount <> 0 Then
        For i = 1 To 5
            RsbotonG.MoveFirst
            RsbotonG.Find ("nBoton=" & i)
            If RsbotonG.EOF Then
               cmdBotonG(i).backColor = vbButtonFace
               cmdBotonG(i).Enabled = True
            Else
               cmdBotonG(i).Enabled = False
               If RsbotonG!tCodigoGrupo = txtCodigo.Text Then
                  txtBotonG.Caption = str(i)
                  cmdBotonG(i).backColor = vbRed
               Else
                  cmdBotonG(i).backColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 5
           cmdBotonG(i).backColor = vbButtonFace
           cmdBotonG(i).Enabled = True
       Next i
    End If
End Sub
Private Sub Botonera()
    Dim i As Integer
    txtBoton.Caption = "NA"
    RsBoton.Filter = "tCodigoGrupo='" & txtCodigo.Text & "'"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 7
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).backColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tCodigoSubGrupo = txtSubGrupo.Text Then
                  txtBoton.Caption = str(i)
                  cmdBoton(i).backColor = vbRed
               Else
                  cmdBoton(i).backColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 7
           cmdBoton(i).backColor = vbButtonFace
           cmdBoton(i).Enabled = True
       Next i
    End If
End Sub

Private Sub txtOrden_GotFocus()
   txtOrden.SelLength = Len(Trim(txtOrden.Text))
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub

Private Sub txtOrdenPatio_Change()
On Error GoTo fin:
    If Me.txtOrdenPatio.Text = "" Then
        Exit Sub
    End If
    If (Not IsNumeric(Me.txtOrdenPatio.Text)) Then
        MsgBox "Valor ingreado no valido!!!"
        Me.txtOrdenPatio.Text = "0"
    End If
    Exit Sub
fin:

End Sub
