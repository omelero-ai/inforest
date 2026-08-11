VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTarjetaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4680
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11070
   Icon            =   "frmTarjetaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   11070
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
      Height          =   3915
      Left            =   9720
      TabIndex        =   22
      Top             =   0
      Width           =   1365
      Begin VB.CommandButton cmdBoton 
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
         Index           =   0
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2760
         Width           =   510
      End
      Begin VB.CommandButton cmdBoton 
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
         TabIndex        =   36
         Top             =   360
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   975
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   975
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1575
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1575
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2190
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2190
         Width           =   510
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boton :"
         Height          =   195
         Left            =   90
         TabIndex        =   38
         Top             =   3555
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
         TabIndex        =   37
         Top             =   3510
         Width           =   540
      End
   End
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
      Height          =   3930
      Left            =   1800
      TabIndex        =   21
      Top             =   0
      Width           =   7980
      Begin VB.Frame fraVisibilidadCanalVenta 
         Caption         =   "Visibilidad Canales de Venta"
         Height          =   2055
         Left            =   4560
         TabIndex        =   46
         Top             =   1800
         Width           =   3255
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal 5"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   52
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal 4"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   51
            Top             =   1440
            Width           =   2655
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal 3"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   50
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal 2"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   49
            Top             =   960
            Width           =   2655
         End
         Begin VB.CheckBox chkCanal 
            Caption         =   "Canal 1"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   48
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox chkVisibilidadCanalVenta 
            Caption         =   "Activar Visibilidad por Canal de Venta"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   270
            Width           =   3015
         End
      End
      Begin VB.TextBox txtCodTarEx 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   41
         Text            =   " "
         Top             =   2760
         Width           =   1230
      End
      Begin VB.CheckBox chkPinPad 
         Alignment       =   1  'Right Justify
         Caption         =   "Utiliza POS :"
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtCuentaContable 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   6
         Text            =   " "
         Top             =   2080
         Width           =   2550
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   6000
         TabIndex        =   8
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox txtTelefono1 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   5
         Text            =   " "
         Top             =   2436
         Width           =   2550
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
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
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
         Left            =   1875
         MaxLength       =   30
         TabIndex        =   2
         Text            =   " "
         Top             =   1012
         Width           =   2550
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
         Left            =   1875
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   656
         Width           =   5025
      End
      Begin VB.TextBox txtFactor 
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
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   4
         Text            =   " 0.00"
         Top             =   1724
         Width           =   1215
      End
      Begin VB.TextBox txtRepresentante 
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
         Left            =   1875
         MaxLength       =   40
         TabIndex        =   3
         Text            =   " "
         Top             =   1368
         Width           =   2550
      End
      Begin MSDataListLib.DataCombo cboTipoIntegracion 
         Height          =   315
         Left            =   4680
         TabIndex        =   44
         Top             =   1320
         Width           =   2280
         _ExtentX        =   4022
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
         Caption         =   "Enlace de Integracion:"
         Height          =   255
         Left            =   4680
         TabIndex        =   45
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cod ApiWeb :"
         Height          =   195
         Left            =   760
         TabIndex        =   42
         Top             =   2800
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable :"
         Height          =   195
         Left            =   495
         TabIndex        =   40
         Top             =   2125
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   701
         Width           =   1650
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   1057
         Width           =   1680
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Factor Retención :"
         Height          =   195
         Left            =   450
         TabIndex        =   26
         Top             =   1769
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1185
         TabIndex        =   25
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Contable Comisión :"
         Height          =   195
         Left            =   45
         TabIndex        =   24
         Top             =   2481
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Representante :"
         Height          =   195
         Left            =   630
         TabIndex        =   23
         Top             =   1413
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11010
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3930
      Width           =   11070
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5250
         TabIndex        =   14
         Top             =   90
         Width           =   5310
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3810
            Picture         =   "frmTarjetaDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4290
            Picture         =   "frmTarjetaDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4770
            Picture         =   "frmTarjetaDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTarjetaDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTarjetaDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTarjetaDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   15
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
            Left            =   1470
            TabIndex        =   39
            Top             =   150
            Width           =   2295
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
         Left            =   7560
         Picture         =   "frmTarjetaDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   6390
         Picture         =   "frmTarjetaDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   8730
         Picture         =   "frmTarjetaDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   9840
         Picture         =   "frmTarjetaDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   3855
      Left            =   60
      Picture         =   "frmTarjetaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1755
   End
End
Attribute VB_Name = "frmTarjetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBoton As Recordset
Dim RsTipoIntegracion As Recordset
Dim rsVisibilidadxCanal As Recordset
Dim rsCanal As Recordset
Private Sub InicializarVisibilidadxCanal()
    Isql = "select tCodigoCanalVenta, tDetallado, lActivo from TCANALVENTA"
    Set rsCanal = Lib.OpenRecordset(Isql, Cn)
    
    Dim datol As String
    
    datol = frmTarjeta.grdGrilla
    
    Dim i As Integer
    For i = 1 To 5
        rsCanal.MoveFirst
        rsCanal.Move i - 1
        chkCanal(i).Caption = rsCanal!tDetallado
    Next i
    
    Dim valorCantidad As Integer
    Isql = "select count(*) as Codigo " & _
       "from TVISIBILIDADTARJETACREDITOXCANAL " & _
       "where tCodigoTarjeta = (" & _
         "select tCodigoTarjeta from TTARJETACREDITO where tDetallado = '" & datol & "'" & _
       " or tCodigoTarjeta = '" & datol & "' )"
    valorCantidad = Calcular(Isql, Cn)
    
    If valorCantidad = 1 Then
        Isql = "select * from TVISIBILIDADTARJETACREDITOXCANAL where tCodigoTarjeta = (" & _
         "select tCodigoTarjeta from TTARJETACREDITO where tDetallado = '" & datol & "'" & _
       "or tCodigoTarjeta = '" & datol & "' )"
        Set rsVisibilidadxCanal = Lib.OpenRecordset(Isql, Cn)
        chkVisibilidadCanalVenta.Value = IIf(IsNull(rsVisibilidadxCanal!lVisibilidad), 0, IIf(rsVisibilidadxCanal!lVisibilidad, 1, 0))
        chkCanal(1).Value = IIf(IsNull(rsVisibilidadxCanal!lCanal1), 0, IIf(rsVisibilidadxCanal!lCanal1, 1, 0))
        chkCanal(2).Value = IIf(IsNull(rsVisibilidadxCanal!lCanal2), 0, IIf(rsVisibilidadxCanal!lCanal2, 1, 0))
        chkCanal(3).Value = IIf(IsNull(rsVisibilidadxCanal!lCanal3), 0, IIf(rsVisibilidadxCanal!lCanal3, 1, 0))
        chkCanal(4).Value = IIf(IsNull(rsVisibilidadxCanal!lCanal4), 0, IIf(rsVisibilidadxCanal!lCanal4, 1, 0))
        chkCanal(5).Value = IIf(IsNull(rsVisibilidadxCanal!lCanal5), 0, IIf(rsVisibilidadxCanal!lCanal5, 1, 0))
    End If
    
End Sub
Private Sub AsignarVisibilidadxCanal()
    Dim valorCantidad As Integer
    Isql = "select count(*) as Codigo from TVISIBILIDADTARJETACREDITOXCANAL where tCodigoTarjeta = '" & txtCodigo.Text & "'"
    valorCantidad = Calcular(Isql, Cn)
    
    If valorCantidad = 1 Then ' Update
        Isql = "update TVISIBILIDADTARJETACREDITOXCANAL set lVisibilidad = " & chkVisibilidadCanalVenta.Value & ", lCanal1 = " & chkCanal(1).Value & ", lCanal2 = " & chkCanal(2).Value & ", lCanal3 = " & chkCanal(3).Value & ", lCanal4 = " & chkCanal(4).Value & ", lCanal5 = " & chkCanal(5).Value & " where tCodigoTarjeta = '" & Trim(txtCodigo.Text) & "'"
        Cn.Execute Isql
    ElseIf valorCantidad = 0 Then ' Insert
        Isql = "insert into TVISIBILIDADTARJETACREDITOXCANAL (lVisibilidad, lCanal1, lCanal2, lCanal3, lCanal4, lCanal5, tCodigoTarjeta) " _
         & "Values (" & chkVisibilidadCanalVenta.Value & ", " & chkCanal(1).Value & "," & chkCanal(2).Value & "," & chkCanal(3).Value & "," & chkCanal(4).Value & "," & chkCanal(5).Value & ",'" & Trim(txtCodigo.Text) & "')"
        Cn.Execute Isql
    End If
    Isql = ""
End Sub
Sub LlenaCombos()
    
    With cboTipoIntegracion
         Isql = "Select * from tintegraciones order by Descripcion"
         Set RsTipoIntegracion = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoIntegracion
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "tCodigo"
    End With
End Sub
Sub Asignar()
    With frmTarjeta.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tCodigoTarjeta), "", !tCodigoTarjeta)
        txtDetallado = IIf(IsNull(!tDetallado), "", !tDetallado)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        txtFactor = Format(IIf(IsNull(!nFactorRetencion), "0.00", !nFactorRetencion), "###,##0.00")
        txtRepresentante = IIf(IsNull(!tRepresentante), "", !tRepresentante)
        txtTelefono1 = IIf(IsNull(!ttelefono1), "", !ttelefono1)
        txtBoton = IIf(IsNull(!nBoton), "", !nBoton)
        txtCuentaContable = IIf(IsNull(!tcuentaContable), "", !tcuentaContable)
        ' se agrego el campo de tCodTarjetaEx
        txtCodTarEx.Text = IIf(IsNull(!tCodTarjetaEx), "", !tCodTarjetaEx)
        cboTipoIntegracion.BoundText = IIf(IsNull(!tcodigoIntegracion), "", !tcodigoIntegracion) 'tCodigoIntegracion""
 
        'Check Box
        chkPinPad = IIf(!lPinPad = True, 1, 0)
        chkActivo = IIf(!lActivo = True, 1, 0)
        Botonera
        If txtBoton = "0" Then
            cmdBoton(0).backColor = vbRed
            cmdBoton(0).Enabled = False
        Else
            cmdBoton(0).backColor = vbButtonFace
            cmdBoton(0).Enabled = True
        End If
    End With
End Sub

Private Sub chkVisibilidadCanalVenta_Click()
    Dim i As Integer
    If (chkVisibilidadCanalVenta.Value = 1) Then
        For i = 1 To 5
            rsCanal.MoveFirst
            rsCanal.Move i - 1
            If rsCanal!lActivo Then
                chkCanal(i).Enabled = True
            End If
        Next i
    Else
     chkCanal(1).Enabled = False
          chkCanal(1).Value = 0
     chkCanal(2).Enabled = False
          chkCanal(2).Value = 0
     chkCanal(3).Enabled = False
          chkCanal(3).Value = 0
     chkCanal(4).Enabled = False
          chkCanal(4).Value = 0
     chkCanal(5).Enabled = False
          chkCanal(5).Value = 0
    End If

End Sub

Private Sub cmdBoton_Click(Index As Integer)
   If val(txtBoton) <> 0 Then
      cmdBoton(val(txtBoton)).backColor = vbButtonFace
      cmdBoton(val(txtBoton)).Enabled = True
   End If
   If txtBoton = "0" Then
      cmdBoton(0).backColor = vbButtonFace
      cmdBoton(0).Enabled = True
   End If
   cmdBoton(Index).backColor = vbRed
   cmdBoton(Index).Enabled = False
   txtBoton.Caption = Index
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTarjeta.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTarjeta.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTarjeta.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTarjeta.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTarjeta.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTarjeta.grdGrilla
    End Select
    InicializarVisibilidadxCanal
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
              ' If frmTarjeta.RsCabecera.RecordCount < 8 Then
                  Sw = True
                  ActivarBotones (False)
                  Blanquear Me
                  chkActivo.Value = 1
                  chkPinPad.Value = 0
                  txtFactor.Text = "0.00"
                  'Cambia el Nombre del Primer Text
                  txtDetallado.SetFocus
                  Botonera
                  InicializarVisibilidadxCanal
'               Else
'                  MsgBox "Ha llegado al límite de 8 Tarjetas Bancarias", vbExclamation, sMensaje
'               End If
          
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim nCorrelaX As String
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                If Calcular("select count(*) as Codigo from ttarjetacredito where tcodigoIntegracion=" & IIf(Me.cboTipoIntegracion.BoundText = "", 0, Me.cboTipoIntegracion.BoundText) & " and tcodigoIntegracion<>0 and  tCodigotarjeta<>'" & txtCodigo.Text & "' ", Cn) > 0 Then
                    MsgBox "El enlace de integraciòn, ya esta siendo usado en otra tarjeta, favor de verificar!!!"
                    'Me.cboTipoIntegracion.BoundText = ""
                    Exit Sub
                End If
                
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoTarjeta) as Codigo from TTARJETACREDITO", Cn)
                  nCorrelaX = Calcular("select COUNT(*) as Codigo from TTARJETACREDITO", Cn)
                 
                 If nCorrelaX >= 24 Then
                    MsgBox "Maximo de tarjetas ya registrado, Cant Max 24!!!"
                    Exit Sub
                 End If
                  
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                                   

                sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETACREDITO", "TARJETA CREDITO", "01", sUsuario, sPasa, "", _
                   "TCODIGOTARJETA", "CODIGO TARJETA", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "tRepresentante", "Representante", txtRepresentante.Text, "nFactorRetencion", "Retencion", val(txtFactor.Text), _
                   "tCuentaContable", "Cuenta Contable", txtCuentaContable.Text, "tTelefono1", "Cuenta Contable Comision", val(txtTelefono1.Text), _
                   "lPinPad", "Flag Pin Pad", IIf(chkPinPad.Value, "Verdadero", "Falso"), "nBoton", "Botonera", val(txtBoton), "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), "tCodigoIntegracion", "Codigo Externo de Integracion ingresado en el sistema", Me.cboTipoIntegracion.BoundText)
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                         
                  'Cambiar el SQL
                  Isql = "insert into TTARJETACREDITO ( " & _
                         "tCodigoTarjeta, tDetallado, tResumido, nFactorRetencion, tRepresentante, " & _
                         "tTelefono1, nBoton, tCuentaContable, lPinPad, lActivo, tUsuario,tCodTarjetaEx,fRegistro, tCodigoIntegracion) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       val(txtFactor.Text) & ", " & _
                                " '" & txtRepresentante.Text & "', " & _
                                " '" & txtTelefono1.Text & "', " & _
                                   val(txtBoton.Caption) & ", " & _
                                " '" & txtCuentaContable.Text & "', " & _
                                       chkPinPad.Value & ", " & _
                                       chkActivo.Value & ", " & _
                                "'" & sUsuario & "'," & _
                                " '" & txtCodTarEx.Text & "', " & _
                                " getdate(), " & IIf(Me.cboTipoIntegracion.BoundText = "", 0, Me.cboTipoIntegracion.BoundText) & " )"
                                  
                  Cn.Execute Isql
                  

                  
                  
                  frmTarjeta.RsCabecera.Sort = "tCodigoTarjeta ASC"
                  frmTarjeta.RsCabecera.Requery
                  RsBoton.Requery
                  frmTarjeta.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & frmTarjeta.RsCabecera.AbsolutePosition & " de " & frmTarjeta.RsCabecera.RecordCount
               Else
               
                    sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETACREDITO", "TARJETA CREDITO", "02", sUsuario, sPasa, "", _
                   "TCODIGOTARJETA", "CODIGO TARJETA", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text, "tResumido", "Descripcion Resumida", txtResumido.Text, _
                   "tRepresentante", "Representante", txtRepresentante.Text, "nFactorRetencion", "Retencion", val(txtFactor.Text), _
                   "tCuentaContable", "Cuenta Contable", txtCuentaContable.Text, "tTelefono1", "Cuenta Contable Comision", val(txtTelefono1.Text), _
                   "lPinPad", "Flag Pin Pad", IIf(chkPinPad.Value, "Verdadero", "Falso"), "nBoton", "Botonera", val(txtBoton), "lActivo", "Flag Activo", IIf(chkActivo.Value, "Verdadero", "Falso"), "tCodigoIntegracion", "Codigo Externo de Integracion ingresado en el sistema", Me.cboTipoIntegracion.BoundText)
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
                  'Cambiar el SQL
                  Isql = "update TTARJETACREDITO set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "nFactorRetencion =" & val(txtFactor.Text) & ", " & _
                         "tRepresentante ='" & txtRepresentante.Text & "', " & _
                         "tTelefono1 ='" & txtTelefono1.Text & "', " & _
                         "tCuentaContable ='" & txtCuentaContable.Text & "', " & _
                         "nBoton =" & val(txtBoton.Caption) & ", " & _
                         "lPinPad =" & chkPinPad.Value & ", " & _
                         "tCodTarjetaEx ='" & txtCodTarEx.Text & "', " & _
                         "lActivo =" & chkActivo.Value & ", tcodigoIntegracion= " & IIf(Me.cboTipoIntegracion.BoundText = "", 0, Me.cboTipoIntegracion.BoundText) & ", lReplica=1 " & _
                         " where tCodigoTarjeta = '" & txtCodigo & "'"
                       
                   Cn.Execute Isql
                   nPos = frmTarjeta.RsCabecera.Bookmark
                   frmTarjeta.RsCabecera.Requery
                   RsBoton.Requery
                   If frmTarjeta.RsCabecera.RecordCount = 0 Then
                      frmTarjeta.RsCabecera.Filter = adFilterNone
                   End If
                   frmTarjeta.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               Cn.Execute "update TTARJETACREDITO set lNuevoModificado=1 where tCodigoTarjeta = '" & txtCodigo & "'"
               cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
               
               AsignarVisibilidadxCanal
          
          Case Is = 2 ' Eliminar
               If frmTarjeta.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar la Tarjeta " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
                    sPasa = txtCodigo.Text
                  
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETACREDITO", "TARJETA CREDITO", "03", sUsuario, sPasa, "", _
                   "TCODIGOTARJETA", "CODIGO TARJETA", sPasa, "tDetallado", "Descripcion Detallada", txtDetallado.Text)
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
               MsgBox "El proceso no elimina el registro solo lo inactiva!!!", vbInformation
               
               'Cambia el Delete
               'Cn.Execute "delete from TTARJETACREDITO where tCodigoTarjeta = '" & txtCodigo & "'"
               Cn.Execute "Update TTARJETACREDITO set lactivo=0 where tCodigoTarjeta = '" & txtCodigo & "'"
               frmTarjeta.RsCabecera.Requery
               If frmTarjeta.RsCabecera.RecordCount <> 0 Then
                  frmTarjeta.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
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
    'Ingrese el SubTitulo
    Me.Caption = " Tarjetas Bancarias "
    fraDetalle.Caption = Me.Caption
    LlenaCombos
    'Botones
    Isql = "select tCodigoTarjeta, nBoton from TTARJETACREDITO order by tCodigoTarjeta"
    Set RsBoton = Lib.OpenRecordset(Isql, Cn)

    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
       chkPinPad.Value = 0
       txtFactor.Text = "0.00"
       Botonera
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & IIf(frmTarjeta.RsCabecera.RecordCount = 0, 0, frmTarjeta.RsCabecera.AbsolutePosition) & " de " & frmTarjeta.RsCabecera.RecordCount
    InicializarVisibilidadxCanal
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmTarjetaDetalle = Nothing
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

Private Sub Botonera()
    Dim i As Integer
    txtBoton.Caption = "0"
    If RsBoton.RecordCount <> 0 Then
        For i = 1 To 8
            RsBoton.MoveFirst
            RsBoton.Find ("nBoton=" & i)
            If RsBoton.EOF Then
               cmdBoton(i).backColor = vbButtonFace
               cmdBoton(i).Enabled = True
            Else
               cmdBoton(i).Enabled = False
               If RsBoton!tCodigoTarjeta = txtCodigo.Text Then
                  txtBoton.Caption = str(i)
                  cmdBoton(i).backColor = vbRed
               Else
                  cmdBoton(i).backColor = vbBlue
               End If
            End If
        Next i
    Else
       For i = 1 To 8
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

