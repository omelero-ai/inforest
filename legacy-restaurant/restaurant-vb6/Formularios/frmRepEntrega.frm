VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Entregas"
   ClientHeight    =   5550
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   6720
      TabIndex        =   30
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Emite"
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
      Left            =   1620
      Picture         =   "frmRepEntrega.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4935
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Impresora"
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
      Left            =   3075
      Picture         =   "frmRepEntrega.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4935
      Width           =   1455
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
      Index           =   2
      Left            =   5985
      Picture         =   "frmRepEntrega.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4935
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar"
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
      Left            =   4530
      Picture         =   "frmRepEntrega.frx":0B56
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4935
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   4935
      Top             =   5055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
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
      Height          =   4920
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9015
      Begin VB.CheckBox chkGrupo 
         Caption         =   "Todos los Grupos"
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
         Left            =   6720
         TabIndex        =   33
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkSubGrupo 
         Caption         =   "Todos los SubGrupos"
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
         Left            =   6720
         TabIndex        =   32
         Top             =   2550
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "Formato 2"
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
         Left            =   6240
         TabIndex        =   31
         Top             =   4455
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "Resumido Por Producto               "
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
         Left            =   1845
         TabIndex        =   29
         Top             =   4125
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.CheckBox chkCanalVenta 
         Caption         =   "Todos los Canales"
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
         Left            =   6720
         TabIndex        =   9
         Top             =   1335
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkEntrega 
         Caption         =   "Todos los Estados"
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
         Left            =   6720
         TabIndex        =   17
         Top             =   3445
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Todos los Clientes"
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
         Left            =   6720
         TabIndex        =   12
         Top             =   1755
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CommandButton cmdBusca 
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
         Height          =   375
         Index           =   0
         Left            =   6030
         Picture         =   "frmRepEntrega.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2925
         Width           =   645
      End
      Begin VB.CommandButton cmdBusca 
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
         Height          =   375
         Index           =   1
         Left            =   6030
         Picture         =   "frmRepEntrega.frx":147A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   645
      End
      Begin VB.TextBox txtProducto 
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
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2955
         Width           =   4185
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1740
         Width           =   4185
      End
      Begin VB.CheckBox chkCancelacion 
         Caption         =   "Todos los Estados"
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
         Left            =   6720
         TabIndex        =   19
         Top             =   3850
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1860
         TabIndex        =   6
         Top             =   795
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51183617
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1860
         TabIndex        =   4
         Top             =   315
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51183617
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   5010
         TabIndex        =   5
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   51183619
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   5010
         TabIndex        =   7
         Top             =   795
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   51183619
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboEntrega 
         Height          =   315
         Left            =   1860
         TabIndex        =   16
         Top             =   3400
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cboCancelacion 
         Height          =   315
         Left            =   1860
         TabIndex        =   18
         Top             =   3830
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.CheckBox chkProducto 
         Caption         =   "Todos los Productos"
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
         Left            =   6720
         TabIndex        =   15
         Top             =   2970
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin MSDataListLib.DataCombo cboCanalVenta 
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Top             =   1320
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Detallado Por Pedido       Formato 1"
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
         Left            =   1845
         TabIndex        =   28
         Top             =   4455
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1860
         TabIndex        =   34
         Top             =   2140
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   1860
         TabIndex        =   35
         Top             =   2550
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo :"
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
         Height          =   195
         Index           =   4
         Left            =   1140
         TabIndex        =   37
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubGrupo :"
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
         Height          =   195
         Index           =   6
         Left            =   810
         TabIndex        =   36
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Canal de Venta :"
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
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   27
         Top             =   1440
         Width           =   1440
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto :"
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
         Height          =   195
         Index           =   7
         Left            =   885
         TabIndex        =   26
         Top             =   3015
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
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
         Height          =   195
         Left            =   660
         TabIndex        =   25
         Top             =   850
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
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
         Height          =   195
         Left            =   555
         TabIndex        =   24
         Top             =   390
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
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
         Height          =   195
         Index           =   2
         Left            =   1065
         TabIndex        =   23
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Entrega :"
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
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   22
         Top             =   3480
         Width           =   1440
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Cancelación :"
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
         Height          =   195
         Index           =   3
         Left            =   -30
         TabIndex        =   21
         Top             =   3850
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmRepEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReporteE As New dsrReporteEntrega
Dim ReporteEF As New dsrReporteEntregaFormato2
Dim ReporteRp As New dsrReporteEntregaResumidoProd
Dim sTitulo As String
Dim rsReporte As Recordset
Dim rsEntrega As New ADODB.Recordset
Dim rsCanalVenta As New ADODB.Recordset
Dim rsCancelacion As New ADODB.Recordset
Dim sCliente As String
Dim sProducto As String

Dim fInicio As Date
Dim fFinal As Date

Private Sub cboGrupo_Change()
cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubGrupo " & IIf(cboGrupo.BoundText = "", "", "where tGrupo = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
    End With
    
    sProducto = ""
    txtProducto.Text = ""
End Sub

Private Sub chkCanalVenta_Click()
  If chkCanalVenta.Value = 1 Then
          cboCanalVenta.Enabled = False
          cboCanalVenta.Text = ""
    Else
          cboCanalVenta.Enabled = True
    End If
End Sub

Private Sub chkCancelacion_Click()
    If chkCancelacion.Value = 1 Then
          cboCancelacion.Enabled = False
          cboCancelacion.Text = ""
    Else
          cboCancelacion.Enabled = True
    End If
End Sub

Private Sub chkCliente_Click()
        If chkCliente.Value = 1 Then
                sCliente = ""
                txtCliente.Text = ""
                cmdBusca(1).Enabled = False
        Else
                cmdBusca(1).Enabled = True
        End If
End Sub

Private Sub chkEntrega_Click()
    If chkEntrega.Value = 1 Then
          cboEntrega.Enabled = False
          cboEntrega.Text = ""
    Else
          cboEntrega.Enabled = True
    End If
End Sub

Private Sub chkGrupo_Click()
If chkGrupo.Value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
   Else
      cboGrupo.Enabled = True
   End If
End Sub

Private Sub chkproducto_Click()
        If Me.chkProducto.Value = 1 Then
                sProducto = ""
                txtProducto.Text = ""
                cmdBusca(0).Enabled = False
        Else
                cmdBusca(0).Enabled = True
        End If
End Sub

Private Sub chkSubGrupo_Click()
If chkSubGrupo.Value = 1 Then
      cboSubGrupo.Enabled = False
      cboSubGrupo.Text = ""
   Else
      cboSubGrupo.Enabled = True
   End If
End Sub

Private Sub cmdBusca_Click(Index As Integer)
    Dim xCriterio As String
    
   Select Case Index
          Case Is = 0
          
               xCriterio = "lActivo = 1 "
               If cboGrupo.Text <> "" Then
                  xCriterio = xCriterio & " and tGrupo ='" & cboGrupo.BoundText & "'"
               End If
               
               If cboSubGrupo.Text <> "" Then
                  xCriterio = xCriterio & " and tSubGrupo ='" & cboSubGrupo.BoundText & "'"
               End If
               
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where " & xCriterio & " order by Descripcion"
               frmBusca.nPredeterm = 3
               Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                                      "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                                      "Plato", 2, "Descripcion", 2900, 0, 0, "", _
                                                      "Precio", 2, "nPrecioVenta", 900, 1, 0, "##,##0.00")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sProducto = sCodigo
               txtProducto.Text = sDescrip
          
          Case Is = 1
               Isql = "Select tCodigoDelivery as Codigo, tNombre + ' ' + tApellido as Descripcion from TDELIVERY order by tCodigoDelivery DESC"
               frmBusca.nPredeterm = 1
               Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1400, 0, 0, "", _
                                                      "Cliente", 2, "Descripcion", 5900, 0, 0, "")
                  
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sCliente = sCodigo
               txtCliente.Text = sDescrip
   End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 2 Then
        Unload Me
    End If
        sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")

    If chkCanalVenta.Value = 0 Then
        If cboCanalVenta.Text = "" Then
                  MsgBox "Debe escoger un Canal de Venta", vbCritical, sMensaje
                  Exit Sub
        End If
        sTitulo = sTitulo & ". Canal de Venta: " & cboCanalVenta.Text
    Else
        sTitulo = sTitulo & ". Todos los Canales de Venta "
    End If
    
    If chkCliente.Value = 0 Then
        If txtCliente.Text = "" Then
                  MsgBox "Debe escoger un Cliente", vbCritical, sMensaje
                  Exit Sub
        End If
        
        sTitulo = sTitulo & ". Cliente : " & txtCliente.Text
    Else
        sTitulo = sTitulo & ". Todos los Clientes "
    End If
    
   If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger el grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sTitulo = sTitulo & ". Grupo='" & cboGrupo.Text & "' "
    Else
      sTitulo = sTitulo & ". Todos los Grupos "
   End If
   
   If chkSubGrupo.Value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger el subgrupo", vbCritical, sMensaje
         Exit Sub
      End If
      sTitulo = sTitulo & ". SubGrupo='" & cboSubGrupo.Text & "' "
   Else
      sTitulo = sTitulo & ". Todos los Subgrupos "
   End If
   
   If chkProducto.Value = 0 Then
        If txtProducto.Text = "" Then
                  MsgBox "Debe escoger un Producto", vbCritical, sMensaje
                  Exit Sub
        End If
        sTitulo = sTitulo & ". Producto : " & txtProducto.Text
   Else
        sTitulo = sTitulo & ". Todos los Productos "
   End If
    
    
   If chkEntrega.Value = 0 Then
        If cboEntrega.Text = "" Then
                  MsgBox "Debe escoger un Estado de Entrega", vbCritical, sMensaje
                  Exit Sub
        End If
        sTitulo = sTitulo & ". Estado Entrega: " & cboEntrega.Text
   Else
        sTitulo = sTitulo & ". Todos los Estados de Entrega"
    End If
    
   If chkCancelacion.Value = 0 Then
        If cboCancelacion.Text = "" Then
                  MsgBox "Debe escoger un Estado de Cancelacion", vbCritical, sMensaje
                  Exit Sub
        End If
        sTitulo = sTitulo & ". Estado de Cancelacion: " & cboCancelacion.Text
    Else
        sTitulo = sTitulo & ". Todos los Estados de Cancelacion"
    End If
 
    Select Case Index
          Case Is = 0 ' Preview
         
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
               If Option1.Value = True Then
                    ReporteE.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    ReporteE.PaperOrientation = crLandscape
                ElseIf Option3.Value = True Then
                     ReporteEF.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    ReporteEF.PaperOrientation = crLandscape
                Else
                    ReporteRp.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    ReporteRp.PaperOrientation = crLandscape
               End If
               
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
               If Option1.Value = True Then
                  ReporteE.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteE.PaperOrientation = crLandscape
                  ReporteE.PrintOut
               ElseIf Option3.Value = True Then
                  ReporteEF.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteEF.PaperOrientation = crLandscape
                  ReporteEF.PrintOut
               Else
                  ReporteRp.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteRp.PaperOrientation = crLandscape
                  ReporteRp.PrintOut
               End If
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If Option1.Value = True Then
                  ReporteE.ExportOptions.FormatType = 21
                  ReporteE.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteE.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteE.Export False
               ElseIf Option3.Value = True Then
                  ReporteEF.ExportOptions.FormatType = 21
                  ReporteEF.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteEF.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteEF.Export False
               Else
                  ReporteRp.ExportOptions.FormatType = 21
                  ReporteRp.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteRp.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteRp.Export False
               End If
   End Select
   
End Sub


Public Sub Genera()
    Screen.MousePointer = vbHourglass
    Dim oComando As clsComando
   
    Set oComando = New clsComando
   
    If Not oComando.CreateCmdSp("spRep_Entregas", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    fInicio = Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:mm")

    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 20, sCliente
    oComando.CreateParameter "@tEstadoEntrega", adVarChar, adParamInput, 20, Me.cboEntrega.BoundText
    oComando.CreateParameter "@tEstadoCanc", adVarChar, adParamInput, 20, Me.cboCancelacion.BoundText
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 10, Me.cboGrupo.BoundText
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 10, Me.cboSubGrupo.BoundText
    oComando.CreateParameter "@tProducto", adVarChar, adParamInput, 20, sProducto
    oComando.CreateParameter "@tCanalVenta", adVarChar, adParamInput, 20, cboCanalVenta.BoundText
    oComando.CreateParameter "@tTipo", adVarChar, adParamInput, 1, IIf(Option2.Value = True, "0", "1")

    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    If Option1.Value = True Then
      ReporteE.DiscardSavedData
      ReporteE.Database.SetDataSource rsReporte
      ReporteE.ReportTitle = sTitulo
      ReporteE.Text13.SetText sRazonSocial
      ReporteE.Text14.SetText localConectado
      frmEmite.CRViewer.ReportSource = ReporteE
    ElseIf Option3.Value = True Then
      ReporteEF.DiscardSavedData
      ReporteEF.Database.SetDataSource rsReporte
      ReporteEF.ReportTitle = sTitulo
      ReporteEF.Text13.SetText sRazonSocial
      ReporteEF.Text14.SetText localConectado
      frmEmite.CRViewer.ReportSource = ReporteEF
    Else
      ReporteRp.DiscardSavedData
      ReporteRp.Database.SetDataSource rsReporte
      ReporteRp.ReportTitle = sTitulo
      ReporteRp.Text13.SetText sRazonSocial
      ReporteRp.Text14.SetText localConectado
      frmEmite.CRViewer.ReportSource = ReporteRp
    End If
        
End Sub
Sub LlenaCombos()
    With cboEntrega
             Isql = "select 'ENTREGADO' as Codigo, 'ENTREGADO' as Descripcion  union select 'PENDIENTE' as Codigo ,'PENDIENTE' as Descripcion"
             Set rsEntrega = Lib.OpenRecordset(Isql, Cn)
             Set .RowSource = rsEntrega
                 .DataField = "Descripcion"
                 .ListField = "Descripcion"
                 .BoundColumn = "Codigo"
    End With
    With cboCancelacion
             Isql = "select 'ANTICIPO' as Codigo, 'ANTICIPO' as Descripcion UNION select 'NO PAGADO' as Codigo, 'NO PAGADO' as Descripcion union select 'PAGADO' as Codigo ,'PAGADO'"
             Set rsCancelacion = Lib.OpenRecordset(Isql, Cn)
             Set .RowSource = rsCancelacion
                 .DataField = "Descripcion"
                 .ListField = "Descripcion"
                 .BoundColumn = "Codigo"
    End With

    With cboCanalVenta
            Isql = "select tCodigoCanalVenta as Codigo,tDetallado as Descripcion From TCANALVENTA Where lActivo = 1 order by 2"
             Set rsCanalVenta = Lib.OpenRecordset(Isql, Cn)
             Set .RowSource = rsCanalVenta
                 .DataField = "Descripcion"
                 .ListField = "Descripcion"
                 .BoundColumn = "Codigo"
    End With
    
        With cboGrupo
         Isql = "Select * from vGrupo order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
           
    With cboSubGrupo
         Isql = "Select * from vSubGrupo order by Descripcion"
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub

Private Sub dtpFecfin_LostFocus()
 If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date + 1
    'cmdBusca.Enabled = False
    sCliente = ""
    sProducto = ""
    txtCliente.Text = ""
    txtProducto.Text = ""
End Sub


