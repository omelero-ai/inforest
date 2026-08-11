VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepAnulado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Transacciones"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmRepAnulado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar Grillado"
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
      Left            =   3690
      Picture         =   "frmRepAnulado.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4860
      Width           =   1455
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
      Height          =   4815
      Left            =   45
      TabIndex        =   25
      Top             =   0
      Width           =   7320
      Begin VB.CheckBox chkUsuarioPedido 
         Caption         =   "Todos los Meseros"
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
         Left            =   5175
         TabIndex        =   12
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkFranjaHoraria 
         Caption         =   "Franja Horaria"
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
         Left            =   5175
         TabIndex        =   17
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Transacción "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   1890
         TabIndex        =   33
         Top             =   3360
         Width           =   4065
         Begin VB.CheckBox Check3 
            Caption         =   "Transferidos"
            Height          =   195
            Left            =   225
            TabIndex        =   20
            Top             =   945
            Value           =   1  'Checked
            Width           =   2445
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Anulados"
            Height          =   195
            Left            =   225
            TabIndex        =   19
            Top             =   630
            Value           =   1  'Checked
            Width           =   2445
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Facturados"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   315
            Value           =   1  'Checked
            Width           =   2445
         End
      End
      Begin VB.CheckBox chkMotivo 
         Caption         =   "Todos los Motivos"
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
         Left            =   5175
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   615
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.ComboBox cboEstado 
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
         ItemData        =   "frmRepAnulado.frx":082E
         Left            =   1905
         List            =   "frmRepAnulado.frx":0838
         TabIndex        =   6
         Text            =   "cboEstado"
         Top             =   1380
         Width           =   3135
      End
      Begin VB.CheckBox chkEstado 
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
         Left            =   5160
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1395
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
         Left            =   4275
         Picture         =   "frmRepAnulado.frx":0853
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1740
         Width           =   765
      End
      Begin VB.TextBox txtTurno 
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1770
         Width           =   2265
      End
      Begin VB.CheckBox chkUsuario 
         Caption         =   "Todos los Usuarios"
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
         Left            =   5175
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1005
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkSalon 
         Caption         =   "Todos los Salones"
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
         Left            =   5175
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkTurno 
         Caption         =   "Todos los Turnos"
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
         Left            =   5175
         TabIndex        =   10
         Top             =   1785
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1905
         TabIndex        =   15
         Top             =   2940
         Width           =   1725
         _ExtentX        =   3043
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
         Format          =   56033281
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1905
         TabIndex        =   13
         Top             =   2520
         Width           =   1725
         _ExtentX        =   3043
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
         Format          =   56033281
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3675
         TabIndex        =   14
         Top             =   2505
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   56033283
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3675
         TabIndex        =   16
         Top             =   2925
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   56033283
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   1905
         TabIndex        =   0
         Top             =   210
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSDataListLib.DataCombo cboUsuario 
         Height          =   315
         Left            =   1905
         TabIndex        =   4
         Top             =   990
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSDataListLib.DataCombo cbomotivo 
         Height          =   315
         Left            =   1905
         TabIndex        =   2
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSDataListLib.DataCombo cboUsuarioPedido 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   2160
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Mesero :"
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
         Left            =   1065
         TabIndex        =   34
         Top             =   2175
         Width           =   750
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo Anulación :"
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
         Left            =   210
         TabIndex        =   32
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Impresión :"
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
         Left            =   225
         TabIndex        =   31
         Top             =   1440
         Width           =   1590
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Anulación :"
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
         Left            =   135
         TabIndex        =   30
         Top             =   1050
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salon :"
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
         Index           =   8
         Left            =   1200
         TabIndex        =   29
         Top             =   270
         Width           =   615
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
         Left            =   690
         TabIndex        =   28
         Top             =   3015
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
         Left            =   585
         TabIndex        =   27
         Top             =   2595
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno :"
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
         Left            =   1185
         TabIndex        =   26
         Top             =   1830
         Width           =   630
      End
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
      Left            =   5145
      Picture         =   "frmRepAnulado.frx":0955
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4860
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
      Left            =   2235
      Picture         =   "frmRepAnulado.frx":0A47
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4860
      Width           =   1455
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
      Left            =   780
      Picture         =   "frmRepAnulado.frx":0F79
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4860
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   435
      Top             =   5010
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepAnulado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSalon As Recordset
Dim RsUsuario As Recordset
Dim RsMEliminacion As Recordset
Dim Rsturno As Recordset
Dim rsReporte As ADODB.Recordset
Dim Reporte As New dsrMovimiento

Dim rsUsuarioPedido As Recordset

Dim sCriterio As String
Dim sTitulo As String
Dim sTurno As String
Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboSalon
         Isql = "Select * from vSalon"
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With

    With cboUsuario
         Isql = "Select * from TUSUARIO where lActivo = 1 and tgrupousuario <>'00' order by tCodigoUsuario"
         Set RsUsuario = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUsuario
         .DataField = "tResumido"
         .ListField = "tResumido"
         .BoundColumn = "tCodigoUsuario"
    End With
    
    With cbomotivo
         Isql = "select * from Vmotivoeliminacion where lActivo=1"
         Set RsMEliminacion = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMEliminacion
         .DataField = "tResumido"
         .ListField = "tResumido"
         .BoundColumn = "Codigo"
    End With
    
    With cboUsuarioPedido
        Isql = "select * from vmozo where SUBSTRING(codigo,1,1)<>'*' and lActivo=1"
         Set rsUsuarioPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsUsuarioPedido
         .DataField = "tResumido"
         .ListField = "tResumido"
         .BoundColumn = "Codigo"
    End With
End Sub

Private Sub chkEstado_Click()
   If chkEstado.Value = 1 Then
      cboEstado.Enabled = False
      cboEstado.Text = ""
   Else
      cboEstado.Enabled = True
   End If
End Sub

Private Sub chkMotivo_Click()
   If chkMotivo.Value = 1 Then
      cbomotivo.Enabled = False
      cbomotivo.Text = ""
   Else
      cbomotivo.Enabled = True
   End If
End Sub

Private Sub chkUsuarioPedido_Click()
   If chkUsuarioPedido.Value = 1 Then
      cboUsuarioPedido.Enabled = False
      cboUsuarioPedido.Text = ""
   Else
      cboUsuarioPedido.Enabled = True
   End If
End Sub

Private Sub cmdBusca_Click()
   Isql = "Select tTurno as Codigo, tCaja, convert(nVarChar,fInicial, 120) as Descripcion, convert(nvarchar,fFinal,120) as fFinal from MTURNO order by tTurno DESC"
   frmBusca.nPredeterm = 2
   Call ConfGrilla(4, frmBusca.grdGrilla, "Turno", 2, "Codigo", 1400, 0, 0, "", _
                                          "Caja", 2, "tCaja", 1100, 0, 0, "", _
                                          "Fecha Inicial", 2, "Descripcion", 2400, 0, 0, "", _
                                          "Fecha Final", 2, "fFinal", 2400, 0, 0, "")
      
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   sTurno = sCodigo
   txtTurno.Text = sCodigo
End Sub

Private Sub chkSalon_Click()
   If chkSalon.Value = 1 Then
      cboSalon.Enabled = False
      cboSalon.Text = ""
   Else
      cboSalon.Enabled = True
   End If
End Sub

Private Sub chkUsuario_Click()
   If chkUsuario.Value = 1 Then
      cboUsuario.Enabled = False
      cboUsuario.Text = ""
   Else
      cboUsuario.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   sTitulo = ""
   
   If ChkTurno.Value = 0 Then
      If sTurno = "" Then
         MsgBox "Debe escoger un Turno", vbCritical, sMensaje
         Exit Sub
      End If
            
      Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
      If Rsturno.RecordCount > 0 Then
         sTitulo = " Del " & Format(Rsturno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(Rsturno!fFinal), Format(Rsturno!finicial, "dd/mmm/yyyy") & " 23:59", Rsturno!fFinal), "dd/mmm/yyyy HH:mm")
      Else
         MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
      End If
   Else
      If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      If ChkFranjaHoraria.Value = 0 Then
            sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
    Else
            sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " Franja Horaria " & Format(dtpHorIni.Value, "HH:mm") & " - " & Format(dtpHorFin.Value, "HH:mm")
    End If
   End If
   fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
   fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
   
   If chkSalon.Value = 0 Then
      If cboSalon.Text = "" Then
         MsgBox "Debe escoger un Salon", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "mpedido.tSalon ='" & cboSalon.BoundText & "'"
   End If
   
   If chkUsuario.Value = 0 Then
      If cboUsuario.Text = "" Then
         MsgBox "Debe escoger un Usuario", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "(mpedido.tUsuarioAnulado ='" & cboUsuario.Text & "' or #DBTRANS.tUsuarioAnulado='" & cboUsuario.Text & "')"
      sTitulo = sTitulo & " filtrado por el Usuario : " & cboUsuario.Text
   End If
      
   If chkMotivo.Value = 0 Then
      If cbomotivo.Text = "" Then
         MsgBox "Debe escoger un Motivo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "(mpedido.tMotivoAnulacion='" & cbomotivo.BoundText & "' or #DBTRANS.tMotivoAnulacion='" & cbomotivo.BoundText & "')"
   End If
               
   If chkEstado.Value = 0 Then
      If cboEstado.Text = "" Then
         MsgBox "Debe escoger un Estado de impresión", vbCritical, sMensaje
         Exit Sub
      End If
      If cboEstado.Text = "Impresos" Then
         sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "lImprime =1"
      Else
         sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "lImprime =0"
      End If
   End If
            
    If chkUsuarioPedido.Value = 0 Then
      If cboUsuarioPedido.Text = "" Then
         MsgBox "Debe escoger un Mesero", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "mpedido.tUsuario ='" & cboUsuarioPedido.Text & "'"
      sTitulo = sTitulo & " filtrado por el Usuario : " & cboUsuario.Text
   End If
               
            
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If

               frmEmite.CRViewer.DisplayGroupTree = True
               Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               Reporte.PaperOrientation = crLandscape
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
               Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               Reporte.PaperOrientation = crLandscape
               Reporte.PrintOut
          
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
                Genera 1
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               ExportaExcel rsReporte
'               Genera
'               Screen.MousePointer = vbDefault
'               If rsReporte.EOF = True Then
'                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
'                  Exit Sub
'               End If
'               Reporte.ExportOptions.FormatType = 21
'               Reporte.ExportOptions.DestinationType = 1
'               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
'               cmdSave.ShowSave
'               If cmdSave.FileName = "" Then
'                  Exit Sub
'               End If
'               Reporte.ExportOptions.DiskFileName = cmdSave.FileName
'               Reporte.Export False
   End Select
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub ChkTurno_Click()
   If ChkTurno.Value = 1 Then
      cmdBusca.Enabled = False
      sTurno = ""
      txtTurno = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
      ChkFranjaHoraria.Enabled = True
      ChkFranjaHoraria.Value = 0
   Else
      cmdBusca.Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
      ChkFranjaHoraria.Enabled = False
      ChkFranjaHoraria.Value = 0
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   LlenaCombos
   sTurno = ""
   dtpFecIni.Value = Date
   dtpFecFin.Value = Date
   cmdBusca.Enabled = False
   sTurno = ""
   cboUsuario.Enabled = False
   cboUsuario.Text = ""
   cboSalon.Enabled = False
   cboSalon.Text = ""
   cboEstado.Enabled = False
   cboEstado.Text = ""
   cbomotivo.Enabled = False
   cbomotivo.Text = ""
   cboUsuarioPedido.Enabled = False
   cboUsuarioPedido.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsSalon = Nothing
   Set RsUsuario = Nothing
   Set rsUsuarioPedido = Nothing
   Set Rsturno = Nothing
   Set rsReporte = Nothing
   Set frmRepAnulado = Nothing
End Sub

Public Sub Genera(Optional ExportaExcel As Integer = 0)
   Dim oComando As clsComando
   Screen.MousePointer = vbHourglass
   
   Set oComando = New clsComando
   If Not oComando.CreateCmdSp("spRep_Anulacion", Cn) Then
      Set oComando = Nothing
      Exit Sub
   End If
   
   oComando.CreateParameter "@lFranjaHoraria", adBoolean, adParamInput, 1, IIf(ChkFranjaHoraria.Value = 1, True, False)
   oComando.CreateParameter "@sTurno", adVarChar, adParamInput, 10, sTurno
   oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
   oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
   oComando.CreateParameter "@lFlag1", adBoolean, adParamInput, 1, Check1.Value
   oComando.CreateParameter "@lFlag2", adBoolean, adParamInput, 1, Check2.Value
   oComando.CreateParameter "@lFlag3", adBoolean, adParamInput, 1, Check3.Value
   oComando.CreateParameter "@sCriterio", adVarChar, adParamInput, 400, sCriterio
   oComando.CreateParameter "@ExportaExcel", adBoolean, adParamInput, 1, ExportaExcel
                 
                     
   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If

   Set rsReporte = oComando.GetSP()
   'rsReporte.Filter = sCriterio
   If ExportaExcel = 1 Then Exit Sub

   Reporte.DiscardSavedData
   Reporte.Database.SetDataSource rsReporte
   Reporte.Text14.SetText sRazonSocial
   Reporte.Text28.SetText localConectado
   Reporte.ReportTitle = IIf(ChkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
   frmEmite.CRViewer.ReportSource = Reporte
End Sub
