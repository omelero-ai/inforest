VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepLiquidacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Cajero"
   ClientHeight    =   6105
   ClientLeft      =   4125
   ClientTop       =   1860
   ClientWidth     =   7950
   Icon            =   "frmRepLiquidacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   90
      Top             =   5370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   3810
      Picture         =   "frmRepLiquidacion.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
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
      Height          =   5340
      Left            =   30
      TabIndex        =   20
      Top             =   0
      Width           =   7785
      Begin VB.CheckBox chkSectorVenta 
         Caption         =   "Todas las Agrupaciones"
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
         Left            =   5235
         TabIndex        =   29
         Top             =   2445
         Value           =   1  'Checked
         Width           =   2500
      End
      Begin VB.CheckBox chkCortesia 
         Caption         =   "Ver Cortesias"
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkResumen 
         Caption         =   "Ver Resumen de Liquidación"
         Height          =   240
         Left            =   1320
         TabIndex        =   13
         Top             =   2880
         Width           =   3435
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cálculo Ticket Promedio"
         Height          =   1575
         Left            =   2805
         TabIndex        =   27
         Top             =   3600
         Width           =   2055
         Begin VB.OptionButton optNeto 
            Caption         =   "Valor Neto"
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optVenta 
            Caption         =   "Valor Venta"
            Height          =   195
            Left            =   195
            TabIndex        =   19
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Presentacion"
         Height          =   1575
         Left            =   645
         TabIndex        =   26
         Top             =   3600
         Width           =   2055
         Begin VB.CheckBox chkConGrafico 
            Caption         =   "Con Gráfico"
            Enabled         =   0   'False
            Height          =   195
            Left            =   600
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Formato 2"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Formato 1"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   285
         Width           =   1815
      End
      Begin VB.CommandButton cmdBusca 
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
         Left            =   3270
         Picture         =   "frmRepLiquidacion.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   765
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
         Left            =   5235
         TabIndex        =   7
         Top             =   660
         Value           =   1  'Checked
         Width           =   1995
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
         Left            =   5235
         TabIndex        =   5
         Top             =   255
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3270
         TabIndex        =   10
         Top             =   1410
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   56098819
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3270
         TabIndex        =   12
         Top             =   1875
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   56098819
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   1875
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm"
         Format          =   56098817
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   1410
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm"
         Format          =   56098817
         CurrentDate     =   37539.2083333333
      End
      Begin MSDataListLib.DataCombo cboUsuario 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   675
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.CheckBox chkDiaContable 
         Caption         =   "Evaluar por Dia Contable"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   940
         Width           =   2715
      End
      Begin MSDataListLib.DataCombo cboSectorVenta 
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.CheckBox chkGenerado 
         Caption         =   "Ver Resumen de Liquidación"
         Height          =   1680
         Left            =   5040
         TabIndex        =   28
         Top             =   3360
         Width           =   1995
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Punto Venta :"
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
         Left            =   135
         TabIndex        =   32
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agrupacion"
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
         Index           =   14
         Left            =   135
         TabIndex        =   31
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
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
         Left            =   540
         TabIndex        =   24
         Top             =   735
         Width           =   780
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
         Left            =   690
         TabIndex        =   23
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label 
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
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   1500
         Width           =   1230
      End
      Begin VB.Label Label 
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
         Index           =   1
         Left            =   195
         TabIndex        =   21
         Top             =   1965
         Width           =   1125
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
      Left            =   5280
      Picture         =   "frmRepLiquidacion.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
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
      Left            =   2355
      Picture         =   "frmRepLiquidacion.frx":0A22
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
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
      Left            =   900
      Picture         =   "frmRepLiquidacion.frx":0F54
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset
Dim RsSumas As Recordset
Dim RsTarjetas As Recordset
Dim RsUsuario As Recordset
Dim RsTc As Recordset
Dim Rsturno As Recordset

Dim sNombre   As String
Dim sSuma     As String
Dim sCriterio As String
Dim sBase     As String
Dim xBase     As String
Dim yBase     As String
Dim sTitulo   As String

Dim sFechaFin As String
Dim nFondoN As Double
Dim nFondoE As Double
Dim i As Long
Dim nCambio As Double
Dim xFecha As Date
Dim nTotalPromedio As Double

Dim nDocumento As Double
Dim nEfectivoN As Double
Dim nEfectivoE As Double
Dim xEfectivoE As Double
Dim nIngresoN As Double
Dim nIngresoE As Double
Dim xIngresoE As Double
Dim nIngresoAN As Double
Dim nIngresoAE As Double
Dim xIngresoAE As Double
Dim nEgresoN As Double
Dim nEgresoE As Double
Dim xEgresoE As Double
'Nota de credito
Dim nNotaCreditoN As Double
Dim nNotaCreditoE As Double
Dim xNotaCreditoE As Double

Dim tOtro(20) As String
Dim nOtro(20) As Double
Dim xOtro(20) As Double
Dim nDiferencia(20) As Double

Dim tTarjeta(24) As String
Dim nTarjeta(24) As Double
Dim nPropina(24) As Double
Dim nPRetencion(24) As Double
Dim nRetencion(24) As Double
Dim nRecibo(24) As Double
Dim nTarjetaT As Double
Dim nRetencionT As Double
Dim nPropinaT As Double
Dim nReciboT As Double

Dim nChequeN As Double
Dim nChequeE As Double
Dim xChequeE As Double
Dim nOtroN As Double
Dim nOtroE As Double
Dim xOtroE As Double
Dim nPunto As Double
Dim nCortesia As Double
      
Dim nTotalCobrado As Double
Dim nTotalCobradoN As Double
Dim nTotalCobradoE As Double
Dim xTotalCobradoE As Double

Dim nTotalEfectivo As Double
Dim nTotalEfectivoN As Double
Dim nTotalEfectivoE As Double
Dim xTotalEfectivoE As Double
Dim nFisEfectivo As Double
Dim nFisEfectivoN As Double

Dim nNeto As Double
Dim nImpuesto1 As Double
Dim nImpuesto2 As Double
Dim nImpuesto3 As Double
Dim nTotal As Double
Dim nDescuento As Double
Dim nRecargo As Double

Dim nTotalCierre As Double
Dim nCobrar As Double
Dim nNoCobrado As Double
Dim nAdulto As Double
Dim nNino As Double
Dim nDolar As Double
Dim xDolar As Double
Dim nOtroDoc As Double
Dim nOtroPed As Double
Dim nFacCPedido As Double

Dim fInicio As Date
Dim fFinal As Date

Dim Reporte As New dsrLiquidacion
Dim ReporteV As New dsrLiquidacionVenta

Dim tFiscal As Double

Dim RsSectorVenta As Recordset

'CESAR
Dim nAdulto2 As Double
Dim nNino2 As Double
Dim nAdulto3 As Double
Dim nNino3 As Double
Dim nAdulto4 As Double
Dim nNino4 As Double
Dim nAdulto5 As Double
Dim nNino5 As Double


Dim nPRetencionPro(24) As Double
Dim nRetencionProT As Double



Sub LlenaCombos()
    With cboUsuario
         Isql = "Select * from TUSUARIO where lActivo = 1 and tGrupoUsuario<>'00' order by tCodigoUsuario"
         Set RsUsuario = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsUsuario
         .DataField = "tResumido"
         .ListField = "tResumido"
         .BoundColumn = "tCodigoUsuario"
    End With
    
    With cboSectorVenta
         Isql = "Select * from vSectorVenta where Activo = 1 order by Codigo"
         Set RsSectorVenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSectorVenta
         .DataField = "Resumido"
         .ListField = "Resumido"
         .BoundColumn = "Codigo"
    End With
    
    
End Sub


Private Sub chkSectorVenta_Click()
   If chkSectorVenta.Value = 1 Then
      cboSectorVenta.Enabled = False
      cboSectorVenta.Text = ""
   Else
      cboSectorVenta.Enabled = True
   End If
End Sub

Private Sub cmdBusca_Click(Index As Integer)
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

Private Sub ChkTurno_Click()
   If ChkTurno.Value = 1 Then
      cmdBusca(0).Enabled = False
      txtTurno.Text = ""
      sTurno = ""
      txtTurno.Text = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
      chkDiaContable.Enabled = True
      chkDiaContable.Value = 0
      
      chkSectorVenta.Enabled = True
      cboSectorVenta.Enabled = False
      
   Else
      cmdBusca(0).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
      chkDiaContable.Enabled = False
      chkDiaContable.Value = 0
      
      chkSectorVenta.Value = 1
      chkSectorVenta.Enabled = False
      cboSectorVenta.Enabled = False
      
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
    tFiscal = 0
   'Fecha del Turno
   If ChkTurno.Value = 0 Then
      xFecha = Calcular("select fInicial as codigo from MTURNO where tTurno='" & sTurno & "'", Cn)
      xFecha = Format(xFecha, "yyyy/mm/dd HH:mm:ss")
   Else
      xFecha = Format(dtpFecIni.Value, "yyyy/mm/dd")
   End If
   
   sCriterio = "isnull(tCaja,'0')<>'0'"
   sBase = "isnull(MDOCUMENTO.tCaja,'0')<>'0'"
   xBase = "isnull(DPAGODOCUMENTO.tDocumento,'0')<>'0'"
   yBase = "isnull(MPEDIDO.tCaja,'0')<>'0'"
      
   If ChkTurno.Value = 0 Then
      If sTurno = "" Then
         MsgBox "No se ha asignado el Turno", vbExclamation, sMensaje
         Exit Sub
      End If
      
      Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
      If Not Rsturno.EOF Then
         nFondoN = IIf(IsNull(Rsturno!nMontoIN), 0, Rsturno!nMontoIN)
         nFondoE = IIf(IsNull(Rsturno!nMontoIE), 0, Rsturno!nMontoIE)
         sFechaFin = IIf(IsNull(Rsturno!fFinal), "", Rsturno!fFinal)
      Else
         nFondoN = 0
         nFondoE = 0
         sFechaFin = ""
      End If
      
      sTitulo = "Turno : " & sTurno & " Del " & Format(xFecha, "dd/mm/yyyy HH:nn") & " Hrs Al " & Format(sFechaFin, "dd/mm/yyyy HH:nn") & " Hrs "
      
      sCriterio = sCriterio & " And (tTurno ='" & sTurno & "') "
      sBase = sBase & " And MDOCUMENTO.tTurno ='" & sTurno & "'"
      xBase = xBase & " And DPAGODOCUMENTO.tTurno ='" & sTurno & "'"
      yBase = yBase & " And MPEDIDO.tTurno ='" & sTurno & "'"
'        If Me.ChkExcluirNC Then
'            sBase = sBase & " And MDOCUMENTO.tdocumento  in ( select tdocumento from mnotacredito where tturno= '" & sTurno & "' and testadodocumento in ('02','05') ) "
'        End If
   Else
      If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
          MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
          Exit Sub
          dtpFecFin.SetFocus
      End If
      nFondoN = 0
      nFondoE = 0
      sFechaFin = dtpFecFin.Value
       If chkDiaContable.Value = 0 Then
       
        sTitulo = "Turno : Todos los Turnos Del " & Format(xFecha, "dd/mm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Hrs Al " & Format(sFechaFin, "dd/mm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm") & " Hrs "
      Else
        sTitulo = ". Por Dia Contable, Todos los Turnos del  " & Format(xFecha, "dd/mm/yyyy") & "   Al " & Format(sFechaFin, "dd/mm/yyyy") & "   "
      End If
      
      sCriterio = sCriterio & " And fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "hh:mm") & "' And fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "'"
      sBase = sBase & " And ((MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' And MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "')) "
      xBase = xBase & " And ((DPAGODOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' And DPAGODOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "')) "
      yBase = yBase & " And MPEDIDO.fRegistro>='" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MPEDIDO.fRegistro<='" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "'"
'        If Me.ChkExcluirNC Then
'            sBase = sBase & " And MDOCUMENTO.tdocumento not in ( select tdocumento from mnotacredito where fregistro >= '" & Format(dtpFecIni.value, "yyyy/MM/dd") & "' " & Format(dtpHorIni.value, "HH:nn") & "' And fRegistro <= '" & Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:nn") & " and testadodocumento in ('02','05') ) "
'        End If
      
   End If
   
   If chkUsuario.Value = 0 Then
      If cboUsuario.Text = "" Then
         MsgBox "No se ha asignado el Usuario", vbExclamation, sMensaje
         Exit Sub
      End If
      
      sTitulo = sTitulo & Chr(13) & "Usuario : " & cboUsuario.Text
      
      sCriterio = sCriterio & " And tUsuario ='" & cboUsuario.Text & "'"
      sBase = sBase & " And MDOCUMENTO.tUsuario ='" & cboUsuario.Text & "'"
      xBase = xBase & " And DPAGODOCUMENTO.tUsuario ='" & cboUsuario.Text & "'"
      yBase = yBase & " And MPEDIDO.tUsuario ='" & cboUsuario.Text & "'"
   Else
      sTitulo = sTitulo & Chr(13) & "Usuario : Todos los Usuarios"
   End If
   
   
   
   
   'diacontable
  If chkDiaContable.Value = 0 Then
            fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
            fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
  Else
            fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd")
            fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd")
  End If
  'diacontable
  
   If chkSectorVenta.Value = 0 Then
      If cboSectorVenta.Text = "" Then
         MsgBox "No se ha asignado el Sector de Venta", vbExclamation, sMensaje
         Exit Sub
      End If
      sTitulo = sTitulo & Chr(13) & "Sectores de Venta : " & cboSectorVenta.Text
      
   Else
      If ChkTurno.Value = 1 Then
         sTitulo = sTitulo & Chr(13) & "Sectores de Venta : Todos los Sectores de Ventas"
      End If
   End If
  
   Select Case Index
          Case Is = 0 ' Preview
                If Option1.Value Then
                    Genera
                    If rsReporte.EOF = True Then
                       MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                       Screen.MousePointer = vbDefault
                       Exit Sub
                    End If
                    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Reporte.PaperOrientation = crPortrait
                    frmEmite.CRViewer.ReportSource = Reporte
                    frmEmite.CRViewer.DisplayGroupTree = True
                    frmEmite.CRViewer.ViewReport
                    frmEmite.Show vbModal
                Else
                    If Option2.Value Then
                        Genera2
                        frmEmite.CRViewer.ReportSource = ReporteV
                        frmEmite.CRViewer.DisplayGroupTree = True
                        frmEmite.CRViewer.ViewReport
                        frmEmite.Show vbModal
                    End If
                End If
          Case Is = 1 ' Imprimir
                If Option1.Value Then
                   Genera
                   Screen.MousePointer = vbDefault
                   If rsReporte.EOF = True Then
                      MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                      Exit Sub
                   End If
                   Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                   Reporte.PaperOrientation = crPortrait
                   Screen.MousePointer = vbDefault
                   Reporte.PrintOut
                Else
                    If Option2.Value Then
                        Genera2
                        Screen.MousePointer = vbDefault
                        If rsReporte.EOF = True Then
                            MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                                           
                        ReporteV.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        ReporteV.PaperOrientation = crPortrait
                        Screen.MousePointer = vbDefault
                        ReporteV.PrintOut
                    End If
                End If
          Case Is = 3 ' Exportar
               On Error GoTo Exportar
               If Option1.Value Then
                    Genera
                    Screen.MousePointer = vbDefault
                    If rsReporte.EOF = True Then
                       MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                       Exit Sub
                    End If
                    Reporte.ExportOptions.FormatType = 21
                    Reporte.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    Screen.MousePointer = vbDefault
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                       Exit Sub
                    End If
                    Reporte.ExportOptions.DiskFileName = cmdSave.FileName
                    Reporte.Export False
               Else
                    If Option2.Value Then
                        Genera2
                        Screen.MousePointer = vbDefault
                        If rsReporte.EOF = True Then
                           MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                           Exit Sub
                        End If
                        ReporteV.ExportOptions.FormatType = 21
                        ReporteV.ExportOptions.DestinationType = 1
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        Screen.MousePointer = vbDefault
                        cmdSave.ShowSave
                        If cmdSave.FileName = "" Then
                           Exit Sub
                        End If
                        ReporteV.ExportOptions.DiskFileName = cmdSave.FileName
                        ReporteV.Export False
                     End If
               End If
                              
   End Select
Exit Sub
Exportar:
    MsgBox err.Number & " " & err.Description & " " & err.Source, vbCritical
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub



Private Sub Form_Load()
 '  sNombre = dbTemporal(sCaja, 18, "tGrupo", "nVarChar(2)", _
                                   "Grupo", "nVarChar(50)", _
                                   "tSubGrupo", "nVarChar(3)", _
                                   "SubGrupo", "nVarChar(50)", _
                                   "tDocumento", "nVarChar(15)", _
                                   "tUsuario", "nVarChar(15)", _
                                   "fFecha", "nVarChar(35)", _
                                   "fPago", "nVarChar(35)", _
                                   "nTC", "Float", _
                                   "nVenta", "Float", _
                                   "nVenta1", "Float", _
                                   "nVenta2", "Float", _
                                   "tObservacion", "nVarChar(250)", _
                                   "tTipoDocumento", "nVarChar(1500)", _
                                   "tInicio", "nVarChar(1500)", _
                                   "tFinal", "nVarChar(1500)", _
                                   "tEmitido", "nVarChar(1500)", _
                                   "tAnulado", "nVarChar(1500)")

   sTurno = ""
   Centrar Me
   LlenaCombos
   cboUsuario.Enabled = False
   dtpFecIni.Value = Date
   dtpFecFin.Value = Date
   dtpHorIni.Value = "12:00 am"
   dtpHorFin.Value = Time
   
 '  sSuma = dbTemporal(sCaja, 7, "tDocumento", "nVarChar(15)", _
                                 "nNeto", "float", _
                                 "nImpuesto1", "float", _
                                 "nImpuesto2", "float", _
                                 "nImpuesto3", "float", _
                                 "nVenta", "float", _
                                 "tTipoPedido", "nVarChar(2)")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsSumas = Nothing
   Set RsTc = Nothing
   Set rsReporte = Nothing
   Set RsUsuario = Nothing
   Set Rsturno = Nothing
   Set frmRepLiquidacion = Nothing
End Sub

Public Sub Genera()
   Dim xparametros As String, tipo As String
   Dim oComando As clsComando: Dim pagadoOtroTurno As Double
   Screen.MousePointer = vbHourglass
   Set oComando = New clsComando: Set Reporte = New dsrLiquidacion
   If lNcOfisis Then
      If Not oComando.CreateCmdSp("spRep_LiquidacionOutPut_NC", Cn) Then Set oComando = Nothing: Exit Sub
   Else
      If Not oComando.CreateCmdSp("spRep_LiquidacionOutPut", Cn) Then Set oComando = Nothing: Exit Sub
   End If
   xparametros = ""
   oComando.CreateParameter "@flagTurno", adBoolean, adParamInput, 1, ChkTurno.Value
   oComando.CreateParameter "@flagDiaContable", adBoolean, adParamInput, 1, chkDiaContable.Value
   oComando.CreateParameter "@sturno", adVarChar, adParamInput, 20, sTurno
   oComando.CreateParameter "@sUsuario", adVarChar, adParamInput, 20, Me.cboUsuario.Text
   oComando.CreateParameter "@sSectorVenta", adVarChar, adParamInput, 50, Me.cboSectorVenta.BoundText
   oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
   oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
   oComando.CreateParameter "@xFecha", adDBDate, adParamInput, 10, xFecha
   oComando.CreateParameter "@cortesia", adInteger, adParamInput, 10, IIf(chkCortesia.Value, 1, 0)
   oComando.CreateParameter "@Ndolar", adDouble, adParamOutput, 10, nDolar
   oComando.CreateParameter "@xDOlar", adDouble, adParamOutput, 10, xDolar
   oComando.CreateParameter "@nOtroDoc", adDouble, adParamOutput, 10, nOtroDoc
   oComando.CreateParameter "@nNoCobrado", adDouble, adParamOutput, 10, nNoCobrado
   oComando.CreateParameter "@nNeto", adDouble, adParamOutput, 10, nNeto
   oComando.CreateParameter "@nImpuesto1", adDouble, adParamOutput, 10, nImpuesto1
   oComando.CreateParameter "@nImpuesto2", adDouble, adParamOutput, 10, nImpuesto2
   oComando.CreateParameter "@nImpuesto3", adDouble, adParamOutput, 10, nImpuesto3
   oComando.CreateParameter "@nVenta", adDouble, adParamOutput, 10, nTotal
   oComando.CreateParameter "@nDescuento", adDouble, adParamOutput, 10, nDescuento
   oComando.CreateParameter "@nRecargo", adDouble, adParamOutput, 10, nRecargo
   oComando.CreateParameter "@nCambio", adDouble, adParamOutput, 10, nCambio
   oComando.CreateParameter "@nAdulto", adDouble, adParamOutput, 10, nAdulto
   oComando.CreateParameter "@nNino", adDouble, adParamOutput, 10, nNino
   oComando.CreateParameter "@nAdulto2", adDouble, adParamOutput, 10, nAdulto2
   oComando.CreateParameter "@nNino2", adDouble, adParamOutput, 10, nNino2
   oComando.CreateParameter "@nAdulto3", adDouble, adParamOutput, 10, nAdulto3
   oComando.CreateParameter "@nNino3", adDouble, adParamOutput, 10, nNino3
   oComando.CreateParameter "@nAdulto4", adDouble, adParamOutput, 10, nAdulto4
   oComando.CreateParameter "@nNino4", adDouble, adParamOutput, 10, nNino4
   oComando.CreateParameter "@nAdulto5", adDouble, adParamOutput, 10, nAdulto5
   oComando.CreateParameter "@nNino5", adDouble, adParamOutput, 10, nNino5
   oComando.CreateParameter "@nPagadosEnOtro", adDouble, adParamOutput, 10, nNoCobrado

   If Not oComando.GetParamOK Then Set oComando = Nothing: Exit Sub
   Set rsReporte = oComando.GetSP()
   nDolar = oComando.GetParameterValue("@ndolar")
   xDolar = oComando.GetParameterValue("@xdolar")
   nOtroDoc = oComando.GetParameterValue("@nOtroDoc")
   nNoCobrado = oComando.GetParameterValue("@nNoCobrado")
   nNeto = oComando.GetParameterValue("@nNeto")
   nImpuesto1 = oComando.GetParameterValue("@nImpuesto1")
   nImpuesto2 = oComando.GetParameterValue("@nImpuesto2")
   nImpuesto3 = oComando.GetParameterValue("@nImpuesto3")
   nTotal = oComando.GetParameterValue("@nVenta")
   nDescuento = oComando.GetParameterValue("@nDescuento")
   nRecargo = oComando.GetParameterValue("@nRecargo")
   nCambio = oComando.GetParameterValue("@ncambio")
   nAdulto = oComando.GetParameterValue("@nadulto")
   nNino = oComando.GetParameterValue("@nnino")
   nAdulto2 = oComando.GetParameterValue("@nadulto2")
   nNino2 = oComando.GetParameterValue("@nnino2")
   nAdulto3 = oComando.GetParameterValue("@nadulto3")
   nNino3 = oComando.GetParameterValue("@nnino3")
   nAdulto4 = oComando.GetParameterValue("@nadulto4")
   nNino4 = oComando.GetParameterValue("@nnino4")
   nAdulto5 = oComando.GetParameterValue("@nadulto5")
   nNino5 = oComando.GetParameterValue("@nnino5")
   pagadoOtroTurno = oComando.GetParameterValue("@nPagadosEnOtro")

   tipo = "'1'"
   xparametros = " ," + str(Me.ChkTurno.Value) + "," + str(Me.chkDiaContable.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkCortesia.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + "," + str(Me.chkGenerado.Value) + ",'" + sTurno + "','" + cboUsuario.Text + "','" + sClub + "','" + Format(xFecha, "yyyy/mm/dd") + "','" + Format(fInicio, "yyyy/mm/dd HH:MM:ss") + "','" + Format(fFinal, "yyyy/mm/dd HH:MM:ss") + "', '" + cboSectorVenta.BoundText + "'"
   If lNcOfisis Then
      Set rsReporte = Lib.OpenRecordset("spRep_Liquidacion_NC" + tipo + xparametros, Cn)
   Else
      Set rsReporte = Lib.OpenRecordset("spRep_Liquidacion" + tipo + xparametros, Cn)
   End If

   Dim X As Long
   Dim TOTAL As Double
   TOTAL = 0
   rsReporte.Filter = "tgrupo='00'"
   If Not (rsReporte.BOF Or rsReporte.EOF) Then
      rsReporte.MoveFirst
      TOTAL = rsReporte.Fields("total00")
      tFiscal = IIf(IsNull(rsReporte.Fields("totalNF")), 0, rsReporte.Fields("totalNF"))
   End If
   nDocumento = TOTAL
   tipo = "'2'"

   If lNcOfisis Then
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion_NC" + tipo + xparametros, Cn)
   Else
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion" + tipo + xparametros, Cn)
   End If

   RsSumas.Filter = "tGrupo='01'"  'Efectivo
   If RsSumas.RecordCount > 0 Then
      nEfectivoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nEfectivoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xEfectivoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nEfectivoN = 0: nEfectivoE = 0: xEfectivoE = 0
   End If

   RsSumas.Filter = "tGrupo='03'"  'Cheques
   If RsSumas.RecordCount > 0 Then
      nChequeN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nChequeE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xChequeE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nChequeN = 0: nChequeE = 0: xChequeE = 0
   End If

   RsSumas.Filter = "tGrupo='04'"  'Otros
   If RsSumas.RecordCount > 0 Then
      nOtroN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nOtroE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xOtroE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nOtroN = 0: nOtroE = 0: xOtroE = 0
   End If

   RsSumas.Filter = "tGrupo='05'"  'Puntos
   If RsSumas.RecordCount > 0 Then
      nPunto = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
   Else
      nPunto = 0
   End If

   RsSumas.Filter = "tGrupo='06'"  'Cancelacion Cortesias
   If RsSumas.RecordCount > 0 Then
      nCortesia = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
   Else
      nCortesia = 0
   End If

   RsSumas.Filter = "tGrupo='07'"  'Cuentas por Cobrar
   If RsSumas.RecordCount > 0 Then
      nCobrar = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
   Else
      nCobrar = 0
   End If

   RsSumas.Filter = "tGrupo='10'"  'Recibos de egreso
   If RsSumas.RecordCount > 0 Then
      nEgresoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nEgresoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xEgresoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nEgresoN = 0: nEgresoE = 0: xEgresoE = 0
   End If

   If lNcOfisis Then
      RsSumas.Filter = "tGrupo='20'"  'Notas de CRedito
      If RsSumas.RecordCount > 0 Then
         nNotaCreditoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
         nNotaCreditoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
         xNotaCreditoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
      Else
         nNotaCreditoN = 0: nNotaCreditoE = 0: xNotaCreditoE = 0
      End If
   Else
      nNotaCreditoN = 0: nNotaCreditoE = 0: xNotaCreditoE = 0
   End If

   Set RsTarjetas = Lib.OpenRecordset("Select tCodigoTarjeta, tDetallado, ISNULL(nFactorRetencion,0) NFACTORRETENCION from TTARJETACREDITO where lActivo=1", Cn)  'Tarjetas
   For i = 1 To 24
       RsTarjetas.Filter = "tCodigoTarjeta='" & Right("00" & LTrim(str(i)), 2) & "'"
       If RsTarjetas.RecordCount > 0 Then
          tTarjeta(i) = IIf(IsNull(RsTarjetas!tDetallado), 0, RsTarjetas!tDetallado)
          nRetencion(i) = IIf(IsNull(RsTarjetas!nFactorRetencion), 0, RsTarjetas!nFactorRetencion)
       Else
          tTarjeta(i) = "": nRetencion(i) = 0
       End If
   Next i
   tipo = "'3'"

   If lNcOfisis Then
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion_NC" + tipo + xparametros, Cn)
   Else
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion" + tipo + xparametros, Cn)
   End If
   nTarjetaT = 0: nPropinaT = 0: nReciboT = 0
   For i = 1 To 24
      RsSumas.Filter = "tSubGrupo='" & Right("00" & LTrim(str(i)), 2) & "'"
      If RsSumas.RecordCount > 0 Then
         nTarjeta(i) = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
         nPropina(i) = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      Else
         nTarjeta(i) = 0: nPropina(i) = 0
      End If
      nTarjetaT = nTarjetaT + nTarjeta(i)
      nPropinaT = nPropinaT + nPropina(i)
   Next i

   tipo = "'4'"  'Recibos de ingreso Efectivo
   If lNcOfisis Then
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion_NC" + tipo + xparametros, Cn)
   Else
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion" + tipo + xparametros, Cn)
   End If

   RsSumas.Filter = "tgrupo='08'"
   If RsSumas.RecordCount > 0 Then
      nIngresoN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nIngresoE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xIngresoE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nIngresoN = 0: nIngresoE = 0: xIngresoE = 0
   End If

   RsSumas.Filter = "tgrupo='09'"  'Recibos de ingreso Anticipos Efectivo
   If RsSumas.RecordCount > 0 Then
      nIngresoAN = IIf(IsNull(RsSumas!nVenta1), 0, RsSumas!nVenta1)
      nIngresoAE = IIf(IsNull(RsSumas!nVenta2), 0, RsSumas!nVenta2)
      xIngresoAE = IIf(IsNull(RsSumas!nVenta3), 0, RsSumas!nVenta3)
   Else
      nIngresoAN = 0: nIngresoAE = 0: xIngresoAE = 0
   End If

   Set RsTarjetas = Lib.OpenRecordset("select tTarjeta, sum(nMonto) as nVenta1 From vIngreso where testadodocumento<>'04' and tTipoPago='02' and " & sCriterio & " Group by tTarjeta", Cn)  'Ingresos por Tarjeta
   For i = 1 To 24
        If i < 10 Then
            RsTarjetas.Filter = "tTarjeta='0" & LTrim(str(i)) & "'"
        Else
            RsTarjetas.Filter = "tTarjeta='" & LTrim(str(i)) & "'"
        End If
       
       If RsTarjetas.RecordCount > 0 Then
          nRecibo(i) = IIf(IsNull(RsTarjetas!nVenta1), 0, RsTarjetas!nVenta1)
       Else
          nRecibo(i) = 0
       End If
       nReciboT = nReciboT + nRecibo(i)
   Next i

   If lNcOfisis Then  'Totales por Tipo Pedido
      If Not oComando.CreateCmdSp("spRep_LiquidacionSuma_NC", Cn) Then Set oComando = Nothing: Exit Sub
   Else
      If Not oComando.CreateCmdSp("spRep_LiquidacionSuma", Cn) Then Set oComando = Nothing: Exit Sub
   End If

   Set RsSumas = Nothing
   oComando.CreateParameter "@flagTurno", adBoolean, adParamInput, 1, ChkTurno.Value
   oComando.CreateParameter "@flagDiaContable", adBoolean, adParamInput, 1, chkDiaContable.Value
   oComando.CreateParameter "@sturno", adVarChar, adParamInput, 20, sTurno
   oComando.CreateParameter "@sUsuario", adVarChar, adParamInput, 20, Me.cboUsuario.Text
   oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
   oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
   oComando.CreateParameter "@sSectorVenta", adVarChar, adParamInput, 50, Me.cboSectorVenta.BoundText

   If Not oComando.GetParamOK Then Set oComando = Nothing: Exit Sub
   Set RsSumas = oComando.GetSP()
   Reporte.Local.SetText IIf(sBoton1 <> "", sBoton1, "n/d") & " :"
   Reporte.Delivery.SetText IIf(sBoton2 <> "", sBoton2, "n/d") & " :"
   Reporte.Llevar.SetText IIf(sBoton3 <> "", sBoton3, "n/d") & " :"
   Reporte.Canal4.SetText IIf(sBoton4 <> "", sBoton4, "n/d") & " :"
   Reporte.Canal5.SetText IIf(sBoton5 <> "", sBoton5, "n/d") & " :"
   Reporte.Text20.SetText localConectado

   Reporte.Canal1x.SetText IIf(sBoton1 <> "", sBoton1, "n/d") & " :"
   Reporte.Canal2x.SetText IIf(sBoton2 <> "", sBoton2, "n/d") & " :"
   Reporte.Canal3x.SetText IIf(sBoton3 <> "", sBoton3, "n/d") & " :"
   Reporte.Canal4x.SetText IIf(sBoton4 <> "", sBoton4, "n/d") & " :"
   Reporte.Canal5x.SetText IIf(sBoton5 <> "", sBoton5, "n/d") & " :"
      
   Dim xTotal As Double
   xTotal = 0
   RsSumas.Filter = "ttipopedido='01'"
   If Not RsSumas.EOF Then
      Reporte.VV1.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto), "###,##0.00")
      Reporte.IMP11.SetText Format(IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1), "###,##0.00")
      Reporte.IMP21.SetText Format(IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2), "###,##0.00")
      Reporte.IMP31.SetText Format(IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3), "###,##0.00")
      Reporte.TOT1.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta), "###,##0.00")
      nTotalPromedio = Format(IIf(IsNull(RsSumas!nTotalPromedio), 0, RsSumas!nTotalPromedio), "###,##0.00") ' Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='01'", Cn)
      Reporte.DOCUM1.SetText Format(nTotalPromedio, "###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR1.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
      Else
         Reporte.FACTOR1.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
      End If
      Reporte.nAdulto.SetText Format(nAdulto, "###,###,##0")
      Reporte.nNino.SetText Format(nNino, "###,###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nAdulto = 0, 1, nAdulto), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      Else
         Reporte.FACTOR.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nAdulto = 0, 1, nAdulto), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      End If
   End If
   
   RsSumas.Filter = "ttipopedido='02'"
   If Not RsSumas.EOF Then
      Reporte.VV2.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto), "###,##0.00")
      Reporte.IMP12.SetText Format(IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1), "###,##0.00")
      Reporte.IMP22.SetText Format(IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2), "###,##0.00")
      Reporte.IMP32.SetText Format(IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3), "###,##0.00")
      Reporte.TOT2.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta), "###,##0.00")
      nTotalPromedio = Format(IIf(IsNull(RsSumas!nTotalPromedio), 0, RsSumas!nTotalPromedio), "###,##0.00") 'Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='02'", Cn)
      Reporte.DOCUM2.SetText Format(nTotalPromedio, "###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR2.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      Else
         Reporte.FACTOR2.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      End If
      Reporte.nAdulto2.SetText Format(nAdulto2, "###,###,##0")
      Reporte.nNino2.SetText Format(nNino2, "###,###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR2x.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nAdulto2 = 0, 1, nAdulto2), "###,###,##0.00")
      Else
         Reporte.FACTOR2x.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nAdulto2 = 0, 1, nAdulto2), "###,###,##0.00")
      End If
   End If

   RsSumas.Filter = "ttipopedido='03'"
   If Not RsSumas.EOF Then
      Reporte.VV3.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto), "###,##0.00")
      Reporte.IMP13.SetText Format(IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1), "###,##0.00")
      Reporte.IMP23.SetText Format(IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2), "###,##0.00")
      Reporte.IMP33.SetText Format(IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3), "###,##0.00")
      Reporte.TOT3.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta), "###,##0.00")
      nTotalPromedio = Format(IIf(IsNull(RsSumas!nTotalPromedio), 0, RsSumas!nTotalPromedio), "###,##0.00") ' Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='03'", Cn)
      Reporte.DOCUM3.SetText Format(nTotalPromedio, "###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR3.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      Else
         Reporte.FACTOR3.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      End If
      Reporte.nAdulto3.SetText Format(nAdulto3, "###,###,##0")
      Reporte.nNino3.SetText Format(nNino3, "###,###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR3x.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nAdulto3 = 0, 1, nAdulto3), "###,###,##0.00")
      Else
         Reporte.FACTOR3x.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nAdulto3 = 0, 1, nAdulto3), "###,###,##0.00")
      End If
   End If
                                                        
   RsSumas.Filter = "ttipopedido='04'"
   If Not RsSumas.EOF Then
      Reporte.VV4.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto), "###,##0.00")
      Reporte.IMP14.SetText Format(IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1), "###,##0.00")
      Reporte.IMP24.SetText Format(IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2), "###,##0.00")
      Reporte.IMP34.SetText Format(IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3), "###,##0.00")
      Reporte.TOT4.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta), "###,##0.00")
      nTotalPromedio = Format(IIf(IsNull(RsSumas!nTotalPromedio), 0, RsSumas!nTotalPromedio), "###,##0.00") 'Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='04'", Cn)
      Reporte.DOCUM4.SetText Format(nTotalPromedio, "###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR4.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
         Reporte.FACTOR4x.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nAdulto4 = 0, 1, nAdulto4), "###,###,##0.00")
      Else
         Reporte.FACTOR4.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
         Reporte.FACTOR4x.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nAdulto4 = 0, 1, nAdulto4), "###,###,##0.00")
      End If
      Reporte.nAdulto4.SetText Format(nAdulto4, "###,###,##0")
      Reporte.nNino4.SetText Format(nNino4, "###,###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR4x.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nAdulto4 = 0, 1, nAdulto4), "###,###,##0.00")
      Else
         Reporte.FACTOR4x.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nAdulto4 = 0, 1, nAdulto4), "###,###,##0.00")
      End If
   End If
                                                        
   RsSumas.Filter = "ttipopedido='05'"
   If Not RsSumas.EOF Then
      Reporte.VV5.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto), "###,##0.00")
      Reporte.IMP15.SetText Format(IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1), "###,##0.00")
      Reporte.IMP25.SetText Format(IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2), "###,##0.00")
      Reporte.IMP35.SetText Format(IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3), "###,##0.00")
      Reporte.TOT5.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta), "###,##0.00")
      nTotalPromedio = Format(IIf(IsNull(RsSumas!nTotalPromedio), 0, RsSumas!nTotalPromedio), "###,##0.00") ' Calcular("select count(tDocumento) as Codigo from " & sSuma & " where tTipoPedido='05'", Cn)
      Reporte.DOCUM5.SetText Format(nTotalPromedio, "###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR5.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
         Reporte.FACTOR5x.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nAdulto5 = 0, 1, nAdulto5), "###,###,##0.00")
      Else
         Reporte.FACTOR5.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
         Reporte.FACTOR5x.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nAdulto5 = 0, 1, nAdulto5), "###,###,##0.00")
      End If
      Reporte.nAdulto5.SetText Format(nAdulto5, "###,###,##0")
      Reporte.nNino5.SetText Format(nNino5, "###,###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR5x.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nAdulto5 = 0, 1, nAdulto5), "###,###,##0.00")
      Else
         Reporte.FACTOR5x.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nAdulto5 = 0, 1, nAdulto5), "###,###,##0.00")
      End If
   End If
   
   Dim ntotalPromedioFiscal, valorventaf As Double
   RsSumas.Filter = "ttipopedido='XX'"
   If Not RsSumas.EOF Then
      Reporte.VV6.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto), "###,##0.00")
      Reporte.IMP16.SetText Format(IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1), "###,##0.00")
      Reporte.IMP26.SetText Format(IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2), "###,##0.00")
      Reporte.IMP36.SetText Format(IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3), "###,##0.00")
      Reporte.TOT6.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta), "###,##0.00")
      nTotalPromedio = Format(IIf(IsNull(RsSumas!nTotalPromedio), 0, RsSumas!nTotalPromedio), "###,##0.00")
      ntotalPromedioFiscal = Format(IIf(IsNull(RsSumas!nTotalPromedio), 0, RsSumas!nTotalPromedio), "###,##0.00")
      Reporte.DOCUM6.SetText Format(nTotalPromedio, "###,##0")
      If optVenta.Value = False Then
         Reporte.FACTOR6.SetText Format(IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
         valorventaf = IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto)
      Else
         Reporte.FACTOR6.SetText Format(IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta) / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
         xTotal = xTotal + IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
         valorventaf = IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta)
      End If
      Reporte.VALORVENTAT.SetText Format(nNeto + IIf(IsNull(RsSumas!nNeto), 0, RsSumas!nNeto), "###,###,##0.00")
      Reporte.IMPUESTO1T.SetText Format(nImpuesto1 + IIf(IsNull(RsSumas!nImpuesto1), 0, RsSumas!nImpuesto1), "###,###,##0.00")
      Reporte.IMPUESTO2T.SetText Format(nImpuesto2 + IIf(IsNull(RsSumas!nImpuesto2), 0, RsSumas!nImpuesto2), "###,###,##0.00")
      Reporte.IMPUESTO3T.SetText Format(nImpuesto3 + IIf(IsNull(RsSumas!nImpuesto3), 0, RsSumas!nImpuesto3), "###,###,##0.00")
      Reporte.TOTALT.SetText Format(nTotal + IIf(IsNull(RsSumas!nVenta), 0, RsSumas!nVenta), "###,###,##0.00")
    Else
      Reporte.VALORVENTAT.SetText Format(nNeto, "###,###,##0.00")
      Reporte.IMPUESTO1T.SetText Format(nImpuesto1, "###,###,##0.00")
      Reporte.IMPUESTO2T.SetText Format(nImpuesto2, "###,###,##0.00")
      Reporte.IMPUESTO3T.SetText Format(nImpuesto3, "###,###,##0.00")
      Reporte.TOTALT.SetText Format(nTotal, "###,###,##0.00")
   End If
   
   Dim subtotal As Double
   RsSumas.Filter = ""
   TOTAL = 0
   If Not (RsSumas.EOF Or RsSumas.BOF) Then
      RsSumas.MoveFirst
      nTotalPromedio = RsSumas!total00
   Else
      nTotalPromedio = TOTAL
   End If
   ntotalPromedioFiscal = IIf(IsNull(ntotalPromedioFiscal), 0, ntotalPromedioFiscal)
   ntotalPromedioFiscal = IIf(nTotalPromedio = 0, 1, nTotalPromedio - ntotalPromedioFiscal)
   subtotal = xTotal - valorventaf
   If Not (RsSumas.EOF Or RsSumas.BOF) Then
      RsSumas.MoveFirst
      nTotalPromedio = RsSumas!total00
   Else
      nTotalPromedio = TOTAL
   End If

   Reporte.DOCUMT.SetText Format(ntotalPromedioFiscal, "###,##0")
   Reporte.FACTORT.SetText Format(subtotal / IIf(ntotalPromedioFiscal = 0, 1, ntotalPromedioFiscal), "###,###,##0.00")
   Reporte.DOCUMTT.SetText Format(nTotalPromedio, "###,##0")
   Reporte.FACTORTT.SetText Format(xTotal / IIf(nTotalPromedio = 0, 1, nTotalPromedio), "###,###,##0.00")
   Call LLenarOrigenventa
   tipo = "'5'"
   If lDesactivaNCFP Then  'Otros Tipos de Pago
      Isql = "select * from vTipoCancelacion where lActivo=1 AND CODIGO<>'002'"
   Else
      Isql = "select * from vTipoCancelacion where lActivo=1"
   End If
   
   If lNcOfisis Then
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion" + tipo + xparametros, Cn)
   Else
      Set RsSumas = Lib.OpenRecordset("spRep_Liquidacion" + tipo + xparametros, Cn)
   End If
   Set RsTarjetas = Lib.OpenRecordset(Isql, Cn)
   
   If RsTarjetas.RecordCount > 0 Then
      RsTarjetas.MoveFirst
      For i = 1 To 20
        If RsTarjetas.EOF Or RsTarjetas.BOF Then
                 tOtro(i) = ""
                 nOtro(i) = 0
        Else
            RsSumas.Filter = "CODIGO='" & RsTarjetas.Fields("codigo") & "'"
            If RsSumas.EOF Then
                tOtro(i) = IIf(IsNull(RsTarjetas!Descripcion), "", RsTarjetas!Descripcion)
                nOtro(i) = 0
            Else
                tOtro(i) = RsTarjetas!Descripcion
                nOtro(i) = RsSumas.Fields("codigo1") '("select sum(nVenta1) as Codigo from " & sNombre & " where tGrupo='04' and tSubGrupo='" & RsSumas!Codigo & "'", Cn)
                xOtro(i) = RsSumas.Fields("codigo2") 'Calcular("select sum(nVenta2) as Codigo from " & sNombre & " where tGrupo='04' and tSubGrupo='" & RsSumas!Codigo & "'", Cn)
                nDiferencia(i) = RsSumas.Fields("diferencia") ' Calcular("select sum(case when round(nVenta1,2) > round(nVenta,2) then nVenta1 - nVenta else 0 end) as Codigo from " & sNombre & " where tGrupo='04' and tSubGrupo='" & RsSumas!Codigo & "'", Cn)
            End If
             RsTarjetas.MoveNext
        End If
      Next i
   End If

   nTotalCobradoN = nEfectivoN + nTarjetaT + nOtroN + nChequeN + nPunto + nCortesia + nNotaCreditoN
   nTotalCobradoE = nEfectivoE + nOtroE + nChequeE + nNotaCreditoE
   xTotalCobradoE = xEfectivoE + xOtroE + xChequeE + xNotaCreditoE
   nTotalCobrado = nTotalCobradoN + xTotalCobradoE

   TOTAL = 0
   rsReporte.Filter = ""
    rsReporte.Filter = "tgrupo='19'"
   If Not (rsReporte.EOF Or rsReporte.BOF) Then
      rsReporte.MoveFirst
      For i = 0 To rsReporte.RecordCount - 1
          TOTAL = TOTAL + IIf(IsNull(rsReporte.Fields("nventa")), 0, rsReporte.Fields("nventa"))
          rsReporte.MoveNext
      Next i
   Else
      TOTAL = 0
   End If
   nFacCPedido = TOTAL
   TOTAL = 0
   rsReporte.Filter = ""

   rsReporte.Filter = "tgrupo='14'"
   If Not (rsReporte.EOF Or rsReporte.BOF) Then
      rsReporte.MoveFirst
      TOTAL = TOTAL + IIf(IsNull(rsReporte.Fields("total14")), 0, rsReporte.Fields("total14"))
   Else
      TOTAL = 0
   End If
    
   nOtroPed = TOTAL ' Calcular("select sum(nVenta1) as Codigo from " & sNombre & " where tGrupo='14'", Cn)
   rsReporte.Filter = ""

   Reporte.Text14.SetText sMonedaN
   Reporte.Text18.SetText sMonedaE
   Reporte.Text9.SetText sMonedaN
   Reporte.Text16.SetText sMonedaE
   Reporte.Text36.SetText sMonedaN
   Reporte.Text40.SetText sMonedaE
   Reporte.Text59.SetText sMonedaN
   Reporte.Text60.SetText sMonedaE
   Reporte.txtsoles.SetText sMonedaN
   Reporte.txtdolares.SetText sMonedaE
   
   Reporte.TipoCambio.SetText Format(nCambio, "###,###,##0.000")
   Reporte.EFETOT.SetText Format(nEfectivoN + xEfectivoE, "###,###,##0.00")
   Reporte.EfeSTot.SetText Format(nEfectivoN, "###,###,##0.00")
   Reporte.EfeDTot.SetText Format(nEfectivoE, "###,###,##0.00")
   Reporte.TOTTAR.SetText Format(nTarjetaT, "###,###,##0.00")
   Reporte.TARS.SetText Format(nTarjetaT, "###,###,##0.00")
   Reporte.TotOtro.SetText Format(nOtroN + xOtroE, "###,###,##0.00")
   Reporte.OTRON.SetText Format(nOtroN, "###,###,##0.00")
   Reporte.OTROE.SetText Format(nOtroE, "###,###,##0.00")
   Reporte.TOTCHE.SetText Format(nChequeN + xChequeE, "###,###,##0.00")
   Reporte.CHEQUEN.SetText Format(nChequeN, "###,###,##0.00")
   Reporte.CHEQUEE.SetText Format(nChequeE, "###,###,##0.00")
   Reporte.PUNTO.SetText Format(nPunto, "###,###,##0.00")
   Reporte.PUNTOS.SetText Format(nPunto, "###,###,##0.00")
   Reporte.Cortesia.SetText Format(nCortesia, "###,###,##0.00")  'Cortesias

   If lNcOfisis Then
      Reporte.Text77.SetText "( 20 )   Notas  De Credito : "
      Reporte.Text80.SetText "(01+02+03+04+05+06+20) "
      Reporte.Text41.SetText "( 01 + 08 + 09 - 10 + 20 )"
      Reporte.Text81.SetText "( 01 + 02 + 03 + 04 + 05 + 06 + 20 + Propinas) "
      Reporte.NotaCreditoT.SetText Format(nNotaCreditoN + xNotaCreditoE, "###,###,##0.00")
      Reporte.NotaCreditoN.SetText Format(nNotaCreditoN, "###,###,##0.00")
      Reporte.NotaCreditoE.SetText Format(nNotaCreditoE, "###,###,##0.00")
   Else
      Reporte.Text80.SetText "(01+02+03+04+05+06) "
      Reporte.Text81.SetText "( 01 + 02 + 03 + 04 + 05 + 06+ Propinas) "
      Reporte.Text77.SetText ""
      Reporte.Text41.SetText "( 01 + 08 + 09 - 10)"
      Reporte.NotaCreditoT.SetText ""
      Reporte.NotaCreditoN.SetText ""
      Reporte.NotaCreditoE.SetText ""
   End If
   Reporte.totCob.SetText Format(nTotalCobrado, "###,###,##0.00")
   Reporte.COBRAN.SetText Format(nTotalCobradoN, "###,###,##0.00")
   Reporte.COBRAE.SetText Format(nTotalCobradoE, "###,###,##0.00")
   Reporte.TotalCobrado.SetText Format(nTotalCobrado + nPropinaT, "###,###,##0.00")
   Reporte.OtroTurno.SetText Format(nOtroDoc, "###,###,##0.00")
   Reporte.Cobrar.SetText Format(nCobrar, "###,###,##0.00")
   Reporte.NOCOBRADO.SetText Format(nNoCobrado + (nDocumento - nTotalCobrado - nCobrar - nNoCobrado + nOtroDoc), "###,###,##0.00")
   Reporte.TotDoc.SetText Format(nDocumento, "###,###,##0.00")
   Reporte.Text19.SetText Format(nOtroPed, "###,###,##0.00")
   Reporte.Text73.SetText Format(nFacCPedido, "###,###,##0.00")

   If sImpuesto1 <> "" Then
      Reporte.Text51.SetText sImpuesto1
   Else
      Reporte.Text51.SetText ""
   End If
   
   If sImpuesto2 <> "" Then
      Reporte.Text52.SetText sImpuesto2
   Else
      Reporte.Text52.SetText ""
   End If
   
   If sImpuesto3 <> "" Then
      Reporte.Text53.SetText sImpuesto3
   Else
      Reporte.Text53.SetText ""
   End If
      
   Reporte.VALORVENTA.SetText Format(nNeto, "###,###,##0.00")
   Reporte.IMPUESTO1.SetText Format(nImpuesto1, "###,###,##0.00")
   Reporte.IMPUESTO2.SetText Format(nImpuesto2, "###,###,##0.00")
   Reporte.IMPUESTO3.SetText Format(nImpuesto3, "###,###,##0.00")
   Reporte.TOTAL.SetText Format(nTotal, "###,###,##0.00")
   Reporte.DESCUENTO.SetText Format(nRecargo - nDescuento, "###,###,##0.00")

   Reporte.TOTING.SetText Format(nIngresoN + xIngresoE, "###,###,##0.00")
   Reporte.INGN.SetText Format(nIngresoN, "###,###,##0.00")
   Reporte.INGE.SetText Format(nIngresoE, "###,###,##0.00")
   Reporte.TOTANT.SetText Format(nIngresoAN + xIngresoAE, "###,###,##0.00")
   Reporte.ANTN.SetText Format(nIngresoAN, "###,###,##0.00")
   Reporte.ANTE.SetText Format(nIngresoAE, "###,###,##0.00")
   Reporte.TOTEGR.SetText Format(nEgresoN + xEgresoE, "###,###,##0.00")
   Reporte.EGRN.SetText Format(nEgresoN, "###,###,##0.00")
   Reporte.EGRE.SetText Format(nEgresoE, "###,###,##0.00")

   nTotalEfectivoN = nEfectivoN + nIngresoN + nIngresoAN - nEgresoN '+ nNotaCreditoN
   nTotalEfectivoE = nEfectivoE + nIngresoE + nIngresoAE - nEgresoE + nNotaCreditoE
   xTotalEfectivoE = xEfectivoE + xIngresoE + xIngresoAE - xEgresoE + xNotaCreditoE
   nTotalEfectivo = nTotalEfectivoN + xTotalEfectivoE
   nFisEfectivoN = nTotalEfectivoN - (xDolar - xEgresoE - xTotalEfectivoE)
   nFisEfectivo = nFisEfectivoN + xDolar - xEgresoE
   Reporte.TOTFON.SetText Format(nFondoN + (nFondoE * nCambio), "###,###,##0.00")
   Reporte.FondoN.SetText Format(nFondoN, "###,###,##0.00")
   Reporte.FondoE.SetText Format(nFondoE, "###,###,##0.00")
   Reporte.TOTALPRO.SetText Format(nTotalEfectivo, "###,###,##0.00")
   Reporte.PRON.SetText Format(nTotalEfectivoN, "###,###,##0.00")
   Reporte.PROE.SetText Format(nTotalEfectivoE, "###,###,##0.00")
   Reporte.TOTALFIS.SetText Format(nFisEfectivo, "###,###,##0.00")
   Reporte.FISN.SetText Format(nFisEfectivoN, "###,###,##0.00")
   Reporte.FISE.SetText Format(nDolar - nEgresoE, "###,###,##0.00")
   
'*********************************tarjetas******************************************
   Reporte.Tarjeta1.SetText tTarjeta(1)
   Reporte.Tarjeta2.SetText tTarjeta(2)
   Reporte.Tarjeta3.SetText tTarjeta(3)
   Reporte.Tarjeta4.SetText tTarjeta(4)
   Reporte.Tarjeta5.SetText tTarjeta(5)
   Reporte.Tarjeta6.SetText tTarjeta(6)
   Reporte.Tarjeta7.SetText tTarjeta(7)
   Reporte.Tarjeta8.SetText tTarjeta(8)
   Reporte.Tarjeta9.SetText tTarjeta(9)
   Reporte.Tarjeta10.SetText tTarjeta(10)
   Reporte.Tarjeta11.SetText tTarjeta(11)
   Reporte.Tarjeta12.SetText tTarjeta(12)
   Reporte.Tarjeta13.SetText tTarjeta(13)
   Reporte.Tarjeta14.SetText tTarjeta(14)
   Reporte.Tarjeta15.SetText tTarjeta(15)
   Reporte.Tarjeta16.SetText tTarjeta(16)
   Reporte.Tarjeta17.SetText tTarjeta(17)
   Reporte.Tarjeta18.SetText tTarjeta(18)
   Reporte.Tarjeta19.SetText tTarjeta(19)
   Reporte.Tarjeta20.SetText tTarjeta(20)
   Reporte.Tarjeta21.SetText tTarjeta(21)
   Reporte.Tarjeta22.SetText tTarjeta(22)
   Reporte.Tarjeta23.SetText tTarjeta(23)
   Reporte.Tarjeta24.SetText tTarjeta(24)
  
   Reporte.Monto1.SetText IIf(tTarjeta(1) = "", "", Format(nTarjeta(1), "###,###,##0.00"))
   Reporte.Monto2.SetText IIf(tTarjeta(2) = "", "", Format(nTarjeta(2), "###,###,##0.00"))
   Reporte.Monto3.SetText IIf(tTarjeta(3) = "", "", Format(nTarjeta(3), "###,###,##0.00"))
   Reporte.Monto4.SetText IIf(tTarjeta(4) = "", "", Format(nTarjeta(4), "###,###,##0.00"))
   Reporte.Monto5.SetText IIf(tTarjeta(5) = "", "", Format(nTarjeta(5), "###,###,##0.00"))
   Reporte.Monto6.SetText IIf(tTarjeta(6) = "", "", Format(nTarjeta(6), "###,###,##0.00"))
   Reporte.Monto7.SetText IIf(tTarjeta(7) = "", "", Format(nTarjeta(7), "###,###,##0.00"))
   Reporte.Monto8.SetText IIf(tTarjeta(8) = "", "", Format(nTarjeta(8), "###,###,##0.00"))
   Reporte.Monto9.SetText IIf(tTarjeta(9) = "", "", Format(nTarjeta(9), "###,###,##0.00"))
   Reporte.Monto10.SetText IIf(tTarjeta(10) = "", "", Format(nTarjeta(10), "###,###,##0.00"))
   Reporte.Monto11.SetText IIf(tTarjeta(11) = "", "", Format(nTarjeta(11), "###,###,##0.00"))
   Reporte.Monto12.SetText IIf(tTarjeta(12) = "", "", Format(nTarjeta(12), "###,###,##0.00"))
   Reporte.Monto13.SetText IIf(tTarjeta(13) = "", "", Format(nTarjeta(13), "###,###,##0.00"))
   Reporte.Monto14.SetText IIf(tTarjeta(14) = "", "", Format(nTarjeta(14), "###,###,##0.00"))
   Reporte.Monto15.SetText IIf(tTarjeta(15) = "", "", Format(nTarjeta(15), "###,###,##0.00"))
   Reporte.Monto16.SetText IIf(tTarjeta(16) = "", "", Format(nTarjeta(16), "###,###,##0.00"))
   Reporte.Monto17.SetText IIf(tTarjeta(17) = "", "", Format(nTarjeta(17), "###,###,##0.00"))
   Reporte.Monto18.SetText IIf(tTarjeta(18) = "", "", Format(nTarjeta(18), "###,###,##0.00"))
   Reporte.Monto19.SetText IIf(tTarjeta(19) = "", "", Format(nTarjeta(19), "###,###,##0.00"))
   Reporte.Monto20.SetText IIf(tTarjeta(20) = "", "", Format(nTarjeta(20), "###,###,##0.00"))
   Reporte.Monto21.SetText IIf(tTarjeta(21) = "", "", Format(nTarjeta(21), "###,###,##0.00"))
   Reporte.Monto22.SetText IIf(tTarjeta(22) = "", "", Format(nTarjeta(22), "###,###,##0.00"))
   Reporte.Monto23.SetText IIf(tTarjeta(23) = "", "", Format(nTarjeta(23), "###,###,##0.00"))
   Reporte.Monto24.SetText IIf(tTarjeta(24) = "", "", Format(nTarjeta(24), "###,###,##0.00"))
   
   Reporte.Prop1.SetText IIf(tTarjeta(1) = "", "", Format(nPropina(1), "###,###,##0.00"))
   Reporte.Prop2.SetText IIf(tTarjeta(2) = "", "", Format(nPropina(2), "###,###,##0.00"))
   Reporte.Prop3.SetText IIf(tTarjeta(3) = "", "", Format(nPropina(3), "###,###,##0.00"))
   Reporte.Prop4.SetText IIf(tTarjeta(4) = "", "", Format(nPropina(4), "###,###,##0.00"))
   Reporte.Prop5.SetText IIf(tTarjeta(5) = "", "", Format(nPropina(5), "###,###,##0.00"))
   Reporte.Prop6.SetText IIf(tTarjeta(6) = "", "", Format(nPropina(6), "###,###,##0.00"))
   Reporte.Prop7.SetText IIf(tTarjeta(7) = "", "", Format(nPropina(7), "###,###,##0.00"))
   Reporte.Prop8.SetText IIf(tTarjeta(8) = "", "", Format(nPropina(8), "###,###,##0.00"))
   Reporte.Prop9.SetText IIf(tTarjeta(9) = "", "", Format(nPropina(9), "###,###,##0.00"))
   Reporte.Prop10.SetText IIf(tTarjeta(10) = "", "", Format(nPropina(10), "###,###,##0.00"))
   Reporte.Prop11.SetText IIf(tTarjeta(11) = "", "", Format(nPropina(11), "###,###,##0.00"))
   Reporte.Prop12.SetText IIf(tTarjeta(12) = "", "", Format(nPropina(12), "###,###,##0.00"))
   Reporte.Prop13.SetText IIf(tTarjeta(13) = "", "", Format(nPropina(13), "###,###,##0.00"))
   Reporte.Prop14.SetText IIf(tTarjeta(14) = "", "", Format(nPropina(14), "###,###,##0.00"))
   Reporte.Prop15.SetText IIf(tTarjeta(15) = "", "", Format(nPropina(15), "###,###,##0.00"))
   Reporte.Prop16.SetText IIf(tTarjeta(16) = "", "", Format(nPropina(16), "###,###,##0.00"))
   Reporte.Prop17.SetText IIf(tTarjeta(17) = "", "", Format(nPropina(17), "###,###,##0.00"))
   Reporte.Prop18.SetText IIf(tTarjeta(18) = "", "", Format(nPropina(18), "###,###,##0.00"))
   Reporte.Prop19.SetText IIf(tTarjeta(19) = "", "", Format(nPropina(19), "###,###,##0.00"))
   Reporte.Prop20.SetText IIf(tTarjeta(20) = "", "", Format(nPropina(20), "###,###,##0.00"))
   Reporte.Prop21.SetText IIf(tTarjeta(21) = "", "", Format(nPropina(21), "###,###,##0.00"))
   Reporte.Prop22.SetText IIf(tTarjeta(22) = "", "", Format(nPropina(22), "###,###,##0.00"))
   Reporte.Prop23.SetText IIf(tTarjeta(23) = "", "", Format(nPropina(23), "###,###,##0.00"))
   Reporte.Prop24.SetText IIf(tTarjeta(24) = "", "", Format(nPropina(24), "###,###,##0.00"))

   
   Reporte.Rec1.SetText IIf(tTarjeta(1) = "", "", Format(nRecibo(1), "###,###,##0.00"))
   Reporte.Rec2.SetText IIf(tTarjeta(2) = "", "", Format(nRecibo(2), "###,###,##0.00"))
   Reporte.Rec3.SetText IIf(tTarjeta(3) = "", "", Format(nRecibo(3), "###,###,##0.00"))
   Reporte.Rec4.SetText IIf(tTarjeta(4) = "", "", Format(nRecibo(4), "###,###,##0.00"))
   Reporte.Rec5.SetText IIf(tTarjeta(5) = "", "", Format(nRecibo(5), "###,###,##0.00"))
   Reporte.Rec6.SetText IIf(tTarjeta(6) = "", "", Format(nRecibo(6), "###,###,##0.00"))
   Reporte.Rec7.SetText IIf(tTarjeta(7) = "", "", Format(nRecibo(7), "###,###,##0.00"))
   Reporte.Rec8.SetText IIf(tTarjeta(8) = "", "", Format(nRecibo(8), "###,###,##0.00"))
   Reporte.Rec9.SetText IIf(tTarjeta(9) = "", "", Format(nRecibo(9), "###,###,##0.00"))
   Reporte.Rec10.SetText IIf(tTarjeta(10) = "", "", Format(nRecibo(10), "###,###,##0.00"))
   Reporte.Rec11.SetText IIf(tTarjeta(11) = "", "", Format(nRecibo(11), "###,###,##0.00"))
   Reporte.Rec12.SetText IIf(tTarjeta(12) = "", "", Format(nRecibo(12), "###,###,##0.00"))
   Reporte.Rec13.SetText IIf(tTarjeta(13) = "", "", Format(nRecibo(13), "###,###,##0.00"))
   Reporte.Rec14.SetText IIf(tTarjeta(14) = "", "", Format(nRecibo(14), "###,###,##0.00"))
   Reporte.Rec15.SetText IIf(tTarjeta(15) = "", "", Format(nRecibo(15), "###,###,##0.00"))
   Reporte.Rec16.SetText IIf(tTarjeta(16) = "", "", Format(nRecibo(16), "###,###,##0.00"))
   Reporte.Rec17.SetText IIf(tTarjeta(17) = "", "", Format(nRecibo(17), "###,###,##0.00"))
   Reporte.Rec18.SetText IIf(tTarjeta(18) = "", "", Format(nRecibo(18), "###,###,##0.00"))
   Reporte.Rec19.SetText IIf(tTarjeta(19) = "", "", Format(nRecibo(19), "###,###,##0.00"))
   Reporte.Rec20.SetText IIf(tTarjeta(20) = "", "", Format(nRecibo(20), "###,###,##0.00"))
   Reporte.Rec21.SetText IIf(tTarjeta(21) = "", "", Format(nRecibo(21), "###,###,##0.00"))
   Reporte.Rec22.SetText IIf(tTarjeta(22) = "", "", Format(nRecibo(22), "###,###,##0.00"))
   Reporte.Rec23.SetText IIf(tTarjeta(23) = "", "", Format(nRecibo(23), "###,###,##0.00"))
   Reporte.Rec24.SetText IIf(tTarjeta(24) = "", "", Format(nRecibo(24), "###,###,##0.00"))
   
   Reporte.Total1.SetText IIf(tTarjeta(1) = "", "", Format(nTarjeta(1) + nPropina(1) + nRecibo(1), "###,###,##0.00"))
   Reporte.Total2.SetText IIf(tTarjeta(2) = "", "", Format(nTarjeta(2) + nPropina(2) + nRecibo(2), "###,###,##0.00"))
   Reporte.Total3.SetText IIf(tTarjeta(3) = "", "", Format(nTarjeta(3) + nPropina(3) + nRecibo(3), "###,###,##0.00"))
   Reporte.Total4.SetText IIf(tTarjeta(4) = "", "", Format(nTarjeta(4) + nPropina(4) + nRecibo(4), "###,###,##0.00"))
   Reporte.Total5.SetText IIf(tTarjeta(5) = "", "", Format(nTarjeta(5) + nPropina(5) + nRecibo(5), "###,###,##0.00"))
   Reporte.Total6.SetText IIf(tTarjeta(6) = "", "", Format(nTarjeta(6) + nPropina(6) + nRecibo(6), "###,###,##0.00"))
   Reporte.Total7.SetText IIf(tTarjeta(7) = "", "", Format(nTarjeta(7) + nPropina(7) + nRecibo(7), "###,###,##0.00"))
   Reporte.Total8.SetText IIf(tTarjeta(8) = "", "", Format(nTarjeta(8) + nPropina(8) + nRecibo(8), "###,###,##0.00"))
   Reporte.Total9.SetText IIf(tTarjeta(9) = "", "", Format(nTarjeta(9) + nPropina(9) + nRecibo(9), "###,###,##0.00"))
   Reporte.Total10.SetText IIf(tTarjeta(10) = "", "", Format(nTarjeta(10) + nPropina(10) + nRecibo(10), "###,###,##0.00"))
   Reporte.Total11.SetText IIf(tTarjeta(11) = "", "", Format(nTarjeta(11) + nPropina(11) + nRecibo(11), "###,###,##0.00"))
   Reporte.Total12.SetText IIf(tTarjeta(12) = "", "", Format(nTarjeta(12) + nPropina(12) + nRecibo(12), "###,###,##0.00"))
   Reporte.Total13.SetText IIf(tTarjeta(13) = "", "", Format(nTarjeta(13) + nPropina(13) + nRecibo(13), "###,###,##0.00"))
   Reporte.Total14.SetText IIf(tTarjeta(14) = "", "", Format(nTarjeta(14) + nPropina(14) + nRecibo(14), "###,###,##0.00"))
   Reporte.Total15.SetText IIf(tTarjeta(15) = "", "", Format(nTarjeta(15) + nPropina(15) + nRecibo(15), "###,###,##0.00"))
   Reporte.Total16.SetText IIf(tTarjeta(16) = "", "", Format(nTarjeta(16) + nPropina(16) + nRecibo(16), "###,###,##0.00"))
   Reporte.Total17.SetText IIf(tTarjeta(17) = "", "", Format(nTarjeta(17) + nPropina(17) + nRecibo(17), "###,###,##0.00"))
   Reporte.Total18.SetText IIf(tTarjeta(18) = "", "", Format(nTarjeta(18) + nPropina(18) + nRecibo(18), "###,###,##0.00"))
   Reporte.Total19.SetText IIf(tTarjeta(19) = "", "", Format(nTarjeta(19) + nPropina(19) + nRecibo(19), "###,###,##0.00"))
   Reporte.Total20.SetText IIf(tTarjeta(20) = "", "", Format(nTarjeta(20) + nPropina(20) + nRecibo(20), "###,###,##0.00"))
   Reporte.Total21.SetText IIf(tTarjeta(21) = "", "", Format(nTarjeta(21) + nPropina(21) + nRecibo(21), "###,###,##0.00"))
   Reporte.Total22.SetText IIf(tTarjeta(22) = "", "", Format(nTarjeta(22) + nPropina(22) + nRecibo(22), "###,###,##0.00"))
   Reporte.Total23.SetText IIf(tTarjeta(23) = "", "", Format(nTarjeta(23) + nPropina(23) + nRecibo(23), "###,###,##0.00"))
   Reporte.Total24.SetText IIf(tTarjeta(24) = "", "", Format(nTarjeta(24) + nPropina(24) + nRecibo(24), "###,###,##0.00"))

   nRetencionT = 0
   For i = 1 To 24
      nPRetencion(i) = (nTarjeta(i) + nPropina(i) + nRecibo(i)) / 100 * nRetencion(i)
      nRetencionT = nRetencionT + nPRetencion(i)
   Next i
   
   Reporte.Ret1.SetText IIf(tTarjeta(1) = "", "", Format(nPRetencion(1), "###,###,##0.00"))
   Reporte.Ret2.SetText IIf(tTarjeta(2) = "", "", Format(nPRetencion(2), "###,###,##0.00"))
   Reporte.Ret3.SetText IIf(tTarjeta(3) = "", "", Format(nPRetencion(3), "###,###,##0.00"))
   Reporte.Ret4.SetText IIf(tTarjeta(4) = "", "", Format(nPRetencion(4), "###,###,##0.00"))
   Reporte.Ret5.SetText IIf(tTarjeta(5) = "", "", Format(nPRetencion(5), "###,###,##0.00"))
   Reporte.Ret6.SetText IIf(tTarjeta(6) = "", "", Format(nPRetencion(6), "###,###,##0.00"))
   Reporte.Ret7.SetText IIf(tTarjeta(7) = "", "", Format(nPRetencion(7), "###,###,##0.00"))
   Reporte.Ret8.SetText IIf(tTarjeta(8) = "", "", Format(nPRetencion(8), "###,###,##0.00"))
   Reporte.Ret9.SetText IIf(tTarjeta(9) = "", "", Format(nPRetencion(9), "###,###,##0.00"))
   Reporte.Ret10.SetText IIf(tTarjeta(10) = "", "", Format(nPRetencion(10), "###,###,##0.00"))
   Reporte.Ret11.SetText IIf(tTarjeta(11) = "", "", Format(nPRetencion(11), "###,###,##0.00"))
   Reporte.Ret12.SetText IIf(tTarjeta(12) = "", "", Format(nPRetencion(12), "###,###,##0.00"))
   Reporte.Ret13.SetText IIf(tTarjeta(13) = "", "", Format(nPRetencion(13), "###,###,##0.00"))
   Reporte.Ret14.SetText IIf(tTarjeta(14) = "", "", Format(nPRetencion(14), "###,###,##0.00"))
   Reporte.Ret15.SetText IIf(tTarjeta(15) = "", "", Format(nPRetencion(15), "###,###,##0.00"))
   Reporte.Ret16.SetText IIf(tTarjeta(16) = "", "", Format(nPRetencion(16), "###,###,##0.00"))
   Reporte.Ret17.SetText IIf(tTarjeta(17) = "", "", Format(nPRetencion(17), "###,###,##0.00"))
   Reporte.Ret18.SetText IIf(tTarjeta(18) = "", "", Format(nPRetencion(18), "###,###,##0.00"))
   Reporte.Ret19.SetText IIf(tTarjeta(19) = "", "", Format(nPRetencion(19), "###,###,##0.00"))
   Reporte.Ret20.SetText IIf(tTarjeta(20) = "", "", Format(nPRetencion(20), "###,###,##0.00"))
   Reporte.Ret21.SetText IIf(tTarjeta(21) = "", "", Format(nPRetencion(21), "###,###,##0.00"))
   Reporte.Ret22.SetText IIf(tTarjeta(22) = "", "", Format(nPRetencion(22), "###,###,##0.00"))
   Reporte.Ret23.SetText IIf(tTarjeta(23) = "", "", Format(nPRetencion(23), "###,###,##0.00"))
   Reporte.Ret24.SetText IIf(tTarjeta(24) = "", "", Format(nPRetencion(24), "###,###,##0.00"))

   nRetencionProT = 0
   For i = 1 To 24
      nPRetencionPro(i) = (nPropina(i)) / 100 * nRetencion(i)
      nRetencionProT = nRetencionProT + nPRetencionPro(i)
   Next i
   
   Genera22

End Sub

Private Sub Genera22()
   Reporte.RetPro1.SetText IIf(tTarjeta(1) = "", "", Format(nPRetencionPro(1), "###,###,##0.00"))
   Reporte.RetPro2.SetText IIf(tTarjeta(2) = "", "", Format(nPRetencionPro(2), "###,###,##0.00"))
   Reporte.RetPro3.SetText IIf(tTarjeta(3) = "", "", Format(nPRetencionPro(3), "###,###,##0.00"))
   Reporte.RetPro4.SetText IIf(tTarjeta(4) = "", "", Format(nPRetencionPro(4), "###,###,##0.00"))
   Reporte.RetPro5.SetText IIf(tTarjeta(5) = "", "", Format(nPRetencionPro(5), "###,###,##0.00"))
   Reporte.RetPro6.SetText IIf(tTarjeta(6) = "", "", Format(nPRetencionPro(6), "###,###,##0.00"))
   Reporte.RetPro7.SetText IIf(tTarjeta(7) = "", "", Format(nPRetencionPro(7), "###,###,##0.00"))
   Reporte.RetPro8.SetText IIf(tTarjeta(8) = "", "", Format(nPRetencionPro(8), "###,###,##0.00"))
   Reporte.RetPro9.SetText IIf(tTarjeta(9) = "", "", Format(nPRetencionPro(9), "###,###,##0.00"))
   Reporte.RetPro10.SetText IIf(tTarjeta(10) = "", "", Format(nPRetencionPro(10), "###,###,##0.00"))
   Reporte.RetPro11.SetText IIf(tTarjeta(11) = "", "", Format(nPRetencionPro(11), "###,###,##0.00"))
   Reporte.RetPro12.SetText IIf(tTarjeta(12) = "", "", Format(nPRetencionPro(12), "###,###,##0.00"))
   Reporte.RetPro13.SetText IIf(tTarjeta(13) = "", "", Format(nPRetencionPro(13), "###,###,##0.00"))
   Reporte.RetPro14.SetText IIf(tTarjeta(14) = "", "", Format(nPRetencionPro(14), "###,###,##0.00"))
   Reporte.RetPro15.SetText IIf(tTarjeta(15) = "", "", Format(nPRetencionPro(15), "###,###,##0.00"))
   Reporte.RetPro16.SetText IIf(tTarjeta(16) = "", "", Format(nPRetencionPro(16), "###,###,##0.00"))
   Reporte.RetPro17.SetText IIf(tTarjeta(17) = "", "", Format(nPRetencionPro(17), "###,###,##0.00"))
   Reporte.RetPro18.SetText IIf(tTarjeta(18) = "", "", Format(nPRetencionPro(18), "###,###,##0.00"))
   Reporte.RetPro19.SetText IIf(tTarjeta(19) = "", "", Format(nPRetencionPro(19), "###,###,##0.00"))
   Reporte.RetPro20.SetText IIf(tTarjeta(20) = "", "", Format(nPRetencionPro(20), "###,###,##0.00"))
   Reporte.RetPro21.SetText IIf(tTarjeta(21) = "", "", Format(nPRetencionPro(21), "###,###,##0.00"))
   Reporte.RetPro22.SetText IIf(tTarjeta(22) = "", "", Format(nPRetencionPro(22), "###,###,##0.00"))
   Reporte.RetPro23.SetText IIf(tTarjeta(23) = "", "", Format(nPRetencionPro(23), "###,###,##0.00"))
   Reporte.RetPro24.SetText IIf(tTarjeta(24) = "", "", Format(nPRetencionPro(24), "###,###,##0.00"))
   
   Reporte.NetPro1.SetText IIf(tTarjeta(1) = "", "", Format(nPropina(1) - nPRetencionPro(1), "###,###,##0.00"))
   Reporte.NetPro2.SetText IIf(tTarjeta(2) = "", "", Format(nPropina(2) - nPRetencionPro(2), "###,###,##0.00"))
   Reporte.NetPro3.SetText IIf(tTarjeta(3) = "", "", Format(nPropina(3) - nPRetencionPro(3), "###,###,##0.00"))
   Reporte.NetPro4.SetText IIf(tTarjeta(4) = "", "", Format(nPropina(4) - nPRetencionPro(4), "###,###,##0.00"))
   Reporte.NetPro5.SetText IIf(tTarjeta(5) = "", "", Format(nPropina(5) - nPRetencionPro(5), "###,###,##0.00"))
   Reporte.NetPro6.SetText IIf(tTarjeta(6) = "", "", Format(nPropina(6) - nPRetencionPro(6), "###,###,##0.00"))
   Reporte.NetPro7.SetText IIf(tTarjeta(7) = "", "", Format(nPropina(7) - nPRetencionPro(7), "###,###,##0.00"))
   Reporte.NetPro8.SetText IIf(tTarjeta(8) = "", "", Format(nPropina(8) - nPRetencionPro(8), "###,###,##0.00"))
   Reporte.NetPro9.SetText IIf(tTarjeta(9) = "", "", Format(nPropina(9) - nPRetencionPro(9), "###,###,##0.00"))
   Reporte.NetPro10.SetText IIf(tTarjeta(10) = "", "", Format(nPropina(10) - nPRetencionPro(10), "###,###,##0.00"))
   Reporte.NetPro11.SetText IIf(tTarjeta(11) = "", "", Format(nPropina(11) - nPRetencionPro(11), "###,###,##0.00"))
   Reporte.NetPro12.SetText IIf(tTarjeta(12) = "", "", Format(nPropina(12) - nPRetencionPro(12), "###,###,##0.00"))
   Reporte.NetPro13.SetText IIf(tTarjeta(13) = "", "", Format(nPropina(13) - nPRetencionPro(13), "###,###,##0.00"))
   Reporte.NetPro14.SetText IIf(tTarjeta(14) = "", "", Format(nPropina(14) - nPRetencionPro(14), "###,###,##0.00"))
   Reporte.NetPro15.SetText IIf(tTarjeta(15) = "", "", Format(nPropina(15) - nPRetencionPro(15), "###,###,##0.00"))
   Reporte.NetPro16.SetText IIf(tTarjeta(16) = "", "", Format(nPropina(16) - nPRetencionPro(16), "###,###,##0.00"))
   Reporte.NetPro17.SetText IIf(tTarjeta(17) = "", "", Format(nPropina(17) - nPRetencionPro(17), "###,###,##0.00"))
   Reporte.NetPro18.SetText IIf(tTarjeta(18) = "", "", Format(nPropina(18) - nPRetencionPro(18), "###,###,##0.00"))
   Reporte.NetPro19.SetText IIf(tTarjeta(19) = "", "", Format(nPropina(19) - nPRetencionPro(19), "###,###,##0.00"))
   Reporte.NetPro20.SetText IIf(tTarjeta(20) = "", "", Format(nPropina(20) - nPRetencionPro(20), "###,###,##0.00"))
   Reporte.NetPro21.SetText IIf(tTarjeta(21) = "", "", Format(nPropina(21) - nPRetencionPro(21), "###,###,##0.00"))
   Reporte.NetPro22.SetText IIf(tTarjeta(22) = "", "", Format(nPropina(22) - nPRetencionPro(22), "###,###,##0.00"))
   Reporte.NetPro23.SetText IIf(tTarjeta(23) = "", "", Format(nPropina(23) - nPRetencionPro(23), "###,###,##0.00"))
   Reporte.NetPro24.SetText IIf(tTarjeta(24) = "", "", Format(nPropina(24) - nPRetencionPro(24), "###,###,##0.00"))
   
   Reporte.TOTTAR1.SetText Format(nTarjetaT, "###,###,##0.00")
   Reporte.TOTPRO.SetText Format(nPropinaT, "###,###,##0.00")
   Reporte.TOTREC.SetText Format(nReciboT, "###,###,##0.00")
   Reporte.TOTRET.SetText Format(nRetencionT, "###,###,##0.00")
   Reporte.TotalDoc.SetText Format(nTarjetaT + nPropinaT + nReciboT, "###,###,##0.00")
   Reporte.TOTRETPRO.SetText Format(nRetencionProT, "###,###,##0.00")
   Reporte.TOTNETPRO.SetText Format(nPropinaT - nRetencionProT, "###,###,##0.00")
   
   Reporte.OTRO1.SetText tOtro(1)
   Reporte.Otro2.SetText tOtro(2)
   Reporte.Otro3.SetText tOtro(3)
   Reporte.Otro4.SetText tOtro(4)
   Reporte.Otro5.SetText tOtro(5)
   Reporte.Otro6.SetText tOtro(6)
   Reporte.Otro7.SetText tOtro(7)
   Reporte.Otro8.SetText tOtro(8)
   Reporte.Otro9.SetText tOtro(9)
   Reporte.Otro10.SetText tOtro(10)
   Reporte.OTRO11.SetText tOtro(11)
   Reporte.Otro12.SetText tOtro(12)
   Reporte.Otro13.SetText tOtro(13)
   Reporte.Otro14.SetText tOtro(14)
   Reporte.Otro15.SetText tOtro(15)
   Reporte.Otro16.SetText tOtro(16)
   Reporte.Otro17.SetText tOtro(17)
   Reporte.Otro18.SetText tOtro(18)
   Reporte.Otro19.SetText tOtro(19)
   Reporte.Otro20.SetText tOtro(20)
       
   Reporte.MONTOO1.SetText IIf(tOtro(1) = "", "", Format(nOtro(1), "####,##0.00"))
   Reporte.MONTOO2.SetText IIf(tOtro(2) = "", "", Format(nOtro(2), "####,##0.00"))
   Reporte.MONTOO3.SetText IIf(tOtro(3) = "", "", Format(nOtro(3), "####,##0.00"))
   Reporte.MONTOO4.SetText IIf(tOtro(4) = "", "", Format(nOtro(4), "####,##0.00"))
   Reporte.MONTOO5.SetText IIf(tOtro(5) = "", "", Format(nOtro(5), "####,##0.00"))
   Reporte.MONTOO6.SetText IIf(tOtro(6) = "", "", Format(nOtro(6), "####,##0.00"))
   Reporte.MONTOO7.SetText IIf(tOtro(7) = "", "", Format(nOtro(7), "####,##0.00"))
   Reporte.MONTOO8.SetText IIf(tOtro(8) = "", "", Format(nOtro(8), "####,##0.00"))
   Reporte.MONTOO9.SetText IIf(tOtro(9) = "", "", Format(nOtro(9), "####,##0.00"))
   Reporte.MONTOO10.SetText IIf(tOtro(10) = "", "", Format(nOtro(10), "####,##0.00"))
   Reporte.MONTOO11.SetText IIf(tOtro(11) = "", "", Format(nOtro(11), "####,##0.00"))
   Reporte.MONTOO12.SetText IIf(tOtro(12) = "", "", Format(nOtro(12), "####,##0.00"))
   Reporte.MONTOO13.SetText IIf(tOtro(13) = "", "", Format(nOtro(13), "####,##0.00"))
   Reporte.MONTOO14.SetText IIf(tOtro(14) = "", "", Format(nOtro(14), "####,##0.00"))
   Reporte.MONTOO15.SetText IIf(tOtro(15) = "", "", Format(nOtro(15), "####,##0.00"))
   Reporte.MONTOO16.SetText IIf(tOtro(16) = "", "", Format(nOtro(16), "####,##0.00"))
   Reporte.MONTOO17.SetText IIf(tOtro(17) = "", "", Format(nOtro(17), "####,##0.00"))
   Reporte.MONTOO18.SetText IIf(tOtro(18) = "", "", Format(nOtro(18), "####,##0.00"))
   Reporte.MONTOO19.SetText IIf(tOtro(19) = "", "", Format(nOtro(19), "####,##0.00"))
   Reporte.MONTOO20.SetText IIf(tOtro(20) = "", "", Format(nOtro(20), "####,##0.00"))
   Reporte.MONTOOE1.SetText IIf(tOtro(1) = "", "", Format(xOtro(1), "####,##0.00"))
   Reporte.MONTOOE2.SetText IIf(tOtro(2) = "", "", Format(xOtro(2), "####,##0.00"))
   Reporte.MONTOOE3.SetText IIf(tOtro(3) = "", "", Format(xOtro(3), "####,##0.00"))
   Reporte.MONTOOE4.SetText IIf(tOtro(4) = "", "", Format(xOtro(4), "####,##0.00"))
   Reporte.MONTOOE5.SetText IIf(tOtro(5) = "", "", Format(xOtro(5), "####,##0.00"))
   Reporte.MONTOOE6.SetText IIf(tOtro(6) = "", "", Format(xOtro(6), "####,##0.00"))
   Reporte.MONTOOE7.SetText IIf(tOtro(7) = "", "", Format(xOtro(7), "####,##0.00"))
   Reporte.MONTOOE8.SetText IIf(tOtro(8) = "", "", Format(xOtro(8), "####,##0.00"))
   Reporte.MONTOOE9.SetText IIf(tOtro(9) = "", "", Format(xOtro(9), "####,##0.00"))
   Reporte.MONTOOE10.SetText IIf(tOtro(10) = "", "", Format(xOtro(10), "####,##0.00"))
   Reporte.MONTOOE11.SetText IIf(tOtro(11) = "", "", Format(xOtro(11), "####,##0.00"))
   Reporte.MONTOOE12.SetText IIf(tOtro(12) = "", "", Format(xOtro(12), "####,##0.00"))
   Reporte.MONTOOE13.SetText IIf(tOtro(13) = "", "", Format(xOtro(13), "####,##0.00"))
   Reporte.MONTOOE14.SetText IIf(tOtro(14) = "", "", Format(xOtro(14), "####,##0.00"))
   Reporte.MONTOOE15.SetText IIf(tOtro(15) = "", "", Format(xOtro(15), "####,##0.00"))
   Reporte.MONTOOE16.SetText IIf(tOtro(16) = "", "", Format(xOtro(16), "####,##0.00"))
   Reporte.MONTOOE17.SetText IIf(tOtro(17) = "", "", Format(xOtro(17), "####,##0.00"))
   Reporte.MONTOOE18.SetText IIf(tOtro(18) = "", "", Format(xOtro(18), "####,##0.00"))
   Reporte.MONTOOE19.SetText IIf(tOtro(19) = "", "", Format(xOtro(19), "####,##0.00"))
   Reporte.MONTOOE20.SetText IIf(tOtro(20) = "", "", Format(xOtro(20), "####,##0.00"))
      
   Reporte.OVuelto1.SetText IIf(tOtro(1) = "", "", Format(nDiferencia(1), "####,##0.00"))
   Reporte.OVuelto2.SetText IIf(tOtro(2) = "", "", Format(nDiferencia(2), "####,##0.00"))
   Reporte.OVuelto3.SetText IIf(tOtro(3) = "", "", Format(nDiferencia(3), "####,##0.00"))
   Reporte.OVuelto4.SetText IIf(tOtro(4) = "", "", Format(nDiferencia(4), "####,##0.00"))
   Reporte.OVuelto5.SetText IIf(tOtro(5) = "", "", Format(nDiferencia(5), "####,##0.00"))
   Reporte.OVuelto6.SetText IIf(tOtro(6) = "", "", Format(nDiferencia(6), "####,##0.00"))
   Reporte.OVuelto7.SetText IIf(tOtro(7) = "", "", Format(nDiferencia(7), "####,##0.00"))
   Reporte.OVuelto8.SetText IIf(tOtro(8) = "", "", Format(nDiferencia(8), "####,##0.00"))
   Reporte.OVuelto9.SetText IIf(tOtro(9) = "", "", Format(nDiferencia(9), "####,##0.00"))
   Reporte.OVuelto10.SetText IIf(tOtro(10) = "", "", Format(nDiferencia(10), "####,##0.00"))
   Reporte.OVuelto11.SetText IIf(tOtro(11) = "", "", Format(nDiferencia(11), "####,##0.00"))
   Reporte.OVuelto12.SetText IIf(tOtro(12) = "", "", Format(nDiferencia(12), "####,##0.00"))
   Reporte.OVuelto13.SetText IIf(tOtro(13) = "", "", Format(nDiferencia(13), "####,##0.00"))
   Reporte.OVuelto14.SetText IIf(tOtro(14) = "", "", Format(nDiferencia(14), "####,##0.00"))
   Reporte.OVuelto15.SetText IIf(tOtro(15) = "", "", Format(nDiferencia(15), "####,##0.00"))
   Reporte.OVuelto16.SetText IIf(tOtro(16) = "", "", Format(nDiferencia(16), "####,##0.00"))
   Reporte.OVuelto17.SetText IIf(tOtro(17) = "", "", Format(nDiferencia(17), "####,##0.00"))
   Reporte.OVuelto18.SetText IIf(tOtro(18) = "", "", Format(nDiferencia(18), "####,##0.00"))
   Reporte.OVuelto19.SetText IIf(tOtro(19) = "", "", Format(nDiferencia(19), "####,##0.00"))
   Reporte.OVuelto20.SetText IIf(tOtro(20) = "", "", Format(nDiferencia(20), "####,##0.00"))
   
   Reporte.OTROTOTAL.SetText Format(nOtro(1) + nOtro(2) + nOtro(3) + nOtro(4) + nOtro(5) + nOtro(6) + nOtro(7) + nOtro(8) + nOtro(9) + nOtro(10) + nOtro(11) + nOtro(12) + nOtro(13) + nOtro(14) + nOtro(15) + nOtro(16) + nOtro(17) + nOtro(18) + nOtro(19) + nOtro(20), "####,##0.00")
   Reporte.OTROTOTALE.SetText Format(xOtro(1) + xOtro(2) + xOtro(3) + xOtro(4) + xOtro(5) + xOtro(6) + xOtro(7) + xOtro(8) + xOtro(9) + xOtro(10) + xOtro(11) + xOtro(12) + xOtro(13) + xOtro(14) + xOtro(15) + xOtro(16) + xOtro(17) + xOtro(18) + xOtro(19) + xOtro(20), "####,##0.00")
   Reporte.VUELTOTOTAL.SetText Format(nDiferencia(1) + nDiferencia(2) + nDiferencia(3) + nDiferencia(4) + nDiferencia(5) + nDiferencia(6) + nDiferencia(7) + nDiferencia(8) + nDiferencia(9) + nDiferencia(10) + nDiferencia(11) + nDiferencia(12) + nDiferencia(13) + nDiferencia(14) + nDiferencia(15) + nDiferencia(16) + nDiferencia(17) + nDiferencia(18) + nDiferencia(19) + nDiferencia(20), "####,##0.00")
   
   Reporte.DiscardSavedData
    If chkResumen.Value Then
        Reporte.Section3.Suppress = True: Reporte.Section6.Suppress = True
        Reporte.Section7.Suppress = True: Reporte.Section8.Suppress = True: Reporte.Section9.Suppress = True
    Else
        Reporte.Section3.Suppress = False: Reporte.Section6.Suppress = False: Reporte.Section7.Suppress = False
        Reporte.Section8.Suppress = False: Reporte.Section9.Suppress = False
    End If
   
   Reporte.Database.SetDataSource rsReporte
   Reporte.ReportTitle = sTitulo
   Reporte.Final.SetText "Caja : " & IIf(ChkTurno.Value, "Todas las Cajas", sCaja)
   Reporte.Text15.SetText sRazonSocial
End Sub


Private Sub chkDiaContable_Click()
    If chkDiaContable.Value = 1 Then
        Me.dtpHorFin.Visible = False
        Me.dtpHorIni.Visible = False
 
    Else
        Me.dtpHorFin.Visible = True
        Me.dtpHorIni.Visible = True
         
    End If
End Sub
Public Sub Genera2()
   Dim oRsSumas As Recordset
    Dim oComando As clsComando
    Set oComando = New clsComando
    Set ReporteV = New dsrLiquidacionVenta
    
    If lNcOfisis Then
        If Not oComando.CreateCmdSp("spRep_LiquidacionSuma_NC", Cn) Then
            Set oComando = Nothing
            Exit Sub
        End If
    Else
         If Not oComando.CreateCmdSp("spRep_LiquidacionSuma", Cn) Then
            Set oComando = Nothing
            Exit Sub
        End If
    End If
        
    oComando.CreateParameter "@flagTurno", adBoolean, adParamInput, 1, ChkTurno.Value
    oComando.CreateParameter "@flagDiaContable", adBoolean, adParamInput, 1, chkDiaContable.Value
    oComando.CreateParameter "@sturno", adVarChar, adParamInput, 30, sTurno
    oComando.CreateParameter "@sUsuario", adVarChar, adParamInput, 30, Me.cboUsuario.Text
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@sSectorVenta", adVarChar, adParamInput, 50, Me.cboSectorVenta.BoundText
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set oRsSumas = oComando.GetSP()
    Dim xparametros As String
    Dim ktTurno As String
    Dim ktUsuario As String
    Dim kfRegistroi As String
    Dim kfRegistrof As String
    If ChkTurno.Value = 0 Then
        xparametros = " '" & sTurno & "', NULL, NULL, "
        ktTurno = " '" & sTurno & "'"
        kfRegistroi = "NULL"
        kfRegistrof = "NULL"
    Else
        xparametros = " NULL, '" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "', '" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "', "
        ktTurno = " NULL"
        
        If chkDiaContable.Value = 0 Then
                kfRegistroi = "'" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "'"
                kfRegistrof = "'" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "'"
        Else
                kfRegistroi = "'" & Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format("00:00:00", "HH:nn:ss") & "'"
                kfRegistrof = "'" & Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format("23:59:59", "HH:nn:ss") & "'"
        End If
    End If
    
    If chkUsuario.Value = 0 Then
        xparametros = xparametros + " '" + cboUsuario.Text + "'"
        ktUsuario = "'" + cboUsuario.Text + "'"
    Else
        xparametros = xparametros + " NULL"
        ktUsuario = " NULL"
    End If
    
    xparametros = xparametros + ",'" + IIf(chkDiaContable.Value, "1", "0") + "', '" + cboSectorVenta.BoundText + "'," + IIf(chkCortesia.Value, "1", "0") + ""
    If lNcOfisis Then
        Set rsReporte = Lib.OpenRecordset("usp_InforestCon_ObtenerReporteLiquidacionVentas_NC" + xparametros, Cn)
        If rsReporte.EOF = True Then
           MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    Else
        Set rsReporte = Lib.OpenRecordset("usp_InforestCon_ObtenerReporteLiquidacionVentas" + xparametros, Cn)
        If rsReporte.EOF = True Then
           MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
    End If
    ReporteV.Text13.SetText localConectado
    ReporteV.DiscardSavedData
    ReporteV.ReportTitle = sTitulo
    ReporteV.Database.SetDataSource rsReporte
    ReporteV.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    ReporteV.PaperOrientation = crPortrait
                 
    Isql = "select Count(*) as Cantidad from vPedidoCorrelativo " & _
            "where Mesa IS NOT NULL And tTipoPedido = '01' " & _
            "And (tTurno = " & ktTurno & " OR " & ktTurno & " IS NULL) " & _
            "And (tUsuario = " & ktUsuario & " OR " & ktUsuario & " IS NULL) " & _
            "And ((fFecha BETWEEN " & kfRegistroi & " AND " & kfRegistrof & ") OR (" & kfRegistroi & " IS NULL AND " & kfRegistrof & " IS NULL))"
    ReporteV.nMesas.SetText "(" & Lib.OpenRecordset(Isql, Cn)!Cantidad & ")"
                               
    Dim fNumerador As Double
    Dim gDenominador As Double
    gDenominador = 0
    fNumerador = 0
    Dim TotalCantidad As Double
    Dim totalPN As Double
    Dim TotalI1 As Double
    Dim TotalI2 As Double
    Dim TotalI3 As Double
    Dim TotalPV As Double
    Dim TotalFactor As Double
 
    If Me.chkConGrafico.Value Then
        ReporteV.ParameterFields.Item(1).SetCurrentValue (2)
    Else
        ReporteV.ParameterFields.Item(1).SetCurrentValue (0)
    End If
        
    oRsSumas.Filter = "ttipopedido='01'"
    '--En El Local
   
    If Not oRsSumas.EOF Then
        fNumerador = Format(IIf(IsNull(oRsSumas!nTotalPromedio), 0, oRsSumas!nTotalPromedio), "###,##0")
        TotalCantidad = TotalCantidad + oRsSumas!nTotalPromedio
        ReporteV.nLOCAL.SetText "(" + CStr(fNumerador) + ")"
    End If
    If Not oRsSumas.EOF Then
        If optVenta.Value = False Then
            gDenominador = Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        Else
            gDenominador = Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        End If
        
        ReporteV.Text2.SetText "Total " & (Calcular("select tdetallado codigo from TCANALVENTA where tCodigoCanalVenta='01'", Cn))
        
        ReporteV.VVLocalPN.SetText Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        ReporteV.VVLocalI1.SetText Format(IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1), "###,##0.00")
        ReporteV.VVLocalI2.SetText Format(IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2), "###,##0.00")
        ReporteV.VVLocalI3.SetText Format(IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3), "###,##0.00")
        ReporteV.VVLocalPV.SetText Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        totalPN = totalPN + IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto)
        TotalI1 = TotalI1 + IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1)
        TotalI2 = TotalI2 + IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2)
        TotalI3 = TotalI3 + IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3)
        TotalPV = TotalPV + IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta)
        If fNumerador > 0 Then
            ReporteV.FactorLocal.SetText Format(CStr(gDenominador / fNumerador), "###,##0.00")
        End If
    End If
    
    gDenominador = 0
    fNumerador = 0
    '--Cantidades Delivery
    oRsSumas.Filter = "ttipopedido='02'"
    If Not oRsSumas.EOF Then
        fNumerador = Format(IIf(IsNull(oRsSumas!nTotalPromedio), 0, oRsSumas!nTotalPromedio), "###,##0")
        TotalCantidad = TotalCantidad + oRsSumas!nTotalPromedio
        ReporteV.nDELIVERY.SetText "(" + CStr(fNumerador) + ")"
    End If
    If Not oRsSumas.EOF Then
        If optVenta.Value = False Then
            gDenominador = Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        Else
            gDenominador = Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        End If
        
        ReporteV.Text3.SetText "Total " & (Calcular("select tdetallado codigo from TCANALVENTA where tCodigoCanalVenta='02'", Cn))
        
        ReporteV.VVDeliveryPN.SetText Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        ReporteV.VVDeliveryI1.SetText Format(IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1), "###,##0.00")
        ReporteV.VVDeliveryI2.SetText Format(IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2), "###,##0.00")
        ReporteV.VVDeliveryI3.SetText Format(IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3), "###,##0.00")
        ReporteV.VVDeliveryPV.SetText Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        totalPN = totalPN + IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto)
        TotalI1 = TotalI1 + IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1)
        TotalI2 = TotalI2 + IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2)
        TotalI3 = TotalI3 + IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3)
        TotalPV = TotalPV + IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta)
        If fNumerador > 0 Then
            ReporteV.FactorDelivery.SetText Format(CStr(gDenominador / fNumerador), "###,##0.00")
        End If
    End If
    gDenominador = 0
    fNumerador = 0
    
    '--Para Llevar
    oRsSumas.Filter = "ttipopedido='03'"
    
    If Not oRsSumas.EOF Then
        fNumerador = Format(IIf(IsNull(oRsSumas!nTotalPromedio), 0, oRsSumas!nTotalPromedio), "###,##0")
        TotalCantidad = TotalCantidad + oRsSumas!nTotalPromedio
        ReporteV.nLlevar.SetText "(" + CStr(fNumerador) + ")"
    End If
    If Not oRsSumas.EOF Then
        If optVenta.Value = False Then
                gDenominador = Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        Else
                gDenominador = Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        End If
        
         ReporteV.Text5.SetText "Total " & (Calcular("select tdetallado codigo from TCANALVENTA where tCodigoCanalVenta='03'", Cn))
        
        ReporteV.VVLlevarPN.SetText Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        ReporteV.VVLlevarI1.SetText Format(IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1), "###,##0.00")
        ReporteV.VVLlevarI2.SetText Format(IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2), "###,##0.00")
        ReporteV.VVLlevarI3.SetText Format(IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3), "###,##0.00")
        ReporteV.VVLlevarPV.SetText Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        totalPN = totalPN + IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto)
        TotalI1 = TotalI1 + IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1)
        TotalI2 = TotalI2 + IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2)
        TotalI3 = TotalI3 + IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3)
        TotalPV = TotalPV + IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta)
        If fNumerador > 0 Then
            ReporteV.FactorLlevar.SetText Format(CStr(gDenominador / fNumerador), "###,##0.00")
        End If
    End If
    gDenominador = 0
    fNumerador = 0
    
    '--Canal4
    oRsSumas.Filter = "ttipopedido='04'"
    ReporteV.Text6.SetText "Total " & (Calcular("select tdetallado codigo from TCANALVENTA where tCodigoCanalVenta='04'", Cn))
    If Not oRsSumas.EOF Then
        fNumerador = Format(IIf(IsNull(oRsSumas!nTotalPromedio), 0, oRsSumas!nTotalPromedio), "###,##0")
        TotalCantidad = TotalCantidad + oRsSumas!nTotalPromedio
        ReporteV.nCanal4.SetText "(" + CStr(fNumerador) + ")"
    End If
    If Not oRsSumas.EOF Then
        If optVenta.Value = False Then
            gDenominador = Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        Else
            gDenominador = Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        End If
        
        
          
        ReporteV.VVCanal4PN.SetText Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        ReporteV.VVCanal4I1.SetText Format(IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1), "###,##0.00")
        ReporteV.VVCanal4I2.SetText Format(IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2), "###,##0.00")
        ReporteV.VVCanal4I3.SetText Format(IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3), "###,##0.00")
        ReporteV.VVCanal4PV.SetText Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        totalPN = totalPN + IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto)
        TotalI1 = TotalI1 + IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1)
        TotalI2 = TotalI2 + IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2)
        TotalI3 = TotalI3 + IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3)
        TotalPV = TotalPV + IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta)
        If fNumerador > 0 Then
            ReporteV.FactorCanal4.SetText Format(CStr(gDenominador / fNumerador), "###,##0.00")
        End If
    End If
    gDenominador = 0
    fNumerador = 0
    
    '--canal5
    oRsSumas.Filter = "ttipopedido='05'"
    ReporteV.Text7.SetText "Total " & (Calcular("select tdetallado codigo from TCANALVENTA where tCodigoCanalVenta='05'", Cn))
    If Not oRsSumas.EOF Then
        fNumerador = Format(IIf(IsNull(oRsSumas!nTotalPromedio), 0, oRsSumas!nTotalPromedio), "###,##0")
        TotalCantidad = TotalCantidad + oRsSumas!nTotalPromedio
        
        ReporteV.nCanal5.SetText "(" + CStr(fNumerador) + ")"
    End If
    
    If Not oRsSumas.EOF Then
        If optVenta.Value = False Then
            gDenominador = Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        Else
            gDenominador = Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        End If
        
        
        
        ReporteV.VVCanal5PN.SetText Format(IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto), "###,##0.00")
        ReporteV.VVCanal5I1.SetText Format(IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1), "###,##0.00")
        ReporteV.VVCanal5I2.SetText Format(IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2), "###,##0.00")
        ReporteV.VVCanal5I3.SetText Format(IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3), "###,##0.00")
        ReporteV.VVCanal5PV.SetText Format(IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta), "###,##0.00")
        totalPN = totalPN + IIf(IsNull(oRsSumas!nNeto), 0, oRsSumas!nNeto)
        TotalI1 = TotalI1 + IIf(IsNull(oRsSumas!nImpuesto1), 0, oRsSumas!nImpuesto1)
        TotalI2 = TotalI2 + IIf(IsNull(oRsSumas!nImpuesto2), 0, oRsSumas!nImpuesto2)
        TotalI3 = TotalI3 + IIf(IsNull(oRsSumas!nImpuesto3), 0, oRsSumas!nImpuesto3)
        TotalPV = TotalPV + IIf(IsNull(oRsSumas!nVenta), 0, oRsSumas!nVenta)
        If fNumerador > 0 Then
            ReporteV.FactorCanal5.SetText Format(CStr(gDenominador / fNumerador), "###,##0.00")
        End If
    End If
    gDenominador = 0
    fNumerador = 0
    ReporteV.TotalCantidad.SetText "(" & TotalCantidad & ")"
    ReporteV.TotalVV.SetText Format(totalPN, "###,##0.00")
    ReporteV.TotalI1.SetText Format(TotalI1, "###,##0.00")
    ReporteV.TotalI2.SetText Format(TotalI2, "###,##0.00")
    ReporteV.TotalI3.SetText Format(TotalI3, "###,##0.00")
    ReporteV.TotalPV.SetText Format(TotalPV, "###,##0.00")
    If TotalCantidad = 0 Then
        ReporteV.TotalFactor.SetText Format(0, "###,##0.00")
    Else
        If optVenta.Value = False Then
            ReporteV.TotalFactor.SetText Format((totalPN / TotalCantidad), "###,##0.00")
        Else
            ReporteV.TotalFactor.SetText Format((TotalPV / TotalCantidad), "###,##0.00")
        End If
    End If
    
    '--Cantidades Boletas y Sumas
    Dim xvvSuma As Double
    'Isql = "Select Count(*) cantidad, Sum(nVenta) suma, Sum(nVenta)/Count(*) Promedio From MDocumento Where (tTurno = " + ktTurno + " OR " + ktTurno + " IS NULL) And (tUsuario = " + ktUsuario + " OR " + ktUsuario + " IS NULL) And ((fRegistro BETWEEN " + kfRegistroi + " AND " + kfRegistrof + ") OR (" + kfRegistroi + " IS NULL AND " + kfRegistrof + " IS NULL)) And tTipoDocumento = '02' And tEstadoDocumento = '02'"
   ' Isql = "Select tipDoc.tDetallado, Count(*) cantidad, Sum(mDoc.nVenta) suma, Sum(mDoc.nVenta)/Count(*) Promedio From MDocumento As mDoc Inner Join TTabla as tipDoc On mDoc.tTipoDocumento = tipDoc.tCodigo And tTabla = 'TIPODOCUMENTO' Where (tTurno = " + ktTurno + " OR " + ktTurno + " IS NULL) And (tUsuario = " + ktUsuario + " OR " + ktUsuario + " IS NULL) And ((fRegistro BETWEEN " + kfRegistroi + " AND " + kfRegistrof + ") OR (" + kfRegistroi + " IS NULL AND " + kfRegistrof + " IS NULL)) And mDoc.tEstadoDocumento = '02' And mDoc.tTipoDocumento <> '00' Group by tipDoc.tDetallado"
    If lNcOfisis Then
        Isql = "Select tipDoc.tDescripcion tdetallado, Count(*) cantidad,Sum(mDoc.nVenta) suma, Sum(mDoc.nVenta)/Count(*) Promedio From MDocumento As mDoc Inner Join TTIPODOCUMENTO as tipDoc On mDoc.tTipoDocumento = tipDoc.tCodigoTipoDocumento  Where (tTurno =  " + ktTurno + " OR  " + ktTurno + " IS NULL) And  (tUsuario =  " + ktUsuario + " OR  " + ktUsuario + " IS NULL) And  ((fRegistro BETWEEN " + kfRegistroi + " AND " + kfRegistrof + ") OR (" + kfRegistroi + " IS NULL AND " + kfRegistrof + " IS NULL)) And  mDoc.tEstadoDocumento = '02' And mDoc.tTipoDocumento <> '00'  Group by tipDoc.tDescripcion" & _
                " union Select tipDoc.tDescripcion tdetallado, Count(*) cantidad,(Sum(mDoc.nVenta)*(-1)) AS  suma, Sum(mDoc.nVenta)/Count(*) Promedio From mnotacredito As mDoc Inner Join TTIPODOCUMENTO as tipDoc On mDoc.tTipoDocumento = tipDoc.tCodigoTipoDocumento  Where (tTurno =  " + ktTurno + " OR  " + ktTurno + " IS NULL) And  (tUsuario =  " + ktUsuario + " OR  " + ktUsuario + " IS NULL) And  ((fRegistro BETWEEN " + kfRegistroi + " AND " + kfRegistrof + ") OR (" + kfRegistroi + " IS NULL AND " + kfRegistrof + " IS NULL)) And  mDoc.tEstadoDocumento = '05' And mDoc.tTipoDocumento <> '00'  Group by tipDoc.tDescripcion"
    Else
        Isql = "Select tipDoc.tDescripcion tdetallado, Count(*) cantidad,Sum(mDoc.nVenta) suma, Sum(mDoc.nVenta)/Count(*) Promedio From MDocumento As mDoc Inner Join TTIPODOCUMENTO as tipDoc On mDoc.tTipoDocumento = tipDoc.tCodigoTipoDocumento  Where (tTurno =  " + ktTurno + " OR  " + ktTurno + " IS NULL) And  (tUsuario =  " + ktUsuario + " OR  " + ktUsuario + " IS NULL) And  ((fRegistro BETWEEN " + kfRegistroi + " AND " + kfRegistrof + ") OR (" + kfRegistroi + " IS NULL AND " + kfRegistrof + " IS NULL)) And  mDoc.tEstadoDocumento = '02' And mDoc.tTipoDocumento <> '00'  Group by tipDoc.tDescripcion"
    End If

   ' Debug.Print Isql
    Set oRsSumas = Lib.OpenRecordset(Isql, Cn)
    
    If Not oRsSumas.EOF Then
        Dim otDetallado As String
        Dim ocantidad As String
        Dim osuma As String
        Dim oPromedioi As String
        Dim oPromedio As String
        Dim oi As Integer
        For oi = 1 To oRsSumas.RecordCount
            otDetallado = otDetallado + "Total " & UCase(oRsSumas!tDetallado) + vbCrLf    'Chr(13)
            ocantidad = ocantidad + "(" + Format(IIf(IsNull(oRsSumas!Cantidad), 0, oRsSumas!Cantidad), "###,##0") + ")" + vbCrLf      'Chr(13) & Chr(10)
            osuma = osuma + Format(IIf(IsNull(oRsSumas!Suma), 0, oRsSumas!Suma), "###,##0.00") + vbCrLf      'Chr(13)
            oPromedioi = oPromedioi + "Prom." + vbCrLf      'Chr(13)
            xvvSuma = xvvSuma + oRsSumas!Suma
            oPromedio = oPromedio + Format(IIf(IsNull(oRsSumas!Promedio), 0, oRsSumas!Promedio), "###,##0.00") + vbCrLf   'Chr(13)
            oRsSumas.MoveNext
        Next oi
    
        ReporteV.txtDetallado.SetText otDetallado
        ReporteV.nBOLETAS.SetText ocantidad '"(" + Format(IIf(IsNull(oRsSumas!Cantidad), 0, oRsSumas!Cantidad), "###,##0") + ")"
        ReporteV.VVBoletas.SetText osuma 'Format(IIf(IsNull(oRsSumas!suma), 0, oRsSumas!suma), "###,##0.00")
        ReporteV.txtProm.SetText oPromedioi
        ReporteV.FactorBoletas.SetText oPromedio 'Format(IIf(IsNull(oRsSumas!Promedio), 0, oRsSumas!Promedio), "###,##0.00")
    End If
    '--Cantidades Adultos
    Isql = "select sum(nAdulto) as Adulto, sum(nNino) as Nino from MPEDIDO where (tTurno = " + ktTurno + " OR " + ktTurno + " IS NULL) And (tUsuario = " + ktUsuario + " OR " + ktUsuario + " IS NULL) And ((fRegistro BETWEEN " + kfRegistroi + " AND " + kfRegistrof + ") OR (" + kfRegistroi + " IS NULL AND " + kfRegistrof + " IS NULL)) and (MPEDIDO.tEstadoPedido <> '03' AND MPEDIDO.tEstadoPedido <> '01')"
    Set oRsSumas = Lib.OpenRecordset(Isql, Cn)
    If Not oRsSumas.EOF Then
        ReporteV.nAdultos.SetText "(" + Format(IIf(IsNull(oRsSumas!Adulto), 0, oRsSumas!Adulto), "###,##0") + ")"
        ReporteV.VVSuma.SetText Format(IIf(IsNull(xvvSuma), 0, xvvSuma), "###,##0.00")
        If oRsSumas!Adulto > 0 Then
            ReporteV.FactorSuma.SetText Format(IIf(IsNull(xvvSuma / oRsSumas!Adulto), 0, xvvSuma / oRsSumas!Adulto), "###,##0.00")
        End If
    End If
    '---------------
End Sub

Private Sub Option1_Click()
    Me.chkConGrafico.Enabled = False
End Sub

Private Sub Option2_Click()
    Me.chkConGrafico.Enabled = True
End Sub
Private Sub LLenarOrigenventa()
 'Totales por Origen Venta
    Dim oComando As clsComando
    Set oComando = New clsComando
   If lNcOfisis Then
        If Not oComando.CreateCmdSp("spRep_LiquidacionOrigenVenta", Cn) Then
            Set oComando = Nothing
            Exit Sub
        End If
   Else
        If Not oComando.CreateCmdSp("spRep_LiquidacionOrigenVenta", Cn) Then
            Set oComando = Nothing
             Exit Sub
         End If
   End If
    
    Set RsSumas = Nothing
    oComando.CreateParameter "@flagTurno", adBoolean, adParamInput, 1, ChkTurno.Value
    oComando.CreateParameter "@flagDiaContable", adBoolean, adParamInput, 1, chkDiaContable.Value
    oComando.CreateParameter "@sturno", adVarChar, adParamInput, 20, sTurno
    oComando.CreateParameter "@sUsuario", adVarChar, adParamInput, 20, Me.cboUsuario.Text
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@sSectorVenta", adVarChar, adParamInput, 50, Me.cboSectorVenta.BoundText
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsSumas = oComando.GetSP()
    If RsSumas.RecordCount > 0 Then
        RsSumas.MoveFirst
        For i = 1 To RsSumas.RecordCount
            If RsSumas.RecordCount < 11 Then
                If i = 1 Then
                    Reporte.txtcanal1.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal1.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal1.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal1.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 2 Then
                    Reporte.txtcanal2.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal2.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal2.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal2.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 3 Then
                    Reporte.txtcanal3.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal3.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal3.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal3.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 4 Then
                    Reporte.txtcanal4.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal4.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal4.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal4.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 5 Then
                    Reporte.txtcanal5.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal5.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal5.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal5.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 6 Then
                    Reporte.txtcanal6.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal6.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal6.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal6.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 7 Then
                    Reporte.txtcanal7.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal7.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal7.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal7.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 8 Then
                    Reporte.txtcanal8.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal8.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal8.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal8.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 9 Then
                    Reporte.txtcanal9.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal9.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal9.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal9.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                ElseIf i = 10 Then
                    Reporte.txtcanal10.SetText IIf(IsNull(RsSumas!OrigenVenta), "", RsSumas!OrigenVenta)
                    Reporte.txtsolescanal10.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMN, "#,####,##0.00"))
                    Reporte.txtdolarescanal10.SetText IIf(IsNull(RsSumas!PagoMN), "", Format(RsSumas!PagoMe, "#,####,##0.00"))
                    Reporte.txttotalcanal10.SetText IIf(IsNull(RsSumas!TOTAL), "", Format(RsSumas!TOTAL, "#,####,##0.00"))
                End If
            End If
            RsSumas.MoveNext
        Next i
    End If
End Sub


