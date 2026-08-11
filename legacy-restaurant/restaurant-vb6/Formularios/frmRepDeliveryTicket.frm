VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepDeliveryTicket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Cajeros Delivery"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   ForeColor       =   &H8000000C&
   Icon            =   "frmRepDeliveryTicket.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
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
      Height          =   3180
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   8445
      Begin VB.CheckBox chkCaja 
         Caption         =   "Todos las Cajas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5625
         TabIndex        =   16
         Top             =   420
         Value           =   1  'Checked
         Width           =   2100
      End
      Begin VB.TextBox txtTurno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1386
         Width           =   2805
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
         Height          =   420
         Left            =   4710
         Picture         =   "frmRepDeliveryTicket.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1386
         Width           =   765
      End
      Begin VB.CheckBox chkMotorizado 
         Caption         =   "Todos los Motorizados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5625
         TabIndex        =   1
         Top             =   855
         Value           =   1  'Checked
         Width           =   2715
      End
      Begin VB.CheckBox chkTurno 
         Caption         =   "Todos los Turnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5625
         TabIndex        =   4
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   435
         Left            =   1755
         TabIndex        =   7
         Top             =   2430
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56295425
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   435
         Left            =   1755
         TabIndex        =   5
         Top             =   1899
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56295425
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   435
         Left            =   3780
         TabIndex        =   6
         Top             =   1905
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   56295427
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   435
         Left            =   3780
         TabIndex        =   8
         Top             =   2430
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   56295427
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboMotorizado 
         Height          =   420
         Left            =   1755
         TabIndex        =   0
         Top             =   870
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   420
         Left            =   1755
         TabIndex        =   17
         Top             =   360
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         Caption         =   "Caja :"
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
         Height          =   240
         Index           =   0
         Left            =   1035
         TabIndex        =   18
         Top             =   450
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorizados :"
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
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   15
         Top             =   963
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
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
         Height          =   240
         Left            =   300
         TabIndex        =   14
         Top             =   2527
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
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
         Height          =   240
         Left            =   195
         TabIndex        =   13
         Top             =   1996
         Width           =   1455
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno :"
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
         Height          =   240
         Index           =   2
         Left            =   915
         TabIndex        =   12
         Top             =   1476
         Width           =   735
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
      Index           =   0
      Left            =   7035
      Picture         =   "frmRepDeliveryTicket.frx":010E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Ticketera"
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
      Left            =   5565
      Picture         =   "frmRepDeliveryTicket.frx":0200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   4950
      Top             =   3285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepDeliveryTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset
Dim RsMotorizado As Recordset
Dim RsCaja As Recordset
Dim RsTurno As Recordset
Dim sTitulo As String
Dim sCriterio As String
Dim sLinea As String
Dim xParte1 As String
Dim xParte2 As String
Dim xParte3 As String
Dim xParte4 As String
Dim nTotal As Double
Dim xTotal As Double
Dim tMotorizado As String
Dim tCaja As String
Dim tTipoPago As String
Dim nVenta As Double

Public sTurno As String
Dim xTC As Double 'Humberto 13/05/08, variable para los pagos con tarjetas de credito
Dim xES As Double 'Humberto 13/05/08, variable para los pagos con efectivo soles
Dim xED As Double 'Humberto 13/05/08, variable para los pagos con efectivo dolares
Dim xCD As Double 'Humberto 13/05/08, variable para los pagos con cheque y deposito
Dim xOt As Double 'Humberto 13/05/08, variable para los pagos con otros
Dim xPt As Double 'Humberto 13/05/08, variable para los pagos con puntos

Dim tTC As Double
Dim tES As Double
Dim tED As Double
Dim tCD As Double
Dim tOt As Double
Dim tPt As Double



Sub LlenaCombos()
    With cboMotorizado
         Isql = "Select * from vMotorizado where lActivo=1"
         Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMotorizado
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    With cboCaja
         Isql = "Select tCaja as Codigo, tDescripcion as Descripcion from TCAJA where lActivo=1"
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub

Private Sub chkMotorizado_Click()
   If chkMotorizado.value = 1 Then
      cboMotorizado.Enabled = False
      cboMotorizado.Text = ""
   Else
      cboMotorizado.Enabled = True
   End If
End Sub

Private Sub ChkCaja_Click()
   If chkCaja.value = 1 Then
      cboCaja.Enabled = False
      cboCaja.Text = ""
   Else
      cboCaja.Enabled = True
   End If
End Sub

Private Sub chkTurno_Click()
   If chkTurno.value = 1 Then
      cmdBusca.Enabled = False
      txtTurno.Text = ""
      sTurno = ""
      txtTurno.Text = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
   Else
      cmdBusca.Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 0 Then
       Unload Me
       Exit Sub
    End If
    
    sCriterio = ""
    sTitulo = ""
    If chkTurno.value = 0 Then
       If sTurno = "" Then
          MsgBox "Debe escoger un Turno", vbCritical, sMensaje
          Exit Sub
       End If
       
       Set RsTurno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
       If RsTurno.RecordCount > 0 Then
          sTitulo = " Del " & Format(RsTurno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(RsTurno!fFinal), Format(RsTurno!finicial, "dd/mmm/yyyy") & " 23:59", RsTurno!fFinal), "dd/mmm/yyyy HH:mm")
          sCriterio = " MDOCUMENTO.tTurno ='" & sTurno & "'"
       Else
          MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
       End If
    Else
       If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
          MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
          Exit Sub
          dtpFecFin.SetFocus
       End If
       sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
       sCriterio = " MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "' "
    End If
    
    If chkCaja.value = 0 Then
       If cboCaja.Text = "" Then
          MsgBox "Debe escoger una Caja", vbCritical, sMensaje
          Exit Sub
       End If
       sCriterio = sCriterio & " and MDOCUMENTO.tCaja ='" & cboCaja.BoundText & "'"
    End If
    
    If chkMotorizado.value = 0 Then
       If cboMotorizado.Text = "" Then
          MsgBox "Debe escoger un Motorizado", vbCritical, sMensaje
          Exit Sub
       End If
       sCriterio = sCriterio & "and MPEDIDO.tMotorizado ='" & cboMotorizado.BoundText & "'"
    End If
      
    Genera
    If rsReporte.EOF = True Then
       MsgBox "No se tienen Comprobantes pendientes de cobro del Canal Delivery", vbExclamation, "Mensaje del Sistema"
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
                
    'Configura la impresora la impresion Font
    Imprimir (sPreCuenta)
    Printer.FontName = sFont
    Printer.FontBold = False
    
    'Cabecera
    sLinea = "Cierre de Delivery " & IIf(chkCaja.value = True, "Todas las Cajas", cboCaja.Text)
    sLinea = String((40 - Len(Trim(sLinea))) / 2, " ") & sLinea
    Printer.Print sLinea
    sLinea = String((40 - Len(Trim(sRazonSocial))) / 2, " ") & sRazonSocial
    Printer.Print sLinea
    Printer.Print ""
    sLinea = IIf(chkTurno.value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
    ImprimeXLinea sLinea, 40, 0
    Printer.Print ""
    Printer.Print "Fecha : " & FechaServidor()
    Printer.Print String(40, "-")
       
    'Detalle
    rsReporte.MoveFirst
       
    CabeceraCaja
    CabeceraMotorizado
    CabeceraTipoPago
    nTotal = 0
    tCaja = IIf(IsNull(rsReporte!tCaja), "000", rsReporte!tCaja)
    tMotorizado = IIf(IsNull(rsReporte!tMotorizado), "0000", rsReporte!tMotorizado)
    tTipoPago = IIf(IsNull(rsReporte!tTipoPago), "00", rsReporte!tTipoPago + rsReporte!Mon)
    
    Do While Not rsReporte.EOF
       If tTipoPago <> IIf(IsNull(rsReporte!tTipoPago), "00", rsReporte!tTipoPago + rsReporte!Mon) Then
          ResumenTipoPago Mid(tTipoPago, 1, 2)
          If tCaja = IIf(IsNull(rsReporte!tCaja), "000", rsReporte!tCaja) Then
             If tMotorizado = IIf(IsNull(rsReporte!tMotorizado), "0000", rsReporte!tMotorizado) Then
                CabeceraTipoPago
             Else
                CabeceraMotorizado
                CabeceraTipoPago
             End If
          Else
             CabeceraCaja
             CabeceraMotorizado
             CabeceraTipoPago
          End If
       End If
                   
       Detalle
       tCaja = IIf(IsNull(rsReporte!tCaja), "000", rsReporte!tCaja)
       tMotorizado = IIf(IsNull(rsReporte!tMotorizado), "0000", rsReporte!tMotorizado)
       tTipoPago = IIf(IsNull(rsReporte!tTipoPago), "00", rsReporte!tTipoPago + rsReporte!Mon)
       rsReporte.MoveNext
    Loop
    
    rsReporte.MoveLast
    ResumenTipoPago Mid(tTipoPago, 1, 2)
      
    Printer.Print "Resumen Pago"
    Printer.Print Left("Tipo Pago: Efectivo " & Calcular("select tResumido as codigo from vmoneda where codigo='01'", Cn) & String(30, " "), 30) & Right(String(8, " ") & Format(tES, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Efectivo " & Calcular("select tResumido as codigo from vmoneda where codigo='02'", Cn) & String(30, " "), 30) & Right(String(8, " ") & Format(tED, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Tarjetas Credito " & String(30, " "), 30) & Right(String(8, " ") & Format(tTC, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Cheque / Deposito " & String(30, " "), 30) & Right(String(8, " ") & Format(tCD, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Otros " & String(25, " "), 30) & Right(String(8, " ") & Format(tOt, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Puntos " & String(30, " "), 30) & Right(String(8, " ") & Format(tPt, "####0.00"), 8)
    Printer.Print
      
    Printer.Print "Resumen Prepago"
    Printer.Print Left("Tipo Pago: Efectivo " & Calcular("select tResumido as codigo from vmoneda where codigo='01'", Cn) & String(30, " "), 30) & Right(String(8, " ") & Format(xES, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Efectivo " & Calcular("select tResumido as codigo from vmoneda where codigo='02'", Cn) & String(30, " "), 30) & Right(String(8, " ") & Format(xED, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Tarjetas Credito " & String(30, " "), 30) & Right(String(8, " ") & Format(xTC, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Cheque / Deposito " & String(30, " "), 30) & Right(String(8, " ") & Format(xCD, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Otros " & String(25, " "), 30) & Right(String(8, " ") & Format(xOt, "####0.00"), 8)
    Printer.Print Left("Tipo Pago: Puntos " & String(30, " "), 30) & Right(String(8, " ") & Format(xPt, "####0.00"), 8)
   
    Printer.Print ""
    Printer.Print String(40, "=")
    Printer.Print "Caja:    " & rsReporte!tCaja
    Printer.Print "Usuario: " & sUsuario
    Printer.EndDoc
    Screen.MousePointer = vbDefault
                
    If sModulo = "INFOREST" Then
       Unload Me
    End If
End Sub

Private Sub cmdBusca_Click()
   Dim xCriterio As String
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

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    dtpFecIni.value = FechaServidor()
    dtpFecFin.value = FechaServidor()
    
    cboCaja.Enabled = False
    cboCaja.Text = ""
    cboMotorizado.Enabled = False
    cboMotorizado.Text = ""
    cmdBusca.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsCaja = Nothing
   Set RsMotorizado = Nothing
   Set frmRepDeliveryTicket = Nothing
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   'Humberto 13/05/08, se cambio el ffecha por fregistro
   Isql = "SELECT dbo.MDOCUMENTO.tCaja, dbo.DPREPAGO.tTipoPago, dbo.vTipoPago.Descripcion AS TipoPago, dbo.MPEDIDO.tMotorizado, dbo.vMotorizado.Descripcion AS Motorizado, dbo.DPREPAGO.tDocumento, dbo.MDOCUMENTO.fRegistro, dbo.MDOCUMENTO.nVenta, " & _
          "dbo.MDOCUMENTO.tTurno, dbo.MDOCUMENTO.tUsuario, dbo.DPREPAGO.tMoneda, dbo.vMoneda.tResumido AS Mon, dbo.DPREPAGO.nTipoCambio, dbo.DPREPAGO.nMonto, dbo.DPREPAGO.nVuelto, dbo.TTARJETACREDITO.tDetallado AS Tarjeta, dbo.DPREPAGO.tNumero, dbo.vTipoCancelacion.Descripcion AS OtroTipo " & _
          "FROM dbo.MPEDIDO LEFT OUTER JOIN dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo RIGHT OUTER JOIN " & _
          "(SELECT DISTINCT tDocumento, tCodigoPedido FROM DDOCUMENTO) T1 INNER JOIN dbo.vMoneda INNER JOIN dbo.vTipoPago INNER JOIN " & _
          "dbo.DPREPAGO ON dbo.vTipoPago.Codigo = dbo.DPREPAGO.tTipoPago ON dbo.vMoneda.Codigo = dbo.DPREPAGO.tMoneda LEFT OUTER JOIN " & _
          "dbo.MDOCUMENTO ON dbo.DPREPAGO.tDocumento = dbo.MDOCUMENTO.tDocumento ON T1.tDocumento = dbo.DPREPAGO.tDocumento ON " & _
          "dbo.MPEDIDO.tCodigoPedido = T1.tCodigoPedido COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN " & _
          "dbo.TTARJETACREDITO ON dbo.DPREPAGO.tTarjeta = dbo.TTARJETACREDITO.tCodigoTarjeta LEFT OUTER JOIN dbo.vTipoCancelacion ON dbo.DPREPAGO.tOtroTipoPago = dbo.vTipoCancelacion.Codigo " & _
          "where tTipoPedido='02' and tEstadoDocumento='01' and " & sCriterio & " " & _
          "ORDER BY dbo.MDOCUMENTO.tCaja, dbo.MPEDIDO.tMotorizado, dbo.DPREPAGO.tTipoPago, dbo.DPREPAGO.tMoneda,  dbo.MDOCUMENTO.tDocumento"
   Set rsReporte = Lib.OpenRecordset(Isql, Cn)
End Sub

Public Sub CabeceraCaja()
   Printer.Print ""
   Printer.Print "Caja " & IIf(IsNull(rsReporte!tCaja), "000", rsReporte!tCaja)
   Printer.Print String(40, "-")
End Sub

Public Sub CabeceraMotorizado()
   Printer.Print ""
   Printer.Print IIf(IsNull(rsReporte!Motorizado), "Sin Motorizado", rsReporte!Motorizado)
   Printer.Print String(40, "=")
End Sub

Public Sub CabeceraTipoPago()
   Printer.Print " "
   Select Case IIf(IsNull(rsReporte!tTipoPago), "00", rsReporte!tTipoPago)
          Case Is = "02" 'TC
               Printer.Print "Tipo Pago: TC " & IIf(IsNull(rsReporte!tipoPago), "Sin Tipo Pago", rsReporte!tarjeta)
               Printer.Print " "
               Printer.Print "Documento          Monto  PrePago Numero"
           
          Case Is = "03" 'Cheque
               Printer.Print "Tipo Pago: Cheque / Deposito"
               Printer.Print " "
               Printer.Print "Documento          Monto  PrePago Numer."
                
          Case Is = "04" 'Documento
               Printer.Print "Tipo Pago: " & rsReporte!OtroTipo
               Printer.Print " "
               Printer.Print "Documento           Monto   PrePago"
                
          Case Is = "05" 'Puntos
               Printer.Print "Tipo Pago: Puntos"
               Printer.Print " "
               Printer.Print "Documento           Monto   PrePago"
                
          Case Else
               Printer.Print "Tipo Pago: Efectivo en " & rsReporte!Mon
               Printer.Print " "
               Printer.Print "Documento           Monto   PrePago"
   End Select
   Printer.Print String(40, "-")
End Sub

Public Sub ResumenTipoPago(tTipoPago As String)
   Printer.Print String(40, "-")
   Select Case tTipoPago
          Case Is = "02" 'Tc
               Printer.Print "Total :       " + Right(String(10, " ") & Format(xTotal, "###,##0.00"), 10) + Right(String(9, " ") & Format(nTotal, "###,##0.00"), 9)
               
          Case Is = "03" 'Cheque
               Printer.Print "Total :       " + Right(String(10, " ") & Format(xTotal, "###,##0.00"), 10) + Right(String(9, " ") & Format(nTotal, "###,##0.00"), 9)
               
          Case Is = "04" 'Documentos
               Printer.Print "Total :        " + Right(String(10, " ") & Format(xTotal, "###,##0.00"), 10) + Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
               
          Case Is = "05" 'Puntos
               Printer.Print "Total :        " + Right(String(10, " ") & Format(xTotal, "###,##0.00"), 10) + Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
               
          Case Else
               Printer.Print "Total :        " + Right(String(10, " ") & Format(xTotal, "###,##0.00"), 10) + Right(String(10, " ") & Format(nTotal, "###,##0.00"), 10)
   End Select
      
   Printer.Print " "
   nTotal = 0
   xTotal = 0
End Sub

Public Sub Detalle()
    Select Case IIf(IsNull(rsReporte!tTipoPago), "00", rsReporte!tTipoPago)
           Case Is = "02" 'TC
                xParte1 = rsReporte!tDocumento
                xParte2 = Right(String(8, " ") & Format(rsReporte!nVenta, "####0.00"), 8)
                xParte3 = Right(String(8, " ") & Format(rsReporte!nMonto, "####0.00"), 8)
                xParte4 = Mid(rsReporte!tNumero, 1, 6)
                xTC = xTC + rsReporte!nMonto
                tTC = tTC + rsReporte!nVenta
                
           Case Is = "03" 'Cheque
                xParte1 = rsReporte!tDocumento
                xParte2 = Right(String(8, " ") & Format(rsReporte!nVenta, "####0.00"), 8)
                xParte3 = Right(String(8, " ") & Format(rsReporte!nMonto, "####0.00"), 8)
                xParte4 = Mid(rsReporte!tNumero, 1, 6)
                xCD = xCD + rsReporte!nMonto
                tCD = tCD + rsReporte!nVenta
           
           Case Is = "04" 'Documento
                xParte1 = rsReporte!tDocumento
                xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta, "##,##0.00"), 9)
                xParte3 = Right(String(9, " ") & Format(rsReporte!nMonto, "##,##0.00"), 9)
                xParte4 = rsReporte!OtroTipo
                xOt = xOt + rsReporte!nMonto
                tOt = tOt + rsReporte!nVenta
           
           Case Is = "05" 'Puntos
                xParte1 = rsReporte!tDocumento
                xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta, "##,##0.00"), 9)
                xParte3 = Right(String(9, " ") & Format(rsReporte!nMonto, "##,##0.00"), 9)
                xParte4 = ""
                xPt = xPt + rsReporte!nMonto
                tPt = tPt + rsReporte!nVenta
           
           Case Else 'Efectivo
                xParte1 = rsReporte!tDocumento
                xParte2 = Right(String(9, " ") & Format(rsReporte!nVenta, "##,##0.00"), 9)
                xParte3 = Right(String(9, " ") & Format(rsReporte!nMonto, "##,##0.00"), 9)
                xParte4 = ""
                If rsReporte!tMoneda = "01" Then
                    xES = xES + rsReporte!nMonto
                    tES = tES + rsReporte!nVenta
                Else
                    xED = xED + rsReporte!nMonto
                    tED = tED + rsReporte!nVenta
                End If
    End Select
    
    nTotal = nTotal + rsReporte!nMonto
    xTotal = xTotal + rsReporte!nVenta
    sLinea = xParte1 & " " & xParte2 & " " & xParte3 & " " & xParte4
    Printer.Print sLinea
End Sub
