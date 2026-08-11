VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmReportNotaCredito 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Notas de Credito"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form2"
   ScaleHeight     =   3780
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Valores"
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton CmdEmite 
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
         Left            =   1440
         Picture         =   "frmReportNotaCredito.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton CmdSalir 
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
         Left            =   3000
         Picture         =   "frmReportNotaCredito.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CheckBox ChkFecha 
         Caption         =   "Fechas"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Frame FrameFecha 
         Height          =   1215
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   2775
         Begin MSComCtl2.DTPicker dtpFechaFinal 
            Height          =   375
            Left            =   1200
            TabIndex        =   6
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   88801281
            CurrentDate     =   42790
         End
         Begin MSComCtl2.DTPicker dtpFechaInicial 
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   88801281
            CurrentDate     =   42790
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Final:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicial:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CheckBox ChkEstadoNota 
         Caption         =   "Estado de Notas de Credito :"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox ChkFranjaHoraria 
         Caption         =   "Franja Horaria"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox ChkCaja 
         Caption         =   "Caja"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox ChkTurno 
         Caption         =   "Turno"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   735
      End
      Begin VB.Frame FrameHorario 
         Height          =   1215
         Left            =   3120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
         Begin MSComCtl2.DTPicker dtpHoraFinal 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   88801283
            UpDown          =   -1  'True
            CurrentDate     =   37541.9993055556
         End
         Begin MSComCtl2.DTPicker dtpHoraInicial 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   88801283
            UpDown          =   -1  'True
            CurrentDate     =   37541
         End
      End
      Begin MSDataListLib.DataCombo cboTurno 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cboEstadoNota 
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   195
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
   End
End
Attribute VB_Name = "frmReportNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dsrRepNotaCredito As New dsrRepNotaCredito
Sub LlenaCombos()
On Error GoTo fin:
Dim i As Integer
Dim RsEstado As New Recordset
Dim RsCaja As New Recordset

    With cboEstadoNota
         Isql = "Select * from vEstadoDocumento where lActivo=1"
         Set RsEstado = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsEstado
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
           
    With cboCaja
         Isql = "select tCaja AS Codigo,tDescripcion as Descripcion from  TCAJA where lActivo=1"
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
   
    Call LlenarCajaTurno(1)
    
    Exit Sub
fin:
      MsgBox ("Error verificar los datos de los Combos")
           
End Sub
Private Sub LlenarCajaTurno(Index As Integer)
Dim Rsturno As New Recordset

    Select Case Index
        Case Is = 1
           With cboTurno
                Isql = "select tTurno as Codigo, tTurno AS Descripcion from MTURNO "
                Set Rsturno = Lib.OpenRecordset(Isql, Cn)
                Set .RowSource = Rsturno
                    .DataField = "Descripcion"
                    .ListField = "Descripcion"
                    .BoundColumn = "Codigo"
            End With
        Case Is = 2
        
          Dim caja As String
          If cboCaja.BoundText <> "" Then
          caja = "where tCaja='" & cboCaja.BoundText & "'"
          End If
          With cboTurno
                Isql = "select tTurno as Codigo, tTurno AS Descripcion from MTURNO  " & caja
                Set Rsturno = Lib.OpenRecordset(Isql, Cn)
                Set .RowSource = Rsturno
                    .DataField = "Descripcion"
                    .ListField = "Descripcion"
                    .BoundColumn = "Codigo"
          End With
          caja = ""
    End Select
             
    
End Sub

Private Sub cboCaja_Change()
cboTurno.Text = ""
Call LlenarCajaTurno(2)
End Sub

Private Sub ChkCaja_Click()
On Error GoTo fin
    If ChkCaja.value = 1 Then
        cboCaja.Enabled = True
        Call LlenarCajaTurno(2)
    Else
        cboCaja.Enabled = False
        cboCaja.Text = ""
        Call LlenarCajaTurno(1)
    End If
Exit Sub
fin:
MsgBox ("Error Interno")
End Sub

Private Sub ChkEstadoNota_Click()
    If ChkEstadoNota = 1 Then
        cboEstadoNota.Enabled = True
    Else
        cboEstadoNota.Enabled = False
        cboEstadoNota.Text = ""
    End If
End Sub

Private Sub ChkFecha_Click()
    If ChkFecha.value = 1 Then
        FrameFecha.Enabled = True
    Else
        FrameFecha.Enabled = False
    End If
End Sub

Private Sub ChkFranjaHoraria_Click()
    If ChkFranjaHoraria.value = 1 Then
        FrameHorario.Enabled = True
    Else
        FrameHorario.Enabled = False
    End If
End Sub

Private Sub ChkTurno_Click()
    If chkTurno.value = 1 Then
        cboTurno.Enabled = True
    Else
        cboTurno.Enabled = False
        cboTurno.Text = ""
    End If
End Sub

Private Sub CmdEmite_Click(Index As Integer)
On Error GoTo fin
Call Genera
Screen.MousePointer = vbDefault
Exit Sub
fin:
Screen.MousePointer = vbDefault
MsgBox (error)
End Sub

Private Sub CmdSalir_Click(Index As Integer)
Unload Me
End Sub


Private Sub Form_Load()
On Error GoTo fin
Centrar Me
cboEstadoNota.Enabled = False
cboCaja.Enabled = False
cboTurno.Enabled = False
FrameFecha.Enabled = False
FrameHorario.Enabled = False

Call LlenaCombos
Exit Sub
fin:
MsgBox ("Error validar")
End Sub

Public Sub Genera()
 Screen.MousePointer = vbHourglass
    Dim sFecha1 As String
    Dim sFecha2 As String
    Dim hora1 As String
    Dim hora2 As String
'    Dim sMes As String
'    Dim sDia As String
'
    'CESAR 0020-2013
'    If chkCFacturados.value = 0 Then
'        'VENTA
''        If optOrigen(0).value = True Then
''            sOrigen = "00"
''            If optValor(0).value Then
''               sTipo = "00"
''               sPrecio = "dbo.DPEDIDO.nVenta"
''               sTexto = " con Precios de Venta"
''            ElseIf optValor(1).value Then
''               sTipo = "00"
''               sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
''               sTexto = " con Precios Netos"
''            Else
''               sTipo = "01"
''               sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
''               sTexto = " con Precios de Costo"
''            End If
''        End If
'        'COMBINACION
''        If optOrigen(4).value = True Then
''            sOrigen = "04"
''            If optValor(0).value Then
''               sTipo = "00"
''               sPrecio = "dbo.CPEDIDO.nVenta"
''               sTexto = " con Precios de Venta"
''            ElseIf optValor(1).value Then
''               sTipo = "00"
''               sPrecio = "dbo.CPEDIDO.nPrecioNeto * dbo.CPEDIDO.nCantidad"
''               sTexto = " con Precios Netos"
''            Else
''               sTipo = "01"
''               sPrecio = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
''               sTexto = " con Precios de Costo"
''            End If
''        End If
'
'    Else
        'VENTA
'        If optOrigen(0).value = True Then
'            sOrigen = "00"
'            If optValor(0).value Then
'                sTipo = "00"
'                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nVenta end "
'                sTexto = " con Precios de Venta"
'            ElseIf optValor(1).value Then
'                sTipo = "00"
'                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad end "
'                sTexto = " con Precios Netos"
'            Else
'               sTipo = "01"
'               sPrecio = " case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad end "
'               sTexto = " con Precios de Costo"
'            End If
'        End If
'        'COMBINACION
'        If optOrigen(4).value = True Then
'            sOrigen = "04"
'            If optValor(0).value Then
'                sTipo = "00"
'                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nVenta end "
'                sTexto = " con Precios de Venta"
'            ElseIf optValor(1).value Then
'                sTipo = "00"
'                sPrecio = "case when dpedido.lregistroventa=0 then 0 else dbo.CPEDIDO.nPrecioNeto * dbo.CPEDIDO.nCantidad end "
'                sTexto = " con Precios Netos"
'            Else
'               sTipo = "01"
'               sPrecio = " case when dpedido.lregistroventa=0 then 0 else (isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad end "
'               sTexto = " con Precios de Costo"
'            End If
'        End If
        
'    End If
'

    
    
    'CESAR 0020-2013
'    If sOrigen = "00" Then
'
'       If sTipo = "00" Then '--------VENTAS y NETO
'           Isql = "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
'                "union all " & _
'                "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
'                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
'       Else '----COSTOS ---
'            Isql = "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TPRODUCTOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nGasto,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'                "LEFT JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTOPROPIEDAD.tProducto AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem AND TPRODUCTOPROPIEDAD.tEnlace<>'' " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
'                "union all " & _
'                "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.DPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TPRODUCTOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nGasto,0) + ISNULL(dbo.TPRODUCTOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido " & _
'                "LEFT JOIN dbo.TPRODUCTOPROPIEDAD ON dbo.DPEDIDO.tCodigoPedido = dbo.TPRODUCTOPROPIEDAD.tCodigoPedido AND dbo.DPEDIDO.tCodigoProducto = dbo.TPRODUCTOPROPIEDAD.tProducto AND dbo.DPEDIDO.tItem = dbo.TPRODUCTOPROPIEDAD.tItem AND TPRODUCTOPROPIEDAD.tEnlace<>'' " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
'                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
'        End If
'    Else
'        If sTipo = "00" Then '--------VENTA y NETO DE COMBINACIONES
'            Isql = "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
'                "union all " & _
'                "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
'                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
'        Else ' ----COSTOS ----
'            Isql = "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TCOMBOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nGasto,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
'                "INNER JOIN dbo.tComboPropiedad ON dbo.tComboPropiedad.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.tComboPropiedad.tProducto = dbo.CPEDIDO.tProductoCombo AND dbo.tComboPropiedad.tItem = dbo.CPEDIDO.tItemCombo " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha1 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) - 1 & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) " & _
'                "union all " & _
'                "(SELECT dbo.vSalon.tLocal AS tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end) AS ANIO, " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end) AS Fecha, " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo, SUM(dbo.CPEDIDO.nCantidad) AS nCantidad, SUM(" & sPrecio & ") + SUM((ISNULL(dbo.TCOMBOPROPIEDAD.nInsumo,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nGasto,0) + ISNULL(dbo.TCOMBOPROPIEDAD.nManoObra,0))) AS nVenta, max(vproducto.grupo) as grupo,max(vproducto.subgrupo) as subgrupo,max(vProducto.descripcion) as nomprod,max(vProducto.Codigo) as tcodigoproducto " & _
'                "FROM dbo.vSalon RIGHT OUTER JOIN dbo.MPEDIDO ON dbo.vSalon.Codigo = dbo.MPEDIDO.tSalon RIGHT OUTER JOIN dbo.DPEDIDO LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo ON " & _
'                "dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.CPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido " & _
'                "INNER JOIN dbo.tComboPropiedad ON dbo.tComboPropiedad.tCodigoPedido = dbo.CPEDIDO.tCodigoPedido AND dbo.tComboPropiedad.tProducto = dbo.CPEDIDO.tProductoCombo AND dbo.tComboPropiedad.tItem = dbo.CPEDIDO.tItemCombo " & _
'                "WHERE dbo.MPEDIDO.tEstadoPedido <> '03' AND dbo.DPEDIDO.tEstadoItem = 'N' AND (dbo.DPEDIDO.tFacturado = 'P' OR dbo.DPEDIDO.tFacturado = 'F') and " & sFecha2 & sCriterio & _
'                "GROUP BY  dbo.vSalon.tLocal, (case when  " & Year(dtpAnual.value) & " = year(Mpedido.fregistro) then year(Mpedido.fRegistro) else year(Mpedido.fRegistro)-1 end), " & _
'                "(case when day(mpedido.fRegistro)=1 and DATEPART(hh,Mpedido.fregistro) >= " & Hour(dtpHora.value) & " then month(Mpedido.fRegistro) else month(dateadd(day, -1, Mpedido.fRegistro))  end), " & _
'                "dbo.vProducto.tGrupo, dbo.vProducto.tSubGrupo,dbo.dpedido.tCodigoProducto) "
'                '"ORDER BY  vproducto.grupo,vproducto.subgrupo,nomprod"
'        End If
'    End If
'
'
'    Set RsTemporal = Lib.OpenRecordset(Isql, Cn)
'
'    If RsTemporal.RecordCount = 0 Then
'      MsgBox "No hay Datos para Mostrar", vbInformation, sMensaje
'      Exit Sub
'    End If
'    Cn.Execute "delete from " & sDetalle
'
'    Dim sLinea As String
'    Dim i As Integer
    
'    RsTemporal.MoveFirst
'    Do While Not RsTemporal.EOF
'
'       If Not IsNull(RsTemporal!Fecha) Then
'
'          If optAgrupacion(0).value = True Then
'             If Calcular("select count(tGrupo1) as Codigo from " & sDetalle & " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tgrupo), "", RsTemporal!tgrupo) & "'", Cn) > 0 Then
'                Isql = "update " & sDetalle & " set C" & RsTemporal!Fecha & "= C" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nCantidad), 0, RsTemporal!nCantidad) & ", M" & RsTemporal!Fecha & "= M" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.value), 0, RsTemporal!nVenta) & ", Mt" & RsTemporal!Fecha & "= Mt" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.value) - 1, 0, RsTemporal!nVenta) & _
'                       " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tgrupo), "", RsTemporal!tgrupo) & "'"
'             Else
'                sLinea = ""
'
'                For i = 1 To 12
'                    If RsTemporal!Fecha <> i Then
'                       sLinea = sLinea & "C" & Trim(str(i)) & ", M" & Trim(str(i)) & ",Mt" & Trim(str(i)) & ","
'                    End If
'                Next i
'                sLinea = sLinea & " C" & RsTemporal!Fecha & ", M" & RsTemporal!Fecha & ", Mt" & RsTemporal!Fecha
'                Isql = "insert " & sDetalle & " (linea,tGrupo1, tGrupo2, " & sLinea & ") " & _
'                       "values ('" & IIf(IsNull(RsTemporal!Grupo), "Sin Grupo...", RsTemporal!Grupo) & "','" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "', '" & IIf(IsNull(RsTemporal!tgrupo), "", RsTemporal!tgrupo) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & RsTemporal!nCantidad & ", " & IIf(RsTemporal!anio = Year(dtpAnual.value) - 1, RsTemporal!nVenta, 0) & "," & IIf(RsTemporal!anio = Year(dtpAnual.value), RsTemporal!nVenta, 0) & " )"
'             End If
'
'          ElseIf optAgrupacion(1).value = True Then
'
'             If Calcular("select count(tGrupo1) as Codigo from " & sDetalle & " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tSubGrupo), "", RsTemporal!tSubGrupo) & "'", Cn) > 0 Then
'                Isql = "update " & sDetalle & " set C" & RsTemporal!Fecha & "= C" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nCantidad), 0, RsTemporal!nCantidad) & ", M" & RsTemporal!Fecha & "= M" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.value), 0, RsTemporal!nVenta) & ", Mt" & RsTemporal!Fecha & "= Mt" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.value) - 1, 0, RsTemporal!nVenta) & _
'                       " where tGrupo1 ='" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "' and tGrupo2 = '" & IIf(IsNull(RsTemporal!tSubGrupo), "", RsTemporal!tSubGrupo) & "'"
'             Else
'                sLinea = ""
'                For i = 1 To 12
'                    If RsTemporal!Fecha <> i Then
'                       sLinea = sLinea & "C" & Trim(str(i)) & ", M" & Trim(str(i)) & ",Mt" & Trim(str(i)) & ","
'                    End If
'                Next i
'                sLinea = sLinea & " C" & RsTemporal!Fecha & ", M" & RsTemporal!Fecha & ", Mt" & RsTemporal!Fecha
'                Isql = "insert " & sDetalle & " (linea,tGrupo1, tGrupo2, " & sLinea & ") " & _
'                       "values ('" & IIf(IsNull(RsTemporal!SubGrupo), "Sin Sub-Grupo...", RsTemporal!SubGrupo) & "','" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "', '" & IIf(IsNull(RsTemporal!tSubGrupo), "", RsTemporal!tSubGrupo) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & RsTemporal!nCantidad & ", " & IIf(RsTemporal!anio = Year(dtpAnual.value) - 1, RsTemporal!nVenta, 0) & "," & IIf(RsTemporal!anio = Year(dtpAnual.value), RsTemporal!nVenta, 0) & " )"
'             End If
'          Else
'
'             If Calcular("select count(tGrupo2) as Codigo from " & sDetalle & " where tGrupo2 ='" & IIf(IsNull(RsTemporal!tCodigoProducto), "", RsTemporal!tCodigoProducto) & "'", Cn) > 0 Then
'                Isql = "update " & sDetalle & " set C" & RsTemporal!Fecha & "= C" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nCantidad), 0, RsTemporal!nCantidad) & ", M" & RsTemporal!Fecha & "= M" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.value), 0, RsTemporal!nVenta) & ", Mt" & RsTemporal!Fecha & "= Mt" & RsTemporal!Fecha & " + " & IIf(IsNull(RsTemporal!nVenta) Or RsTemporal!anio = Year(dtpAnual.value) - 1, 0, RsTemporal!nVenta) & _
'                       " where tGrupo2 ='" & IIf(IsNull(RsTemporal!tCodigoProducto), "", RsTemporal!tCodigoProducto) & "'"
'             Else
'                sLinea = ""
'
'
'                For i = 1 To 12
'                    If RsTemporal!Fecha <> i Then
'                       sLinea = sLinea & "C" & Trim(str(i)) & ", M" & Trim(str(i)) & ",Mt" & Trim(str(i)) & ","
'                    End If
'                Next i
'                sLinea = sLinea & " C" & RsTemporal!Fecha & ", M" & RsTemporal!Fecha & ", Mt" & RsTemporal!Fecha
'                Isql = "insert " & sDetalle & " (linea,tGrupo1, tGrupo2, " & sLinea & ") " & _
'                       "values ('" & IIf(IsNull(RsTemporal!nomprod), "Sin Nombre...", RsTemporal!nomprod) & "','" & IIf(IsNull(RsTemporal!tLocal), "", RsTemporal!tLocal) & "', '" & IIf(IsNull(RsTemporal!tCodigoProducto), "", RsTemporal!tCodigoProducto) & "', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, " & RsTemporal!nCantidad & ", " & IIf(RsTemporal!anio = Year(dtpAnual.value) - 1, RsTemporal!nVenta, 0) & "," & IIf(RsTemporal!anio = Year(dtpAnual.value), RsTemporal!nVenta, 0) & " )"
'             End If
'          End If
'          Cn.Execute Isql
'       End If
'       RsTemporal.MoveNext
'    Loop

    Dim Estado As String
    Dim rsReporte As Recordset
    If ChkEstadoNota.value = 1 Then
        If cboEstadoNota.Text = "" Then MsgBox ("seleccione un Estado de Nota de Credito"): Exit Sub
        Estado = " WHERE  MN.tEstadoDocumento ='" & cboEstadoNota.BoundText & "'"
    Else
        Estado = ""
    End If
    
    If ChkFecha.value = 1 Then
        If dtpFechaInicial.value > dtpFechaFinal.value Then
            MsgBox ("Fechas Incorrectas")
            Exit Sub
        Else
            sFecha1 = str(dtpFechaInicial.value)
            sFecha2 = str(dtpFechaFinal.value)
            If Estado = "" Then
                Estado = Estado & " where convert(date,MN.fRegistro) between convert(date,'" & sFecha1 & "') and convert(date,'" & sFecha2 & "')"
            Else
                Estado = Estado & " and convert(date,MN.fRegistro) between convert(date,'" & sFecha1 & "') and convert(date,'" & sFecha2 & "')"
            End If
            'sAno1 = LTrim(str(Year(dtpAnual.value)))
            'sAno2 = LTrim(str(Year(dtpAnual.value) - 1))
        
            'sFecha1 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno2 & "/01/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno2 & "/12/31') "
            'sFecha2 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno1 & "/01/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno1 & "/12/31') "
        End If
    End If
    
   If ChkFranjaHoraria.value = 1 Then
        If dtpHoraInicial.value > dtpHoraFinal.value Then
            MsgBox ("Horas Incorrectas")
            Exit Sub
           
        Else
            hora1 = Format(dtpHoraInicial.value, "HH:mm") 'str(dtpHoraInicial.value)
            hora2 = Format(dtpHoraFinal.value, "HH:mm")
            If Estado = "" Then
                Estado = Estado & " where convert(time,MN.fRegistro) between convert(time,'" & hora1 & "') and convert(time,'" & hora2 & "')"
            Else
                Estado = Estado & " and convert(time,MN.fRegistro) between convert(time,'" & hora1 & "') and convert(time,'" & hora2 & "')"
            End If
            'sAno1 = LTrim(str(Year(dtpAnual.value)))
            'sAno2 = LTrim(str(Year(dtpAnual.value) - 1))
        
            'sFecha1 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno2 & "/01/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno2 & "/12/31') "
            'sFecha2 = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno1 & "/01/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno1 & "/12/31') "
        End If
    End If
    
    If ChkCaja.value = 1 Then
           If cboCaja.Text = "" Then MsgBox ("seleccione una Caja"): Exit Sub
           If Estado = "" Then
                Estado = " where  mn.tCaja ='" & cboCaja.BoundText & "'"
           Else
                Estado = Estado & " and mn.tCaja ='" & cboCaja.BoundText & "'"
           End If
    End If
    If chkTurno.value = 1 Then
           If cboTurno.Text = "" Then MsgBox ("seleccione una Turno"): Exit Sub
           If Estado = "" Then
                Estado = " where  mn.tTurno ='" & cboTurno.BoundText & "'"
           Else
                Estado = Estado & " and mn.tTurno ='" & cboTurno.BoundText & "'"
           End If
    End If

    Isql = "set dateformat dmy SELECT MN.tTurno AS TURNO,mn.tCaja as CAJA,MN.tNotaCredito, MN.tDocumento AS NCDOCUMENEnlace,ISNULL(mn.nNeto,0) AS EnlNETO,ISNULL(mn.nImpuesto1,0) AS EnlIMPUESTO1 ,ISNULL(mn.nImpuesto1,0) AS EnlIMPUESTO2," & _
            " MN.nVenta as MontoCreacion, isnull(dp.tDocumento,'') as NCDOCUMENPagado , ISNULL(MD.nNeto,0) AS PagoNETO, ISNULL(MD.nPrecioImpuesto1,0) AS PagoIMPUESTO1,ISNULL(MD.nPrecioImpuesto2,0) AS PagoIMPUESTO2," & _
            " ISNULL(MD.nVenta,0) AS PagoVENTA, ISNULL(MD.nVuelto,0) AS PagoVUELTO," & _
            " isnull((select top 1 tCodigoPedido from DDOCUMENTO where tDocumento=md.tdocumento),'') as CodPedido," & _
            " (SELECT LTRIM(RTRIM(Descripcion)) FROM vEstadoDocumento WHERE CODIGO= MN.tEstadoDocumento) AS ESTADO," & _
            " convert(datetime,MN.fRegistro) as FECHA_NOTACREDITO" & _
            " ,ISNULL((SELECT LTRIM(RTRIM(tEmpresa)) FROM TCLIENTE WHERE tCodigoCliente=MD.tCodigoCliente) ,'')AS CLIENTE_PAGO" & _
            " ,ISNULL((SELECT LTRIM(RTRIM(tEmpresa)) FROM TCLIENTE WHERE tCodigoCliente=(SELECT top 1 tCodigoCliente FROM MDOCUMENTO WHERE tDocumento=MN.tDocumento)) ,'')AS CLIENTE_NOTACREDITO" & _
            " FROM MNOTACREDITO MN LEFT OUTER JOIN DPAGODOCUMENTO DP  ON MN.tNotaCredito= DP.tNumero" & _
            " LEFT OUTER JOIN  MDOCUMENTO MD on DP.tDocumento=MD.tDocumento AND MD.tEstadoDocumento='02' " & Estado & " ORDER BY MN.tTurno  DESC " 'MN.fRegistro
            
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
      
    dsrRepNotaCredito.DiscardSavedData
    dsrRepNotaCredito.Database.SetDataSource rsReporte
    frmEmite.CRViewer.ReportSource = dsrRepNotaCredito
    dsrRepNotaCredito.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    dsrRepNotaCredito.PaperOrientation = crLandscape
    frmEmite.CRViewer.DisplayGroupTree = False
    frmEmite.CRViewer.ViewReport
    frmEmite.Show vbModal
Screen.MousePointer = vbDefault
End Sub




Private Sub Label1_Click()

End Sub

