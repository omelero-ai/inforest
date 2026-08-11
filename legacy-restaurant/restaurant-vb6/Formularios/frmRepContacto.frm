VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepContacto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Contactos"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmRepContacto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
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
      Left            =   4830
      Picture         =   "frmRepContacto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3135
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
      Height          =   3030
      Left            =   0
      TabIndex        =   13
      Top             =   45
      Width           =   7725
      Begin VB.CheckBox chkTipoProducto 
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
         Left            =   5400
         TabIndex        =   24
         Top             =   1179
         Value           =   1  'Checked
         Width           =   2220
      End
      Begin VB.CheckBox chkContacto 
         Caption         =   "Todos los Contactos"
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
         Left            =   5400
         TabIndex        =   20
         Top             =   375
         Value           =   1  'Checked
         Width           =   2070
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
         Index           =   1
         Left            =   4500
         Picture         =   "frmRepContacto.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   315
         Width           =   765
      End
      Begin VB.TextBox txtContacto 
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
         TabIndex        =   18
         Top             =   360
         Width           =   2535
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
         Left            =   4500
         Picture         =   "frmRepContacto.frx":0930
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1536
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
         TabIndex        =   2
         Top             =   1566
         Width           =   2535
      End
      Begin VB.CheckBox chkEstadoPedido 
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
         Left            =   5400
         TabIndex        =   1
         Top             =   777
         Value           =   1  'Checked
         Width           =   2220
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
         Left            =   5400
         TabIndex        =   12
         Top             =   1581
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1905
         TabIndex        =   6
         Top             =   2400
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   63963137
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1905
         TabIndex        =   4
         Top             =   1968
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   63963137
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3810
         TabIndex        =   5
         Top             =   1953
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3810
         TabIndex        =   7
         Top             =   2385
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboEstadoPedido 
         Height          =   315
         Left            =   1905
         TabIndex        =   0
         Top             =   765
         Width           =   3360
         _ExtentX        =   5927
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
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   1905
         TabIndex        =   23
         Top             =   1170
         Width           =   3360
         _ExtentX        =   5927
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
         Caption         =   "Tipo de Producto :"
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
         Left            =   195
         TabIndex        =   22
         Top             =   1230
         Width           =   1605
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto :"
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
         Index           =   10
         Left            =   900
         TabIndex        =   21
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Pedido :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   825
         Width           =   1680
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
         Left            =   675
         TabIndex        =   16
         Top             =   2475
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
         Left            =   570
         TabIndex        =   15
         Top             =   2040
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
         Left            =   1170
         TabIndex        =   14
         Top             =   1620
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
      Left            =   6285
      Picture         =   "frmRepContacto.frx":0A32
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3135
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
      Left            =   3375
      Picture         =   "frmRepContacto.frx":0B24
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3135
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
      Left            =   1920
      Picture         =   "frmRepContacto.frx":1056
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3135
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1170
      Top             =   3195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sInfhotelRuta As String
Dim sInfhotelMDB As String
Dim rsReporte As Recordset
Dim RsTipoProducto As Recordset
Dim RsEstadoPedido As Recordset
Dim Rsturno As Recordset
Dim Reporte As New dsrContacto
Dim sContacto As String
Dim sCriterio As String
Dim sTitulo As String

Sub LlenaCombos()
    With cboEstadoPedido
         Isql = "Select * from vEstadoPedido"
         Set RsEstadoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsEstadoPedido
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    With cboTipoProducto
         Isql = "Select * from vTipoProducto"
         Set RsTipoProducto = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoProducto
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub

Private Sub chkContacto_Click()
   If chkContacto.Value = 1 Then
      sContacto = ""
      txtContacto.Text = ""
      cmdBusca(1).Enabled = False
   Else
      cmdBusca(1).Enabled = True
   End If
End Sub

Private Sub chkTipoProducto_Click()
   If chkTipoProducto.Value = 1 Then
      cboTipoProducto.Enabled = False
      cboTipoProducto.Text = ""
   Else
      cboTipoProducto.Enabled = True
   End If
End Sub

Private Sub chkEstadoPedido_Click()
   If chkEstadoPedido.Value = 1 Then
      cboEstadoPedido.Enabled = False
      cboEstadoPedido.Text = ""
   Else
      cboEstadoPedido.Enabled = True
   End If
End Sub


Private Sub cmdBusca_Click(Index As Integer)
    If Index = 1 Then
       Isql = "SELECT " & sInfhotelMDB & ".dbo.TCOMPANIA.tCodigoCompania + " & sInfhotelMDB & ".dbo.TCONTACTO.tCodigoContacto as Codigo, " & sInfhotelMDB & ".dbo.TCOMPANIA.tRazonSocialCompania, rtrim(" & sInfhotelMDB & ".dbo.TCONTACTO.tCarnet) + ' - ' +Rtrim(" & sInfhotelMDB & ".dbo.TCONTACTO.tPaterno) + ' ' + Rtrim(" & sInfhotelMDB & ".dbo.TCONTACTO.tMaterno) as Descripcion, " & sInfhotelMDB & ".dbo.TCONTACTO.tNombre, " & sInfhotelMDB & ".dbo.TCONTACTO.tCarnet " & _
              "FROM " & sInfhotelMDB & ".dbo.TCOMPANIA INNER JOIN " & sInfhotelMDB & ".dbo.TCONTACTO ON " & sInfhotelMDB & ".dbo.TCOMPANIA.tCodigoCompania = " & sInfhotelMDB & ".dbo.TCONTACTO.tCodigoCompania " & _
              "Where lInforest = 1 order by tpaterno, tmaterno, tnombre"
       frmBusca.cboCriterio.Enabled = True
       frmBusca.nPredeterm = 2
       Call ConfGrilla(3, frmBusca.grdGrilla, "Compañía", 2, "tRazonSocialCompania", 3400, 0, 0, "", _
                                              "Apellido", 2, "Descripcion", 3000, 0, 0, "", _
                                              "Nombre", 2, "tNombre", 1800, 0, 0, "")
       frmBusca.Show vbModal
       If Not wEnter = True Then
          Exit Sub
       End If
       sContacto = Mid(sCodigo, 6, 4)
       txtContacto.Text = sDescrip
    Else
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
    End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   sTitulo = ""
   
   If chkTurno.Value = 0 Then
      If sTurno = "" Then
         MsgBox "Debe escoger un Turno", vbCritical, sMensaje
         Exit Sub
      End If
            
      Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
      If Rsturno.RecordCount > 0 Then
         sTitulo = " Del " & Format(Rsturno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(Rsturno!fFinal), Format(Rsturno!finicial, "dd/mmm/yyyy") & " 23:59", Rsturno!fFinal), "dd/mmm/yyyy HH:mm")
         sCriterio = " MPEDIDO.tTurno ='" & sTurno & "'"
      Else
         MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
      End If
      
   Else
      If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
      sCriterio = " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "' "
   End If
   
   If chkTipoProducto.Value = 0 Then
      If cboTipoProducto.Text = "" Then
         MsgBox "Debe escoger un Tipo de producto", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and vProducto.tTipoProducto ='" & cboTipoProducto.BoundText & "'"
   End If
         
   If chkEstadoPedido.Value = 0 Then
      If cboEstadoPedido.Text = "" Then
         MsgBox "Debe escoger un Estado", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and tEstadoPedido ='" & cboEstadoPedido.BoundText & "'"
   End If
   
   If chkContacto.Value = 0 Then
      If sContacto = "" Then
         MsgBox "Debe escoger un Contacto", vbCritical, sMensaje
         Exit Sub
       End If
       sCriterio = sCriterio & " and MPEDIDO.tContacto = '" & sContacto & "'"
   End If
      
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If

               frmEmite.CRViewer.DisplayGroupTree = False
               Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               Reporte.PaperOrientation = crPortrait
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
               Reporte.PaperOrientation = crPortrait
               Reporte.PrintOut
          
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If

               Reporte.ExportOptions.FormatType = 21
               Reporte.ExportOptions.DestinationType = 1
               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
               cmdSave.ShowSave
               If cmdSave.FileName = "" Then
                  Exit Sub
               End If
               Reporte.ExportOptions.DiskFileName = cmdSave.FileName
               Reporte.Export False
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub ChkTurno_Click()
   If chkTurno.Value = 1 Then
      cmdBusca(0).Enabled = False
      sTurno = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
   Else
      cmdBusca(0).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
      
    Open App.Path & "\INFHOTEL.INI" For Input As #1   ' Abre el archivo para recibir los datos.
    Do While Not EOF(1)                               ' Repite el bucle hasta el final del archivo.
        Input #1, sInfhotelRuta, sInfhotelMDB, sCajaInfhotel
    Loop
    Close #1   ' Cierra el archivo.
    sInfhotelRuta = IIf(Right(Trim(sInfhotelRuta), 1) = "\", sInfhotelRuta, sInfhotelRuta)
    sInfhotelMDB = IIf(UCase(Right(Trim(sInfhotelMDB), 4)) = ".MDB", sInfhotelMDB, sInfhotelMDB)
    Set CnInfhotel = New Connection
    CnInfhotel.Provider = "SQLOLEDB"
    CnInfhotel.CursorLocation = adUseServer
    CnInfhotel.ConnectionString = "User ID=" & sUserName & _
                                  ";password=" & sUserPassword & _
                                  ";Data Source=" & sInfhotelRuta & _
                                  ";Initial Catalog=" & sInfhotelMDB
    CnInfhotel.Open
    
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    sTurno = ""
    cboEstadoPedido.Enabled = False
    cboTipoProducto.Enabled = False
    cboEstadoPedido.Text = ""
    cboTipoProducto.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set Rsturno = Nothing
   Set RsTipoProducto = Nothing
   Set RsEstadoPedido = Nothing
   Set frmRepPedido = Nothing
End Sub

Public Sub Genera()
   Dim sOrden As String
   Screen.MousePointer = vbHourglass
   
   Isql = "SELECT dbo.MPEDIDO.fRegistro, dbo.DPEDIDO.tCodigoPedido, " & sInfhotelMDB & ".dbo.TCONTACTO.tCarnet AS Codigo, RTRIM(" & sInfhotelMDB & ".dbo.TCONTACTO.tPaterno) + ' ' + RTRIM(" & sInfhotelMDB & ".dbo.TCONTACTO.tMaterno) + ' ' + RTRIM(" & sInfhotelMDB & ".dbo.TCONTACTO.tNombre) AS Socio, " & _
          "(case when dbo.MPEDIDO.tTipoPedido='03' then 'R/S' else (case when dbo.MPEDIDO.tTipoPedido='01' then cast(dbo.MPEDIDO.nAdulto as nvarchar(4)) else '1' end) end) as Pax, SUM(dbo.DPEDIDO.nVenta) AS nVenta, SUM(dbo.DPEDIDO.nDescuento) AS nDescuento, " & _
          "SUM(dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad) AS nNeto, dbo.vEstadoPedido.Descripcion AS EstadoPedido " & _
          "FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN " & sInfhotelMDB & ".dbo.TCONTACTO ON dbo.MPEDIDO.tCompania = " & sInfhotelMDB & ".dbo.TCONTACTO.tCodigoCompania AND dbo.MPEDIDO.tContacto = " & sInfhotelMDB & ".dbo.TCONTACTO.tCodigoContacto INNER JOIN " & _
          "dbo.vEstadoPedido ON dbo.MPEDIDO.tEstadoPedido = dbo.vEstadoPedido.Codigo LEFT OUTER JOIN dbo.vProducto ON dbo.DPEDIDO.tCodigoProducto = dbo.vProducto.Codigo " & _
          "WHERE (dbo.MPEDIDO.tEstadoPedido <> '03') AND (dbo.DPEDIDO.tEstadoItem = 'N') and " & sCriterio & " " & _
          "GROUP BY dbo.DPEDIDO.tCodigoPedido, dbo.MPEDIDO.nAdulto, dbo.MPEDIDO.fRegistro, " & sInfhotelMDB & ".dbo.TCONTACTO.tCarnet, RTRIM(" & sInfhotelMDB & ".dbo.TCONTACTO.tPaterno) + ' ' + RTRIM(" & sInfhotelMDB & ".dbo.TCONTACTO.tMaterno) + ' ' + RTRIM(" & sInfhotelMDB & ".dbo.TCONTACTO.tNombre), " & _
          "dbo.vEstadoPedido.Descripcion, (case when dbo.MPEDIDO.tTipoPedido='03' then 'R/S' else (case when dbo.MPEDIDO.tTipoPedido='01' then cast(dbo.MPEDIDO.nAdulto as nvarchar(4)) else '1' end) end)"
    
   Set rsReporte = Lib.OpenRecordset(Isql, Cn)
   Reporte.DiscardSavedData
   Reporte.Database.SetDataSource rsReporte
   Reporte.Text13.SetText sRazonSocial
   Reporte.Text1.SetText localConectado
   Reporte.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
   frmEmite.CRViewer.ReportSource = Reporte
         
End Sub

