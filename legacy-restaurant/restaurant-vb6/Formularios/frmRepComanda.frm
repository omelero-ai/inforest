VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepComanda 
   Caption         =   "Reporte de Comandas"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   Icon            =   "frmRepComanda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
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
      Left            =   3210
      Picture         =   "frmRepComanda.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3330
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1860
      Picture         =   "frmRepComanda.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3330
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   30
      TabIndex        =   15
      Top             =   0
      Width           =   5955
      Begin MSDataListLib.DataCombo cboUsuario 
         Height          =   360
         Left            =   1515
         TabIndex        =   2
         Top             =   780
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboMozo 
         Height          =   360
         Left            =   1515
         TabIndex        =   0
         Top             =   330
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton OptResumen 
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1530
         TabIndex        =   10
         Top             =   2880
         Width           =   1260
      End
      Begin VB.CheckBox chkTodoMozo 
         Caption         =   "Todos los Meseros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3870
         TabIndex        =   1
         Top             =   382
         Width           =   1965
      End
      Begin VB.CheckBox chkTodoUsuario 
         Caption         =   "Todos los Usuarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3870
         TabIndex        =   3
         Top             =   772
         Width           =   1905
      End
      Begin VB.ComboBox cboOrden 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmRepComanda.frx":0D60
         Left            =   1515
         List            =   "frmRepComanda.frx":0D6A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2010
         Width           =   2175
      End
      Begin VB.OptionButton OptDetalle 
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1530
         TabIndex        =   9
         Top             =   2490
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   315
         Left            =   1515
         TabIndex        =   6
         Top             =   1590
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   315
         Left            =   1515
         TabIndex        =   4
         Top             =   1170
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   330
         Left            =   3870
         TabIndex        =   5
         Top             =   1170
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   330
         Left            =   3870
         TabIndex        =   7
         Top             =   1590
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   5340
         Top             =   2610
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   660
         TabIndex        =   21
         Top             =   795
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   690
         TabIndex        =   20
         Top             =   390
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenado :"
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
         Left            =   480
         TabIndex        =   19
         Top             =   2070
         Width           =   960
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
         Index           =   0
         Left            =   210
         TabIndex        =   18
         Top             =   1200
         Width           =   1230
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
         Left            =   315
         TabIndex        =   17
         Top             =   1650
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reporte :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   690
         TabIndex        =   16
         Top             =   2505
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdOpcion 
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
      Height          =   615
      Index           =   1
      Left            =   4680
      Picture         =   "frmRepComanda.frx":0D7F
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3330
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Emite"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   495
      Picture         =   "frmRepComanda.frx":0E71
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3330
      Width           =   1350
   End
End
Attribute VB_Name = "frmRepComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsReporte As Recordset
Dim RsUser    As Recordset
Dim RsMozo    As Recordset
Dim sTitulo   As String
Dim sCriterio As String
Dim sOrden    As String
Dim RsEmpresa As Recordset
Dim RsDocumento As Recordset

Dim CrtDetalleC As New dsrComandaD
Dim CrtResumenC As New dsrComandaR

Dim sReporte As String

Dim fInicio As Date
Dim fFinal As Date

Private Sub ChkTodomozo_Click()
    If chkTodoMozo.Value = 0 Then
       cboMozo.Enabled = True
    Else
       cboMozo.Text = ""
       cboMozo.Enabled = False
    End If
End Sub

Private Sub ChkTodousuario_Click()
    If chkTodoUsuario.Value = 0 Then
       cboUsuario.Enabled = True
    Else
       cboUsuario.Text = ""
       cboUsuario.Enabled = False
    End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

 sCriterio = ""
 Screen.MousePointer = vbHourglass

 sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " al " & Format(dtpFecFin.Value, "dd/mmm/yyyy")
    
   If chkTodoMozo.Value = 0 Then
      If cboMozo.Text = "" Then
         MsgBox "Debe escoger un Mesero", vbCritical, sMensaje
        Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " Mozo = '" & Me.cboMozo.BoundText & "'"
   End If
   If chkTodoUsuario.Value = 0 Then
      If cboUsuario.Text = "" Then
         MsgBox "Debe escoger un Usuario", vbCritical, sMensaje
        Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "usuario ='" & Me.cboUsuario.BoundText & "'"
   End If
   
   Select Case Index
          Case Is = 0 'Emite
         'Preliminar
          If OptDetalle.Value = True Then
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               frmEmite.CRViewer.DisplayGroupTree = False
               CrtDetalleC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               CrtDetalleC.PaperOrientation = crLandscape
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
               
          ElseIf OptResumen.Value = True Then
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               frmEmite.CRViewer.DisplayGroupTree = False
               CrtResumenC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               CrtResumenC.PaperOrientation = crPortrait
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          End If
          
    Case Is = 2 'Impresora
       
       'Print
       If OptDetalle.Value = True Then
          Genera
          Screen.MousePointer = vbDefault
          If rsReporte.EOF = True Then
             MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
             Exit Sub
          End If
          CrtDetalleC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
          CrtDetalleC.PaperOrientation = crLandscape
          CrtDetalleC.PrintOut
       
       ElseIf OptResumen.Value = True Then
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
              MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
              Exit Sub
            End If
            CrtResumenC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            CrtResumenC.PaperOrientation = crPortrait
            CrtResumenC.PrintOut
       End If
    
    Case Is = 3 'Excel
         If OptDetalle.Value = True Then
           Genera
           Screen.MousePointer = vbDefault
           If rsReporte.EOF = True Then
              MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
              Exit Sub
           End If
           CrtDetalleC.ExportOptions.FormatType = 21
           CrtDetalleC.ExportOptions.DestinationType = 1
           cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
           cmdSave.ShowSave
           If cmdSave.FileName = "" Then
              Exit Sub
           End If

           CrtDetalleC.ExportOptions.DiskFileName = cmdSave.FileName
           CrtDetalleC.Export False
         
         ElseIf OptResumen.Value = True Then
             Genera
             Screen.MousePointer = vbDefault
             If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
             End If
             CrtResumenC.ExportOptions.FormatType = 21
             CrtResumenC.ExportOptions.DestinationType = 1
             cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
             cmdSave.ShowSave
             If cmdSave.FileName = "" Then
                Exit Sub
             End If
             CrtResumenC.ExportOptions.DiskFileName = cmdSave.FileName
             CrtResumenC.Export False
         End If
    
   Case Is = 1
        Screen.MousePointer = vbDefault
        Unload Me
   End Select
End Sub
Sub LlenaCombos()

    With cboMozo
         Isql = "Select * from vMozo where substring(Codigo,1,1)<>'*'"
         Set RsMozo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMozo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "tResumido"
    End With
  
    With cboUsuario
      Isql = "Select * from tUsuario where lActivo = 1 order by tCodigoUsuario "
      Set RsUser = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsUser
          .DataField = "tDetallado"
          .ListField = "tDetallado"
          .BoundColumn = "tResumido"
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
  dtpFecIni.Value = Date
  dtpFecFin.Value = Date
   
  chkTodoMozo.Value = 1
  cboMozo.Enabled = False
 
  chkTodoUsuario.Value = 1
  cboUsuario.Enabled = False
  OptDetalle.Value = True
  Call LlenaCombos
  cboOrden.Text = "Pedido"
  
End Sub
Public Sub Genera()
   Dim oComando As clsComando
    fInicio = Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHoraIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHoraFin.Value, "HH:mm")
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_Comanda", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, OptDetalle.Value
    oComando.CreateParameter "@SOrden", adVarChar, adParamInput, 50, cboOrden.Text
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@sCriterio", adVarChar, adParamInput, 400, sCriterio
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    'rsReporte.Filter = "Mozo='RPALOMINO'"
   If OptDetalle.Value = True Then
        CrtDetalleC.DiscardSavedData
        CrtDetalleC.Database.SetDataSource rsReporte
        CrtDetalleC.ReportTitle = sTitulo
        CrtDetalleC.Text13.SetText sRazonSocial
        CrtDetalleC.Text5.SetText localConectado
        frmEmite.CRViewer.ReportSource = CrtDetalleC
   Else
        CrtResumenC.DiscardSavedData
        CrtResumenC.Database.SetDataSource rsReporte
        CrtResumenC.ReportTitle = sTitulo
        CrtResumenC.Text13.SetText sRazonSocial
        CrtResumenC.Text3.SetText localConectado
        frmEmite.CRViewer.ReportSource = CrtResumenC
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
End Sub

