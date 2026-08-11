VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepTiempoKDS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiempos Pedidos KDS"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   3765
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   7245
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Reporte "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   5295
         Begin VB.OptionButton optOpcion 
            Caption         =   "Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   3600
            TabIndex        =   5
            Top             =   240
            Width           =   1275
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Producto"
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
            Index           =   0
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1545
         End
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3030
         Width           =   2445
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
         Left            =   4200
         Picture         =   "frmRepTiempoKDS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3000
         Width           =   615
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
         Left            =   4920
         TabIndex        =   11
         Top             =   3045
         Value           =   1  'Checked
         Width           =   2070
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
         Left            =   4920
         TabIndex        =   9
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2175
      End
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
         Left            =   4920
         TabIndex        =   7
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   825
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   63963137
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   63963137
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3765
         TabIndex        =   1
         Top             =   360
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3765
         TabIndex        =   3
         Top             =   825
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   2040
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   2520
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
         Left            =   600
         TabIndex        =   22
         Top             =   3120
         Width           =   945
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
         Left            =   480
         TabIndex        =   20
         Top             =   2520
         Width           =   975
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
         Left            =   840
         TabIndex        =   19
         Top             =   2040
         Width           =   645
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
         Left            =   450
         TabIndex        =   18
         Top             =   915
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
         Left            =   345
         TabIndex        =   17
         Top             =   450
         Width           =   1230
      End
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
      Left            =   1500
      Picture         =   "frmRepTiempoKDS.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3855
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
      Left            =   2940
      Picture         =   "frmRepTiempoKDS.frx":0634
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3855
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
      Left            =   5820
      Picture         =   "frmRepTiempoKDS.frx":0B66
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3855
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
      Left            =   4380
      Picture         =   "frmRepTiempoKDS.frx":0C58
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3855
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   5520
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepTiempoKDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsReporte As Recordset
Dim RsGrupo As ADODB.Recordset
Dim RsSubgrupo As ADODB.Recordset


Dim Resumido As New dsrTiempoKdsR
Dim Pedido As New dsrTiempoKdsP

Dim sCriterio As String
Dim sCondicion As String

Dim sProducto As String

Dim fInicio As Date
Dim fFinal As Date

Private Sub chkGrupo_Click()
   If chkGrupo.Value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
   Else
      cboGrupo.Enabled = True
   End If
End Sub

Private Sub chkproducto_Click()
   If chkProducto.Value = 1 Then
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


Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   
   If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If

   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
               If optOpcion(0).Value = True Then
               Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               Resumido.PaperOrientation = crPortrait
               Else
               Pedido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               Pedido.PaperOrientation = crPortrait
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
               If optOpcion(0).Value = True Then
               Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               Resumido.PaperOrientation = crPortrait
               Resumido.PrintOut
               Else
               Pedido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               Pedido.PaperOrientation = crPortrait
               Pedido.PrintOut
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
               
             If optOpcion(0).Value = True Then
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave

                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
              Else
                  Pedido.ExportOptions.FormatType = 21
                  Pedido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Pedido.ExportOptions.DiskFileName = cmdSave.FileName
                  Pedido.Export False
               End If
               
   End Select

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
   End Select

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
    dtpFecFin.Value = Date
End Sub

Sub LlenaCombos()

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


Public Sub Genera()
   Dim oComando As clsComando
   sCriterio = "Del " & dtpFecIni.Value & " al " & dtpFecFin
   sCondicion = ""
   If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger un Grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCondicion = sCondicion & IIf(sCondicion = "", "", " and ") & " tGrupo ='" & cboGrupo.BoundText & "'"
   End If
   
   If chkSubGrupo.Value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger un Sub Grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sCondicion = sCondicion & IIf(sCondicion = "", "", " and ") & " tSubGrupo ='" & cboSubGrupo.BoundText & "'"
   End If
   
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    
If optOpcion(0).Value = True Then
    If Not oComando.CreateCmdSp("USP_KDS_ResporteTiempoProducto", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@tGrupo", adVarChar, adParamInput, 10, Me.cboGrupo.BoundText
    oComando.CreateParameter "@tSubGrupo", adVarChar, adParamInput, 10, Me.cboSubGrupo.BoundText
    oComando.CreateParameter "@tProducto", adVarChar, adParamInput, 30, sProducto
Else
    If Not oComando.CreateCmdSp("USP_KDS_ResporteTiempoPedido", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
End If
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
     'rsReporte.Filter = sCondicion
    
    If optOpcion(0).Value = True Then
       Resumido.DiscardSavedData
       Resumido.Database.SetDataSource rsReporte
       Resumido.ReportTitle = sCriterio
       Resumido.Text12.SetText "Tiempo de Productos KDS"
       Resumido.Text15.SetText sRazonSocial
       'Resumido.Text1.SetText localConectado
       frmEmite.CRViewer.ReportSource = Resumido
     Else
       Pedido.DiscardSavedData
       Pedido.Database.SetDataSource rsReporte
       Pedido.ReportTitle = sCriterio
       Pedido.Text12.SetText "Tiempo de Pedidos KDS"
       Pedido.Text15.SetText sRazonSocial
       'Pedido.Text1.SetText localConectado
       frmEmite.CRViewer.ReportSource = Pedido
    End If

   Screen.MousePointer = vbDefault
End Sub

Private Sub optOpcion_Click(Index As Integer)
   If Index = 7 Then
      chkGrupo.Enabled = False
      chkSubGrupo.Enabled = False
      chkProducto.Enabled = False
   Else
      chkGrupo.Enabled = True
      chkSubGrupo.Enabled = True
      chkProducto.Enabled = True
   End If
End Sub

