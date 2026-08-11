VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepControlMotorizado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de motorizados"
   ClientHeight    =   2895
   ClientLeft      =   4125
   ClientTop       =   1860
   ClientWidth     =   6135
   Icon            =   "frmRepControlMotorizado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   180
      Picture         =   "frmRepControlMotorizado.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
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
      Left            =   1635
      Picture         =   "frmRepControlMotorizado.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
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
      Left            =   4560
      Picture         =   "frmRepControlMotorizado.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
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
      Height          =   1980
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5865
      Begin VB.CheckBox chkMotorizado 
         Caption         =   "Todos los Motorizados"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin MSDataListLib.DataCombo cboMotorizado 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54525953
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54525953
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   1080
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
         Format          =   54525955
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   1560
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
         Format          =   54525955
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorizado :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1020
      End
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
      Left            =   3090
      Picture         =   "frmRepControlMotorizado.frx":0B62
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   60
      Top             =   4410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepControlMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset
Dim RsMotorizado As Recordset
Dim RsPrinter As Recordset
Dim sNombre   As String
Dim sCriterio As String
Dim sTitulo   As String
Dim sTarifa   As String
Dim fInicio As Date
Dim fFinal As Date
Dim ReporteD As New dsrControlMotorizado


Sub LlenaCombos()
    With cboMotorizado
     Isql = "Select * from vMotorizado order by descripcion "
     Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
     Set .RowSource = RsMotorizado
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub


Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
 
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
      
   If chkMotorizado.Value = 0 Then
      If cboMotorizado.Text = "" Then
         MsgBox "Debe escoger el Motorizado", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
   fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")

   Select Case Index
          Case Is = 0 ' Preview
            Genera
               If rsReporte.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteD.PaperOrientation = crLandscape
               
               frmEmite.CRViewer.Zoom 100
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
            Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteD.PaperOrientation = crLandscape
               ReporteD.PrintOut
          
          Case Is = 3 ' Exportar
            Genera Index
            
            Screen.MousePointer = vbDefault

            If rsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            rsReporte.MoveFirst
            ExportaExcel rsReporte
            Screen.MousePointer = vbDefault
            'On Error GoTo Cancelar

'            ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
'            ReporteD.PaperOrientation = crLandscape
'
'            ReporteD.ExportOptions.FormatType = 21
'            ReporteD.ExportOptions.DestinationType = 1
'            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
'            cmdSave.ShowSave
'            If cmdSave.FileName = "" Then
'               Exit Sub
'            End If
'            ReporteD.ExportOptions.DiskFileName = cmdSave.FileName
'            ReporteD.Export False
         
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

Private Sub Form_Unload(Cancel As Integer)
 

   Set rsReporte = Nothing
 
   Set RsMotorizado = Nothing
   Set frmRepPlanillaMovilidadMotorizado = Nothing
End Sub

Private Sub chkMotorizado_Click()
If chkMotorizado.Value = 1 Then
   cboMotorizado.Enabled = False
   cboMotorizado.Text = ""
Else
   cboMotorizado.Enabled = True
End If
End Sub

Public Sub Genera(Optional Index As Integer)
    Dim oComando As clsComando
    Screen.MousePointer = vbHourglass
    
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_ControlMotorizado", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    sTarifa = ""
    
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 20, cboMotorizado.BoundText
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
             
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
        
    Set rsReporte = oComando.GetSP()
    If Index = 3 Then
        Exit Sub
    End If
    
      ReporteD.DiscardSavedData
      ReporteD.Database.SetDataSource rsReporte
    
     'ReporteD.Text14.SetText ""
     
      ReporteD.Text23.SetText sRazonComercial
      ReporteD.Text22.SetText sRUC
      ReporteD.Text21.SetText sRazonSocial
      If Format(dtpFecIni.Value, "dd/MM/yyyy") = Format(dtpFecFin.Value, "dd/MM/yyyy") Then
       ReporteD.Text10.SetText Format(dtpFecIni.Value, "dd/MM/yyyy")
      Else
       ReporteD.Text10.SetText Format(dtpFecIni.Value, "dd/MM/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & "  al " & Format(dtpFecFin.Value, "dd/MM/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
      End If
      frmEmite.CRViewer.ReportSource = ReporteD
      
End Sub
 
 
'Public Sub Genera1()
'    Dim oComando As clsComando
'    Screen.MousePointer = vbHourglass
'
'    Set oComando = New clsComando
'    If Not oComando.CreateCmdSp("spRep_PlanillaMovilidadGeneral", Cn) Then
'       Set oComando = Nothing
'       Exit Sub
'    End If
'
'    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 20, cboMotorizado.BoundText
'    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
'    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
'
'    If Not oComando.GetParamOK Then
'       Set oComando = Nothing
'       Exit Sub
'    End If
'
'    Set rsReporte = oComando.GetSP()
'
'       ReporteG.DiscardSavedData
'       ReporteG.Database.SetDataSource rsReporte
'
'       ReporteG.Text23.SetText sRazonComercial
'       ReporteG.Text22.SetText sRUC
'       ReporteG.Text21.SetText sRazonSocial
'       ReporteG.Text9.SetText Format(dtpFecIni.Value, "dd/MM/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " al " & Format(dtpFecFin.Value, "dd/MM/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
'       frmEmite.CRViewer.ReportSource = ReporteG
'
'End Sub




