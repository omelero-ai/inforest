VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepPlanillaMovilidadMotorizado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planillas de Movilidad"
   ClientHeight    =   4305
   ClientLeft      =   4125
   ClientTop       =   1860
   ClientWidth     =   8730
   Icon            =   "frmRepPlanillaMovilidadMotorizado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Tarifa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2040
      TabIndex        =   19
      Top             =   2280
      Width           =   4080
      Begin VB.OptionButton optValor 
         Caption         =   "Todos"
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
         Left            =   360
         TabIndex        =   22
         Top             =   270
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optValor 
         Caption         =   "Oficial"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Top             =   270
         Width           =   975
      End
      Begin VB.OptionButton optValor 
         Caption         =   "Extra"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   20
         Top             =   270
         Width           =   855
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
      Left            =   1380
      Picture         =   "frmRepPlanillaMovilidadMotorizado.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
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
      Left            =   2835
      Picture         =   "frmRepPlanillaMovilidadMotorizado.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
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
      Left            =   5760
      Picture         =   "frmRepPlanillaMovilidadMotorizado.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
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
      Height          =   3540
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   8505
      Begin VB.OptionButton optOpcion 
         Caption         =   "General"
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
         Index           =   2
         Left            =   5160
         TabIndex        =   23
         Top             =   3120
         Width           =   1185
      End
      Begin VB.OptionButton optOpcion 
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
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   3120
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optOpcion 
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
         Height          =   240
         Index           =   1
         Left            =   3720
         TabIndex        =   10
         Top             =   3120
         Width           =   1425
      End
      Begin VB.TextBox txtMonto 
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
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1800
         Width           =   4095
      End
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
         Left            =   6120
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin MSDataListLib.DataCombo cboMotorizado 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   360
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
         Left            =   1920
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   4800
         TabIndex        =   17
         Top             =   840
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
         Format          =   63176707
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   4800
         TabIndex        =   18
         Top             =   1320
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
         Format          =   63176707
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Consulta :"
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
         Index           =   2
         Left            =   630
         TabIndex        =   16
         Top             =   3120
         Width           =   1230
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
         Left            =   780
         TabIndex        =   15
         Top             =   840
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
         Left            =   870
         TabIndex        =   14
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mayor al Monto :"
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
         Index           =   5
         Left            =   510
         TabIndex        =   13
         Top             =   1800
         Width           =   1350
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
         Left            =   840
         TabIndex        =   12
         Top             =   360
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
      Left            =   4290
      Picture         =   "frmRepPlanillaMovilidadMotorizado.frx":0B62
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
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
Attribute VB_Name = "frmRepPlanillaMovilidadMotorizado"
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
Dim ReporteD As New dsrPlanillaMovilidadPorDiaDetallado
Dim ReporteR As New dsrPlanillaMovilidadPorDiaResumido
Dim ReporteG As New dsrPlanillaMovilidadGeneral

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
            If optOpcion(2).Value Then
             Genera1
            Else
             Genera
            End If
               If rsReporte.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(0).Value = True Then
                  ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteD.PaperOrientation = crLandscape
                  
               ElseIf optOpcion(1).Value = True Then
                  ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteR.PaperOrientation = crLandscape
                  
               Else
                  ReporteG.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteG.PaperOrientation = crLandscape
               End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
            If optOpcion(2).Value Then
             Genera1
            Else
             Genera
            End If
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(0).Value = True Then
                  ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteD.PaperOrientation = crLandscape
                  ReporteD.PrintOut
                  
               ElseIf optOpcion(1).Value = True Then
                  ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteR.PaperOrientation = crLandscape
                  ReporteR.PrintOut
                  
               Else
                  ReporteG.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteG.PaperOrientation = crLandscape
                  ReporteG.PrintOut
                  
               End If
          
          Case Is = 3 ' Exportar
            If optOpcion(2).Value Then
             Genera1
            Else
             Genera
            End If
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               Screen.MousePointer = vbDefault
               'On Error GoTo Cancelar

               If optOpcion(0).Value = True Then
                  ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteD.PaperOrientation = crLandscape
                  
                  ReporteD.ExportOptions.FormatType = 21
                  ReporteD.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteD.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteD.Export False
                  
               ElseIf optOpcion(1).Value = True Then
                  ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteR.PaperOrientation = crPortrait
                  ReporteR.ExportOptions.FormatType = 21
                  ReporteR.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteR.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteR.Export False
                  
               Else
                  ReporteG.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReporteG.PaperOrientation = crPortrait
                  ReporteG.ExportOptions.FormatType = 21
                  ReporteG.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReporteG.ExportOptions.DiskFileName = cmdSave.FileName
                  ReporteG.Export False
               End If
         
          Case Is = 3 ' Salir
               Unload Me
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

Public Sub Genera()
    Dim oComando As clsComando
    Screen.MousePointer = vbHourglass
    
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_PlanillaMovilidad", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    If optValor(1).Value Then
       sTarifa = "O"
    ElseIf optValor(2).Value Then
       sTarifa = "E"
    Else
       sTarifa = ""
    End If
    
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 20, cboMotorizado.BoundText
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@nMonto", adDouble, adParamInput, 10, Val(txtMonto.Text)
    oComando.CreateParameter "@tTipo", adVarChar, adParamInput, 1, IIf(Me.optOpcion(0).Value, "D", "R")
                      
    oComando.CreateParameter "@tTarifa", adVarChar, adParamInput, 50, sTarifa
                     
                      
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
        
    Set rsReporte = oComando.GetSP()
    
        
    If optOpcion(0).Value Then
       ReporteD.DiscardSavedData
       ReporteD.Database.SetDataSource rsReporte
     
       If optValor(2).Value Then
          ReporteD.Text14.SetText "Extra"
       Else
          ReporteD.Text14.SetText ""
       End If
       ReporteD.Text23.SetText sRazonComercial
       ReporteD.Text22.SetText sRUC
       ReporteD.Text21.SetText sRazonSocial
       If Format(dtpFecIni.Value, "dd/MM/yyyy") = Format(dtpFecFin.Value, "dd/MM/yyyy") Then
        ReporteD.Text10.SetText Format(dtpFecIni.Value, "dd/MM/yyyy")
       Else
        ReporteD.Text10.SetText Format(dtpFecIni.Value, "dd/MM/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & "  al " & Format(dtpFecFin.Value, "dd/MM/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
       End If
       
       If Month(dtpFecIni.Value) = Month(dtpFecFin.Value) Then
       
       ReporteD.Text45.SetText MonthName(Month(dtpFecIni.Value))
       Else
       ReporteD.Text45.SetText MonthName(Month(dtpFecIni.Value)) & " - " & MonthName(Month(dtpFecFin.Value))
       End If

       frmEmite.CRViewer.ReportSource = ReporteD
    Else
       ReporteR.DiscardSavedData
       ReporteR.Database.SetDataSource rsReporte
       
       If optValor(2).Value Then
          ReporteR.Text14.SetText "Extra"
       Else
          ReporteR.Text14.SetText ""
       End If
       
       ReporteR.Text23.SetText sRazonComercial
       ReporteR.Text22.SetText sRUC
       ReporteR.Text21.SetText sRazonSocial
       ReporteR.Text9.SetText Format(dtpFecIni.Value, "dd/MM/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " al " & Format(dtpFecFin.Value, "dd/MM/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
       frmEmite.CRViewer.ReportSource = ReporteR
    End If
End Sub
 
 
Public Sub Genera1()
    Dim oComando As clsComando
    Screen.MousePointer = vbHourglass
    
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_PlanillaMovilidadGeneral", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
       
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 20, cboMotorizado.BoundText
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                      
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
        
    Set rsReporte = oComando.GetSP()
    
       ReporteG.DiscardSavedData
       ReporteG.Database.SetDataSource rsReporte
             
       ReporteG.Text23.SetText sRazonComercial
       ReporteG.Text22.SetText sRUC
       ReporteG.Text21.SetText sRazonSocial
       ReporteG.Text9.SetText Format(dtpFecIni.Value, "dd/MM/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " al " & Format(dtpFecFin.Value, "dd/MM/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
       frmEmite.CRViewer.ReportSource = ReporteG

End Sub


Private Sub optOpcion_Click(Index As Integer)

    If optOpcion(2).Value Then
       txtMonto.Enabled = False
       Frame4.Enabled = False
    Else
       txtMonto.Enabled = True
       Frame4.Enabled = True
    End If

End Sub

Private Sub txtMonto_Change()
   If txtMonto.Text = "" Or Not IsNumeric(txtMonto) Then
      txtMonto.Text = "0.00"
   End If
End Sub

Private Sub txtMonto_gotfocus()
   txtMonto.SelLength = Len(Trim(txtMonto.Text))
End Sub

Private Sub txtMonto_LostFocus()
   txtMonto.Text = Format(txtMonto.Text, "###,##0.00")
End Sub


