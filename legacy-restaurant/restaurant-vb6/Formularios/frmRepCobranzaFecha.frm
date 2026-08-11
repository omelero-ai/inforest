VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRepCobranzaFecha 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobranza Mensual por Fechas"
   ClientHeight    =   4770
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepCobranzaFecha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6015
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
      Left            =   3060
      Picture         =   "frmRepCobranzaFecha.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4125
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
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
      Height          =   4080
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5970
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2640
         TabIndex        =   19
         Top             =   3000
         Width           =   2490
         Begin VB.OptionButton optTipo 
            Caption         =   "Hoja A4"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   2265
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Hoja Extendida"
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   11
            Top             =   600
            Width           =   2310
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2655
         TabIndex        =   16
         Top             =   1620
         Width           =   2490
         Begin VB.OptionButton optValor 
            Caption         =   "Fecha del Documento"
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   9
            Top             =   945
            Width           =   2310
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Fecha del Pedido"
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   8
            Top             =   585
            Width           =   1950
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Fecha Cancelación"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   2265
         End
      End
      Begin VB.ComboBox CmbMes 
         Height          =   360
         Left            =   2655
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   746
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker dtpAnual 
         Height          =   405
         Left            =   2655
         TabIndex        =   4
         Top             =   240
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   714
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
         CustomFormat    =   "yyyy"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   375
         Left            =   2655
         TabIndex        =   6
         Top             =   1200
         Width           =   2490
         _ExtentX        =   4392
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
         CustomFormat    =   "HH"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   38818
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo De Reporte :"
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
         Index           =   4
         Left            =   960
         TabIndex        =   18
         Top             =   3120
         Width           =   1545
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Origen del Pago :"
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
         Index           =   3
         Left            =   990
         TabIndex        =   17
         Top             =   1755
         Width           =   1500
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
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
         Index           =   2
         Left            =   2010
         TabIndex        =   15
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora de Cierre y Apertura  :"
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
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   1275
         Width           =   2355
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
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
         Index           =   0
         Left            =   2025
         TabIndex        =   13
         Top             =   315
         Width           =   465
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
      Left            =   4515
      Picture         =   "frmRepCobranzaFecha.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4125
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
      Left            =   1605
      Picture         =   "frmRepCobranzaFecha.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4125
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
      Left            =   150
      Picture         =   "frmRepCobranzaFecha.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4125
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepCobranzaFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTemporal As String
Dim ReporteMensual As New dsrCobranzaFecha
Dim ReporteHorizontal As New dsrCobranzaMensualHorizontal
Dim RsTemporal As Recordset
Dim RsLocal As Recordset
Dim rsReporte As Recordset
Dim sCriterio As String
Dim meses(11) As String
Dim sPrefijo As String
Dim sBases As String
Dim sTexto As String

Sub LlenaCombos()
Dim i As Integer
   For i = 0 To 11
       CmbMes.AddItem meses(i)
   Next i
   CmbMes.ListIndex = Month(Date) - 1
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
                
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optTipo(0).Value = True Then
                frmEmite.CRViewer.DisplayGroupTree = False
                ReporteMensual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteMensual.PaperOrientation = crLandscape
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
               Else
                frmEmite.CRViewer.DisplayGroupTree = False
                ReporteHorizontal.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteHorizontal.PaperOrientation = crLandscape
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
               End If
                    
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optTipo(0).Value = True Then
                ReporteMensual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteMensual.PaperOrientation = crLandscape
                ReporteMensual.PrintOut
               Else
                ReporteHorizontal.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteHorizontal.PaperOrientation = crLandscape
                ReporteHorizontal.PrintOut
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
                If optTipo(0).Value = True Then
                    Screen.MousePointer = vbDefault
                    ReporteMensual.ExportOptions.FormatType = 21
                    ReporteMensual.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                       Exit Sub
                    End If
                    ReporteMensual.ExportOptions.DiskFileName = cmdSave.FileName
                    ReporteMensual.Export False
                Else
                    Screen.MousePointer = vbDefault
                    ReporteHorizontal.ExportOptions.FormatType = 21
                    ReporteHorizontal.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                       Exit Sub
                    End If
                    ReporteHorizontal.ExportOptions.DiskFileName = cmdSave.FileName
                    ReporteHorizontal.Export False
                
                End If
   End Select
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpAnual.Value = Date
   
    meses(0) = "Enero"
    meses(1) = "Febrero"
    meses(2) = "Marzo"
    meses(3) = "Abril"
    meses(4) = "Mayo"
    meses(5) = "Junio"
    meses(6) = "Julio"
    meses(7) = "Agosto"
    meses(8) = "setiembre"
    meses(9) = "Octubre"
    meses(10) = "Noviembre"
    meses(11) = "Diciembre"
    LlenaCombos
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Cn.Execute "drop table " & sTemporal
   Set RsTemporal = Nothing
   Set rsReporte = Nothing
End Sub

Public Sub Genera()
    Screen.MousePointer = vbHourglass
    Dim oComando As clsComando
    Dim sFecha As String
    Dim sFechaDocumento As String
    Dim sFechaPedido As String
    Dim sAno As String
    Dim sMes As String
    Dim sDia As String
    Dim j, i As Integer
    j = 0
    
    If optValor(0).Value Then
       sPrefijo = "dbo.DPAGODOCUMENTO."
       sTexto = " Basado en la Fecha de Cancelación"
    ElseIf optValor(1).Value Then
       sPrefijo = "dbo.DPAGODOCUMENTO."
       sTexto = " Basado en la Fecha del Pedido"
    Else
       sPrefijo = "dbo.MDOCUMENTO."
       sTexto = " Basado en la Fecha del Documento"
    End If
    
    sFecha = ""
    sAno = LTrim(str(Year(dtpAnual.Value)))
    sMes = IIf(CmbMes.ListIndex + 1 > 9, LTrim(str(CmbMes.ListIndex + 1)), "0" & LTrim(str(CmbMes.ListIndex + 1)))
        
    Select Case CmbMes.ListIndex + 1
           Case 2
                If Year(dtpAnual.Value) Mod 4 = 0 Then
                   sFecha = sPrefijo & "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and " & sPrefijo & "fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                   sFechaDocumento = "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                   sFechaPedido = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                Else
                   sFecha = sPrefijo & "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and " & sPrefijo & "fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                   sFechaDocumento = "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                   sFechaPedido = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                End If
                
           Case 1, 3, 5, 7, 8, 10, 12
                   sFecha = sPrefijo & "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and " & sPrefijo & "fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
                   sFechaDocumento = "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
                   sFechaPedido = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
                   
           Case 4, 6, 9, 11
                   sFecha = sPrefijo & "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and " & sPrefijo & "fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
                   sFechaDocumento = "fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
                   sFechaPedido = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"
    End Select
        
     Set oComando = New clsComando
     If Not oComando.CreateCmdSp("spRep_CobranzaFecha", Cn) Then
      Set oComando = Nothing
      Exit Sub
     End If
     oComando.CreateParameter "@flagTIPOREPORTE", adBoolean, adParamInput, 1, optTipo(0).Value
     oComando.CreateParameter "@flagValor", adBoolean, adParamInput, 1, optValor(1).Value
     oComando.CreateParameter "@sPrefijo", adVarChar, adParamInput, 40, sPrefijo
     oComando.CreateParameter "@sAno", adVarChar, adParamInput, 4, sAno
     oComando.CreateParameter "@sMes", adVarChar, adParamInput, 4, sMes
     oComando.CreateParameter "@sFechaPedido", adVarChar, adParamInput, 400, sFechaPedido
     oComando.CreateParameter "@sFechaDocumento", adVarChar, adParamInput, 400, sFechaDocumento
     oComando.CreateParameter "@sFecha", adVarChar, adParamInput, 400, sFecha
     oComando.CreateParameter "@dHour", adDouble, adParamInput, 8, Val(Hour(dtpHora))
                     
   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If

   Set rsReporte = oComando.GetSP()
   If Me.optTipo(0).Value = True Then
    ReporteMensual.DiscardSavedData
    ReporteMensual.Database.SetDataSource rsReporte
    ReporteMensual.ReportTitle = "Mes de : " & CmbMes.Text & sTexto
    Dim xTar As String
    xTar = Calcular("select tResumido as Codigo from TTARJETACREDITO where tCodigoTarjeta='01'", Cn)
    ReporteMensual.Text7.SetText IIf(xTar = "0", "no/defin", Mid(xTar, 1, 10))
    xTar = Calcular("select tResumido as Codigo from TTARJETACREDITO where tCodigoTarjeta='02'", Cn)
    ReporteMensual.Text8.SetText IIf(xTar = "0", "no/defin", Mid(xTar, 1, 10))
    xTar = Calcular("select tResumido as Codigo from TTARJETACREDITO where tCodigoTarjeta='03'", Cn)
    ReporteMensual.Text9.SetText IIf(xTar = "0", "no/defin", Mid(xTar, 1, 10))
    xTar = Calcular("select tResumido as Codigo from TTARJETACREDITO where tCodigoTarjeta='04'", Cn)
    ReporteMensual.Text10.SetText IIf(xTar = "0", "no/defin", Mid(xTar, 1, 10))
    ReporteMensual.Text20.SetText sRazonSocial
    ReporteMensual.Text1.SetText localConectado
    frmEmite.CRViewer.ReportSource = ReporteMensual
    Else
    
    ReporteHorizontal.DiscardSavedData
    ReporteHorizontal.Text20.SetText sRazonSocial
    ReporteHorizontal.Text1.SetText localConectado
    ReporteHorizontal.Database.SetDataSource rsReporte
    ReporteHorizontal.ReportTitle = "Mes de : " & CmbMes.Text & sTexto
    frmEmite.CRViewer.ReportSource = ReporteHorizontal
    
    
    End If
End Sub
