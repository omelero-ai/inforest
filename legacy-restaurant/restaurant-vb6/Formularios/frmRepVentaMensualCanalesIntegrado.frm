VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRepVentaMensualCanalesIntegrado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venta Mensual Por Canales"
   ClientHeight    =   5220
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepVentaMensualCanalesIntegrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
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
      Left            =   1620
      Picture         =   "frmRepVentaMensualCanalesIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
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
      Left            =   3105
      Picture         =   "frmRepVentaMensualCanalesIntegrado.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
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
      Left            =   7485
      Picture         =   "frmRepVentaMensualCanalesIntegrado.frx":132E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
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
      Left            =   4560
      Picture         =   "frmRepVentaMensualCanalesIntegrado.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton CmdActualiza 
      Caption         =   "Actualizar"
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
      Left            =   6015
      Picture         =   "frmRepVentaMensualCanalesIntegrado.frx":1C42
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      UseMaskColor    =   -1  'True
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
      Height          =   4485
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10515
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
         Height          =   855
         Left            =   2655
         TabIndex        =   17
         Top             =   1680
         Width           =   2295
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.ComboBox CmbMes 
         Height          =   360
         Left            =   2655
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   2265
      End
      Begin VB.Frame Frame3 
         Caption         =   "Servidores En Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4395
         Left            =   5880
         TabIndex        =   16
         Top             =   0
         Width           =   4575
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar Todo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2760
            TabIndex        =   8
            Top             =   3960
            Width           =   1695
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   3735
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   6588
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
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
         Height          =   855
         Left            =   2640
         TabIndex        =   15
         Top             =   2760
         Width           =   2295
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Local"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Value           =   -1  'True
            Width           =   2000
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Sector"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   2000
         End
      End
      Begin MSComCtl2.DTPicker dtpAnual 
         Height          =   405
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
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
         Format          =   58458115
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1200
         Width           =   2265
         _ExtentX        =   3995
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
         Format          =   58458115
         UpDown          =   -1  'True
         CurrentDate     =   38818
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
         TabIndex        =   22
         Top             =   315
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Hora de Cierre:"
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
         Left            =   1200
         TabIndex        =   21
         Top             =   1275
         Width           =   1305
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
         TabIndex        =   20
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Valorización :"
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
         Index           =   5
         Left            =   1320
         TabIndex        =   19
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Reporte :"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   3000
         Width           =   1245
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   0
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepVentaMensualCanalesIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTemporal As String
Dim ReporteMensual As New dsrVentaMensualCanalesIntegrada
Dim RsTemporal As Recordset
Dim RsLocal As Recordset
Dim RsReporte As Recordset
Dim sector(500) As String
Dim sCriterio As String
Dim meses(11) As String
Dim sPrecio As String
Dim sTexto As String

Sub LlenaCombos()
Dim i As Integer
   For i = 0 To 11
       CmbMes.AddItem meses(i)
   Next i
   CmbMes.ListIndex = Month(Date) - 1
End Sub

Public Sub activaCheck(Tipo As Boolean)
   Dim o As Integer
    For o = 1 To lstServidores.ListItems.Count
        If Tipo = True Then
            If lstServidores.ListItems.Item(o).SubItems(3) <> "Desconectado" Then

                lstServidores.ListItems.Item(o).Checked = Tipo
            End If
        Else
            lstServidores.ListItems.Item(o).Checked = Tipo
        End If
    Next o
     '   validaCantidadSeleccion
End Sub

Private Sub chkTodo_Click()
        activaCheck IIf(chkTodo.value = 1, True, False)
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Cn.Execute "delete from " & sTemporal
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
    
   Dim r, p As Integer
   r = 0
   p = 0
   For p = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(p).Checked = True Then
            r = r + 1
        End If
   Next p
   If r = 0 Then
      MsgBox "Debe seleccionar al menos un Item de la Lista de Servidores", vbCritical, sMensaje
   Exit Sub
   End If
   sCriterio = ""
                
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If RsReporte.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               frmEmite.CRViewer.DisplayGroupTree = False
          
               
               ReporteMensual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteMensual.PaperOrientation = crLandscape
          
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
                    
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If RsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
          
               ReporteMensual.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteMensual.PaperOrientation = crLandscape
               ReporteMensual.PrintOut
          
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
                
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
              
   End Select
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpAnual.value = Date
    sTemporal = dbTemporal(sCaja, 7, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "tipo", "nvarchar(50)", "ventas", "float", "pax", "float", "pedidos", "float")
    meses(0) = "Enero"
    meses(1) = "Febrero"
    meses(2) = "Marzo"
    meses(3) = "Abril"
    meses(4) = "Mayo"
    meses(5) = "Junio"
    meses(6) = "Julio"
    meses(7) = "Agosto"
    meses(8) = "Setiembre"
    meses(9) = "Octubre"
    meses(10) = "Noviembre"
    meses(11) = "Diciembre"
    LlenaCombos
    'configuraListaServidores
    obtieneListaServidores
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsTemporal = Nothing
   Set RsReporte = Nothing
   Cn.Execute "drop table " & sTemporal

End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   On Error GoTo sigue
   Dim oComando As clsComando
   Dim RsDatos As New ADODB.Recordset
   Dim conectado As Boolean
   Dim w As Integer
   Dim z As Long
   Dim cs As String
   Dim cl As String
   Dim sector As String
   Dim llocal As String
   
    Dim sFecha As String
    Dim sAno As String
    Dim sMes As String
    Dim sDia As String
    Dim j, i As Integer
    Set oComando = New clsComando
    j = 0
    If optValor(0).value Then
       sPrecio = "dbo.DPEDIDO.nVenta"
       sTexto = " Venta"
    Else
       sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
       sTexto = " Neto"
    End If
    
    sFecha = ""
    sAno = LTrim(str(Year(dtpAnual.value)))
    sMes = IIf(CmbMes.ListIndex + 1 > 9, LTrim(str(CmbMes.ListIndex + 1)), "0" & LTrim(str(CmbMes.ListIndex + 1)))
        
    Select Case CmbMes.ListIndex + 1
           Case 2
                If Year(dtpAnual.value) Mod 4 = 0 Then
                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/29')"
                Else
                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/28')"
                End If
                
           Case 1, 3, 5, 7, 8, 10, 12
                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/31')"
                   
           Case 4, 6, 9, 11
                   sFecha = "dbo.MPEDIDO.fRegistro >= DATEADD(HH," & Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/01') and dbo.MPEDIDO.fRegistro <= DATEADD(HH," & 24 + Hour(dtpHora) & ", '" & sAno & "/" & sMes & "/30')"

    End Select
    
    For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
                conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                If conectado = False Then
                    MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                    lstServidores.ListItems.Item(w).Checked = False
                    GoTo sigue
                End If
                 If Not oComando.CreateCmdSp("spRep_VentaMensualCanalesIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                       Set oComando = Nothing
                     Exit Sub
                 End If
                  oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 60, sPrecio
                  oComando.CreateParameter "@sAno", adVarChar, adParamInput, 4, sAno
                  oComando.CreateParameter "@sMes", adVarChar, adParamInput, 4, sMes
                  oComando.CreateParameter "@sFecha", adVarChar, adParamInput, 400, sFecha
                  oComando.CreateParameter "@dHour", adDouble, adParamInput, 8, Val(Hour(dtpHora))
                                  
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Sub
                End If
                    'cs = LTrim(Mid(lstServidores.ListItems.Item(w), 1, 2))
                    cl = lstServidores.ListItems.Item(w)
                    sector = lstServidores.ListItems.Item(w).SubItems(1)
                    llocal = lstServidores.ListItems.Item(w).SubItems(2)
                   Set RsDatos = oComando.GetSP()
                If Not (RsDatos.EOF Or RsDatos.BOF) Then
                            RsDatos.MoveFirst
                                For z = 0 To RsDatos.RecordCount - 1
                                        Isql = "insert into " & sTemporal & " values ( '" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "'," & RsDatos.Fields(1) & "," & RsDatos.Fields(2) & "," & RsDatos.Fields(3) & " )"
'                                                 Debug.Print Isql
                                        Cn.Execute Isql
                                    RsDatos.MoveNext
                                Next z
                         End If
                End If
sigue:
    Next w
    Dim Tipo As String
    If optOpcion(0).value = True Then ' sector
        Isql = "select sector,tipo + char(13)+'   Pax' + char(13)+'   Pedidos' AS TIPO , sum(ventas) as ventas,sum(pax) as pax,sum(pedidos) as pedidos from " & sTemporal & " group by sector,tipo + char(13)+'   Pax' + char(13)+'   Pedidos' order by 1"
        Tipo = "Sector"
    Else ' local
        Tipo = "Local"
        Isql = "select local as sector,tipo + char(13)+'   Pax' + char(13)+'   Pedidos' AS TIPO  ,sum(ventas) as ventas, sum(pax) as pax,sum(pedidos) as pedidos from " & sTemporal & " group by local,tipo + char(13)+'   Pax' + char(13)+'   Pedidos'  order by 1"
    End If
    
        Set RsReporte = Lib.OpenRecordset(Isql, Cn)
        ReporteMensual.DiscardSavedData
        ReporteMensual.Database.SetDataSource RsReporte
        ReporteMensual.PaperOrientation = crLandscape
        
        ReporteMensual.ReportTitle = CmbMes.Text & " / " & Year(Me.dtpAnual.value) & "- Hora de Cierre: " & Format(Me.dtpHora.value, "HH") & " horas"  'sTexto
        
        ReporteMensual.Text1.SetText "Agrupación por: " & Tipo & " (Montos Expresado en Soles)"
        ReporteMensual.Text20.SetText sRazonSocial
        frmEmite.CRViewer.ReportSource = ReporteMensual
    
End Sub
Private Sub configuraListaServidores()
 Dim rsS As New ADODB.Recordset
 Set rsS = devuelveServidoresConectados
 Me.lstServidores.Checkboxes = True
 lstServidores.ColumnHeaders.Clear
 Me.lstServidores.ListItems.Clear
    Dim Item As ListItem
    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
         .ColumnHeaders.Add , , " Local ", 1100
         .ColumnHeaders.Add , , " Estado ", 1100
         .ColumnHeaders.Add , , "Ip", 0
         .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
        
     Dim X As Integer
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(4)
        rsS.MoveNext
        Next X
            End With
End Sub

Private Sub lstServidores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
       End If
   ' validaCantidadSeleccion
End Sub
Public Sub validaCantidadSeleccion()
Dim p, w As Integer
Dim s As Integer
    p = 0
    s = 0
    limpiaSector
    For w = 0 To lstServidores.ListItems.Count - 1
        If lstServidores.ListItems.Item(w + 1).Checked = True Then
            p = p + 1
            If verificaSectorSeleccionado(lstServidores.ListItems(w + 1).SubItems(1)) Then
                sector(w) = lstServidores.ListItems(w + 1).SubItems(1)
            End If
        End If
    Next w
    For w = 0 To 500
        If sector(w) <> "" Then
            s = s + 1
        End If
    Next w
    If Me.optOpcion(1).value = True Then
        If p > 9 Then
            Me.cmdOpcion(0).Enabled = False
        Else
            Me.cmdOpcion(0).Enabled = True
        End If
    Else
        If s > 9 Then
            Me.cmdOpcion(0).Enabled = False
        Else
            Me.cmdOpcion(0).Enabled = True
        End If
    End If
End Sub
Public Function verificaSectorSeleccionado(sectors As String) As Boolean
    verificaSectorSeleccionado = True
    Dim q As Integer
        For q = 0 To 500
            If sector(q) = sectors Then
                verificaSectorSeleccionado = False
                Exit Function
            End If
        Next q
End Function

Public Sub limpiaSector()
    Dim q As Integer
        For q = 0 To 500
            sector(q) = ""
        Next q
End Sub
Private Sub CmdActualiza_Click()
If MsgBox("Esta operación puede tardar varios minutos" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
    Screen.MousePointer = vbHourglass
    configuraListaServidores
    Me.chkTodo.value = 0
    MsgBox "Lista De Servidores Actualizada", vbInformation, sMensaje
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub obtieneListaServidores()
 Dim rsS As New ADODB.Recordset
 Set rsS = rsListaServidores
 Me.lstServidores.Checkboxes = True
 lstServidores.ColumnHeaders.Clear
 Me.lstServidores.ListItems.Clear
    Dim Item As ListItem
    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
         .ColumnHeaders.Add , , " Local ", 1100
         .ColumnHeaders.Add , , " Estado ", 1100
         .ColumnHeaders.Add , , "Ip", 0
         .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
        
     Dim X As Integer
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(4)
        rsS.MoveNext
        Next X
            End With

End Sub
