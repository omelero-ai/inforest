VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmRepCtaCteIntegrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Corrientes"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   Icon            =   "frmRepCtaCteIntegrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
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
      Left            =   1590
      Picture         =   "frmRepCtaCteIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
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
      Left            =   3075
      Picture         =   "frmRepCtaCteIntegrado.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
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
      Left            =   7455
      Picture         =   "frmRepCtaCteIntegrado.frx":132E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
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
      Left            =   4530
      Picture         =   "frmRepCtaCteIntegrado.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
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
      Left            =   5985
      Picture         =   "frmRepCtaCteIntegrado.frx":1C42
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame4 
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
      Height          =   4840
      Left            =   5880
      TabIndex        =   20
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox chkTodo 
         Caption         =   "Seleccionar Todo"
         Height          =   360
         Left            =   2760
         TabIndex        =   9
         Top             =   4440
         Width           =   1695
      End
      Begin MSComctlLib.ListView lstServidores 
         Height          =   4215
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7435
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
   Begin VB.Frame Frame 
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
      Height          =   4840
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   5865
      Begin VB.Frame Frame5 
         Caption         =   "Agrupado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3120
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Cliente"
            Height          =   240
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   1515
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Local"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Sector"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   1440
         Left            =   840
         TabIndex        =   16
         Top             =   1440
         Width           =   2025
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Consolidado"
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   1545
         End
      End
      Begin MSMAPI.MAPIMessages MAPIM 
         Left            =   3000
         Top             =   3360
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession MAPIS 
         Left            =   3960
         Top             =   3360
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   0   'False
         LogonUI         =   0   'False
         NewSession      =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   780
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
         Format          =   56360961
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   300
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
         Format          =   56360961
         CurrentDate     =   37539.2083333333
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   2160
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   375
         Left            =   3555
         TabIndex        =   1
         Top             =   300
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   56360963
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   375
         Left            =   3540
         TabIndex        =   3
         Top             =   780
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   56360963
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
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
         Left            =   270
         TabIndex        =   18
         Top             =   375
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
         Left            =   375
         TabIndex        =   17
         Top             =   870
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRepCtaCteIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset
Dim RsCliente As Recordset
Dim RsTipo As Recordset
Dim RsSubTipo As Recordset
Dim ReportC As New dsrRepCtaCteIntegradoConsolidado
Dim ReportD As New dsrRepCtaCteIntegradoDetallado
Dim ReportXCD As New dsrRepCtaCteIntegradoDetalladoXC
Dim ReportXCC As New dsrRepCtaCteIntegradoConsolidadoXC
Dim sCriterio As String
Dim sEstadoPedido As String
Dim sTitulo As String
Dim fInicio As Date
Dim fFinal As Date

Private Sub chkTodo_Click()
 activaCheck IIf(chkTodo.value = 1, True, False)
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

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   Dim p, r As Integer
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
   If dtpFecIni.value + dtpHoraIni.value > dtpFecFin.value + dtpHoraFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
   
   sTitulo = ""

   
    sTitulo = sTitulo & "Rango del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
       
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
               If optOpcion(0).value = True And optTipo(2).value = False Then
                  ReportC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportC.PaperOrientation = crPortrait
              
               ElseIf optOpcion(2).value = True And optTipo(2).value = False Then
                  ReportD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportD.PaperOrientation = crPortrait
               ElseIf optOpcion(2).value = True And optTipo(2).value = True Then
                  ReportXCD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportXCD.PaperOrientation = crPortrait
               ElseIf optOpcion(0).value = True And optTipo(2).value = True Then
                  ReportXCC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportXCC.PaperOrientation = crPortrait
               End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Screen.MousePointer = vbDefault
                   Exit Sub
               End If
               If optOpcion(0).value = True Then
                  ReportC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportC.PaperOrientation = crPortrait
                  Screen.MousePointer = vbDefault
                  ReportC.PrintOut

               Else
                  ReportD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportD.PaperOrientation = crPortrait
                  Screen.MousePointer = vbDefault
                  ReportD.PrintOut
               End If
          

         
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               
               If optOpcion(0).value = True Then
                  ReportC.ExportOptions.FormatType = 21
                  ReportC.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReportC.ExportOptions.DiskFileName = cmdSave.FileName
                  ReportC.Export False
                  

               Else
                  ReportD.ExportOptions.FormatType = 21
                  ReportD.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReportD.ExportOptions.DiskFileName = cmdSave.FileName
                  ReportD.Export False
               End If
         
         Case Is = 2 ' Salir
              Unload Me
   End Select
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Centrar Me

    dtpFecIni.value = Date - 3
    dtpFecFin.value = Date + 3
 

    sCliente = ""
     obtieneListaServidores
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsCliente = Nothing
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
    On Error GoTo sigue
   Dim conectado As Boolean
   Dim oComando As clsComando
   Dim sOrden As String
   Dim RsDatos As New ADODB.Recordset
   Dim w, z As Long
   Dim cs As String
   Dim cl As String
   Dim sector As String
   Dim llocal As String
   Dim sNombre As String
   Set oComando = New clsComando
   Dim Tipo As String
   Dim Cabecera As String
   Dim tipopedido As String
   Dim Estado As String
    If optOpcion(0).value = True Then ' consolidado
        sNombre = dbTemporal(sCaja, 10, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "descripcion", "nvarchar(200)", "tIdentidad", "NVARCHAR(50)", "tCodigoPedido", "NVARCHAR(15)", "ffecha", "nvarchar(20)", "testadopedido", "nvarchar(50)", "tdocumento", "nvarchar(20)", "Nventa", "FLOAT")
        Tipo = "2"
    
    Else ' detallado
        sNombre = dbTemporal(sCaja, 12, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "descripcion", "nvarchar(200)", "tIdentidad", "NVARCHAR(50)", "tCodigoPedido", "NVARCHAR(15)", "ffecha", "nvarchar(20)", "testadopedido", "nvarchar(50)", "Nventa", "FLOAT", "producto", "nvarchar(200)", "ncantidad", "float", "tdocumento", "nvarchar(20)")
        Tipo = "0"
    End If
     
   fInicio = Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHoraIni.value, "HH:mm")
   fFinal = Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHoraFin.value, "HH:mm")
   
   
    For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
        
                conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                If conectado = False Then
                    MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                    lstServidores.ListItems.Item(w).Checked = False
                     GoTo sigue
                End If
                
                If Not oComando.CreateCmdSp("spRep_CtaCteIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                Exit Sub
                End If
                
                    oComando.CreateParameter "@flagTipo", adVarChar, adParamInput, 2, Tipo
                    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
                    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                      If Not oComando.GetParamOK Then
                            Set oComando = Nothing
                            Exit Sub
                      End If
                 
                    cl = lstServidores.ListItems.Item(w)
                    sector = lstServidores.ListItems.Item(w).SubItems(1)
                    llocal = lstServidores.ListItems.Item(w).SubItems(2)
                  Set RsDatos = oComando.GetSP()
                 
                 
                 
                 
                 
                 
                 If Not (RsDatos.EOF Or RsDatos.BOF) Then
                    RsDatos.MoveFirst
                        For z = 0 To RsDatos.RecordCount - 1
                            
                            If optOpcion(0).value = True Then ' consolidado
                              Isql = "insert into " & sNombre & " values ('" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "','" & RsDatos.Fields(5) & "','" & Format(RsDatos.Fields(6), "dd/mm/yyyy hh:mm") & "','" & RsDatos.Fields(9) & "','" & RsDatos.Fields(8) & "'," & RsDatos.Fields(10) & " )"
                              '  Debug.Print Isql
                            Else ' detallado
                                Isql = "insert into " & sNombre & " values ('" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "','" & RsDatos.Fields(5) & "','" & Format(RsDatos.Fields(6), "dd/mm/yyyy hh:mm") & "','" & RsDatos.Fields(8) & "'," & RsDatos.Fields(9) & ",'" & RsDatos.Fields(11) & "'," & RsDatos.Fields(12) & ",'" & RsDatos.Fields(13) & "' )"
                            End If
                         '   Debug.Print Isql
                                Cn.Execute Isql
                            RsDatos.MoveNext
                        Next z
                 End If
        End If
sigue:
    Next w
    
    
    
   If optOpcion(0).value = True Then ' consolidado
        If optTipo(0).value = True Then ' sector
          
            tipopedido = "Local"
            Isql = "select sector as local,descripcion, tidentidad, local as tcodigopedido , ffecha,testadopedido,tdocumento,nventa from " & sNombre & " order by descripcion asc, tcodigopedido desc"
        ElseIf optTipo(1).value = True Then  ' local
          
            tipopedido = "Pedido"
            Isql = "select local,descripcion, tidentidad,tcodigopedido ,ffecha,testadopedido,tdocumento,nventa from " & sNombre & " order by tcodigopedido desc, ffecha desc"
         ElseIf optTipo(2).value = True Then
                 tipopedido = "Pedido"
            Isql = "select local,descripcion, tidentidad,tcodigopedido ,ffecha,testadopedido,tdocumento,nventa from " & sNombre & " order by 2 asc,1 asc, tcodigopedido desc, ffecha desc"
        End If
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
    
        If optTipo(2).value = True Then
            ReportXCC.DiscardSavedData
            ReportXCC.Database.SetDataSource rsReporte
            ReportXCC.ReportTitle = "Tipo: Resumido. " & sTitulo
            ReportXCC.Text8.SetText sRazonSocial
            ReportXCC.Text9.SetText tipopedido
            frmEmite.CRViewer.DisplayGroupTree = False
            frmEmite.CRViewer.ReportSource = ReportXCC
        Else
        
            ReportC.DiscardSavedData
            ReportC.Database.SetDataSource rsReporte
            ReportC.ReportTitle = "Tipo: Resumido. " & sTitulo
            ReportC.Text8.SetText sRazonSocial
            ReportC.Text9.SetText tipopedido
            frmEmite.CRViewer.DisplayGroupTree = False
            frmEmite.CRViewer.ReportSource = ReportC
        End If
   
   Else ' detallado
   
        If optTipo(0).value = True Then ' sector
            Cabecera = "Sector"
            tipopedido = "Local"
                Isql = "select sector as local,descripcion, tidentidad, local as tcodigopedido , ffecha,testadopedido,nventa,producto,ncantidad, tdocumento from " & sNombre & " order by descripcion asc, tcodigopedido desc"
        ElseIf optTipo(1).value = True Then  ' local
            Cabecera = "Local"
            tipopedido = "Pedido"
                Isql = "select local,descripcion, tidentidad,tcodigopedido ,ffecha,testadopedido,nventa,producto,ncantidad, tdocumento from " & sNombre & " order by tcodigopedido desc, ffecha desc"
        ElseIf optTipo(2).value = True Then
            Cabecera = "Cliente"
            tipopedido = "Local"
                Isql = "select local,descripcion, tidentidad,tcodigopedido ,ffecha,testadopedido,nventa,producto,ncantidad, tdocumento from " & sNombre & " order by 1,tcodigopedido desc, ffecha desc"
        End If
        
      Set rsReporte = Lib.OpenRecordset(Isql, Cn)
        If optTipo(2).value = True Then
            ReportXCD.DiscardSavedData
            ReportXCD.Database.SetDataSource rsReporte
            ReportXCD.ReportTitle = "Tipo: Detallado. " & sTitulo
            ReportXCD.Text8.SetText sRazonSocial
            ReportXCD.Text9.SetText tipopedido
            frmEmite.CRViewer.DisplayGroupTree = False
            frmEmite.CRViewer.ReportSource = ReportXCD
        Else
        
            ReportD.DiscardSavedData
            ReportD.Database.SetDataSource rsReporte
            ReportD.ReportTitle = "Tipo: Detallado. " & sTitulo
            ReportD.Text8.SetText sRazonSocial
            ReportD.Text9.SetText tipopedido
            frmEmite.CRViewer.DisplayGroupTree = False
            frmEmite.CRViewer.ReportSource = ReportD
        End If
   End If
   'Cn.Execute "drop table " & sNombre
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
      '  validaCantidadSeleccion
End Sub



Private Sub lstServidores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   Dim p, w As Integer
    Dim s As Integer
    If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
    End If

End Sub
