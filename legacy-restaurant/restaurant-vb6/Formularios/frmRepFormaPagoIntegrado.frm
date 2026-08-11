VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepFormaPagoIntegrado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consolidado De Liquidaciones de Caja"
   ClientHeight    =   5700
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   12405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepFormaPagoIntegrado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   480
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   2580
      Picture         =   "frmRepFormaPagoIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
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
      Left            =   4035
      Picture         =   "frmRepFormaPagoIntegrado.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
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
      Left            =   6930
      Picture         =   "frmRepFormaPagoIntegrado.frx":132E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      UseMaskColor    =   -1  'True
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
      Index           =   2
      Left            =   5490
      Picture         =   "frmRepFormaPagoIntegrado.frx":13F6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
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
      Height          =   4950
      Left            =   45
      TabIndex        =   17
      Top             =   45
      Width           =   12330
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
         Height          =   1455
         Left            =   2280
         TabIndex        =   23
         Top             =   1560
         Width           =   3375
         Begin VB.CheckBox chkForma 
            Caption         =   "Propina"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   1000
            Value           =   1  'Checked
            Width           =   3000
         End
         Begin VB.CheckBox chkForma 
            Caption         =   "Crédito"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkForma 
            Caption         =   "Contado"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   200
            Value           =   1  'Checked
            Width           =   2415
         End
      End
      Begin VB.CheckBox chkFHoraria 
         Caption         =   "Franja Horaria"
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
         Left            =   5760
         TabIndex        =   11
         Top             =   840
         Width           =   1695
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
         Height          =   4695
         Left            =   7680
         TabIndex        =   21
         Top             =   120
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
            TabIndex        =   6
            Top             =   4200
            Width           =   1695
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   3975
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7011
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
      Begin VB.Frame Frame2 
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
         Left            =   2280
         TabIndex        =   20
         Top             =   3120
         Width           =   3375
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Local"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Value           =   -1  'True
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Sector"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   200
            Width           =   3015
         End
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   2340
         TabIndex        =   9
         Top             =   1125
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   52232193
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   2340
         TabIndex        =   7
         Top             =   600
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   52232193
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   4320
         TabIndex        =   8
         Top             =   600
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   52232195
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   4320
         TabIndex        =   10
         Top             =   1125
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   52232195
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Reporte :"
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
         Index           =   3
         Left            =   690
         TabIndex        =   24
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Ingreso :"
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
         Left            =   870
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   1080
         TabIndex        =   19
         Top             =   1305
         Width           =   1125
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Left            =   975
         TabIndex        =   18
         Top             =   780
         Width           =   1230
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
      Index           =   3
      Left            =   8370
      Picture         =   "frmRepFormaPagoIntegrado.frx":1C18
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepFormaPagoIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCliente As Recordset
Dim RsEstado As Recordset
Dim RsTipoDocumento As Recordset
Dim Consolidado As New dsrFormaPagoDiarioIntegradoLocales
Dim ConsolidadoSector As New dsrFormaPagoDiarioIntegradoSector
Dim Identidad As String
Dim RsPrinter As Recordset
Dim sector(500) As String
Dim sReporte As String
Dim sNombre As String
Dim fInicio As Date
Dim fFinal As Date

Private Sub chkTodo_Click()
        activaCheck IIf(chkTodo.value = 1, True, False)
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   
   Cn.Execute " delete from " & sNombre
 '  Cn.Execute " delete from " & sReporte
   If Index = 3 Then
      Unload Me
      Exit Sub
   End If
   Dim p, r As Integer
    r = 0
   p = 0
   
   If chkForma(0).value = 0 And chkForma(1).value = 0 And chkForma(2).value = 0 Then
    MsgBox "Debe seleccionar una Forma de Pago"
   Exit Sub
   End If
   For p = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(p).Checked = True Then
            r = r + 1
        End If
   Next p
   If r = 0 Then
      MsgBox "Debe seleccionar al menos un Item de la Lista de Servidores", vbCritical, sMensaje
   Exit Sub
   End If
   If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
         
   cmdOpcion(Index).Enabled = False
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If RsPrinter.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
                  Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Consolidado.PaperOrientation = crLandscape
    
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
        
                   Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                   Consolidado.PaperOrientation = crLandscape
                   Consolidado.PrintOut
      
          
          Case Is = 2 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbDefault
               'On Error GoTo Cancelar

                  Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Consolidado.PaperOrientation = crLandscape
                  Consolidado.ExportOptions.FormatType = 21
                  Consolidado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Consolidado.ExportOptions.DiskFileName = cmdSave.FileName
                  Consolidado.Export False
       
         
          Case Is = 3 ' Salir
               Unload Me
   End Select
   cmdOpcion(Index).Enabled = True
Cancelar:
   Exit Sub
End Sub


Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   dtpFecIni.value = Date
   dtpFecFin.value = Date
   sNombre = dbTemporal(sCaja, 6, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "Etiqueta", "nvarchar(50)", "Grupo", "nvarchar(250)", "nventa1", "Float")
  ' sReporte = dbTemporal(sCaja, 11, "formaPago", "nvarchar(250)", "l0", "float", "l1", "float", "l2", "float", "l3", "float", "l4", "float", "l5", "float", "l6", "float", "l7", "float", "l8", "float", "l9", "float")
   'configuraListaServidores
   obtieneListaServidores
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
   Dim X As Integer
   Dim j As Integer
   Dim rsFormas As New ADODB.Recordset
   Dim cl As String
   Dim sector As String
   Dim rsvalor As New ADODB.Recordset
   Dim tiendas(9) As String
   Dim llocal As String
   Dim filtro As String
   Dim tcondicion1 As String
   Dim tcondicion2 As String
   Dim tcondicion3 As String
   
   tcondicion1 = ""
   tcondicion2 = ""
   tcondicion3 = ""
   
   filtro = ""
   If chkForma(0).value = 1 Then
        tcondicion1 = "CONTADO"
   End If
   
   If chkForma(1).value = 1 Then
        tcondicion2 = "CREDITO"
   End If
   If chkForma(2).value = 1 Then
        tcondicion3 = "PROPINA"
   End If
   
   If Len(tcondicion1) <> 0 Then
            filtro = "'" & tcondicion1 & "'"
   End If
   
   If Len(filtro) <> 0 Then
        If Len(tcondicion2) <> 0 Then
            filtro = filtro & ", '" & tcondicion2 & "'"
        End If
    Else
        If Len(tcondicion2) <> 0 Then
            filtro = "'" & tcondicion2 & "'"
        End If
   End If
   
    If Len(filtro) <> 0 Then
        If Len(tcondicion3) <> 0 Then
            filtro = filtro & ", '" & tcondicion3 & "'"
        End If
    Else
        If Len(tcondicion3) <> 0 Then
            filtro = "'" & tcondicion3 & "'"
        End If
   End If
 

   Set oComando = New clsComando
       If Me.chkFHoraria.value = 1 Then
            Dim dias, t As Integer
            dias = Me.dtpFecFin.value - Me.dtpFecIni.value
            For t = 0 To dias
                   fInicio = Format(dtpFecIni.value + t, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
                   fFinal = Format(dtpFecIni.value + t, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
                    For w = 1 To lstServidores.ListItems.Count
                If lstServidores.ListItems.Item(w).Checked = True Then
                        conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                        If conectado = False Then
                            MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                            lstServidores.ListItems.Item(w).Checked = False
                           GoTo sigue1
                        End If
                    
                        If Not oComando.CreateCmdSp("spRep_FormaPagoIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                            Set oComando = Nothing
                        Exit Sub
                        End If
                              oComando.CreateParameter "@fRegistroi", adDBDate, adParamInput, 20, fInicio
                              oComando.CreateParameter "@fRegistrof", adDBDate, adParamInput, 20, fFinal
                              If Not oComando.GetParamOK Then
                                    Set oComando = Nothing
                                    Exit Sub
                              End If
                           ' cs = LTrim(Mid(lstServidores.ListItems.Item(w), 1, 2))
                            cl = lstServidores.ListItems.Item(w)
                            sector = lstServidores.ListItems.Item(w).SubItems(1)
                            llocal = lstServidores.ListItems.Item(w).SubItems(2)
                       Set RsDatos = oComando.GetSP()
                     
                         If Not (RsDatos.EOF Or RsDatos.BOF) Then
                            RsDatos.MoveFirst
                                For z = 0 To RsDatos.RecordCount - 1
                                        Isql = "insert into " & sNombre & " values ( '" & sector & "','" & cl & "','" & llocal & "','" & LTrim(RsDatos.Fields(0)) & "', '" & LTrim(RsDatos.Fields(1)) & "', " & IIf(IsNull(RsDatos.Fields(2)), 0, RsDatos.Fields(2)) & " )"
        '                                 Debug.Print Isql
                                        Cn.Execute Isql
                                    RsDatos.MoveNext
                                Next z
                         End If
                End If
sigue1:
            Next w
            Next t
       Else
            fInicio = Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm")
            fFinal = Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:mm")
            For w = 1 To lstServidores.ListItems.Count
                If lstServidores.ListItems.Item(w).Checked = True Then
                        conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                        If conectado = False Then
                            MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                            lstServidores.ListItems.Item(w).Checked = False
                           GoTo sigue
                        End If
                    
                        If Not oComando.CreateCmdSp("spRep_FormaPagoIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                            Set oComando = Nothing
                        Exit Sub
                        End If
                              oComando.CreateParameter "@fRegistroi", adDBDate, adParamInput, 20, fInicio
                              oComando.CreateParameter "@fRegistrof", adDBDate, adParamInput, 20, fFinal
                              If Not oComando.GetParamOK Then
                                    Set oComando = Nothing
                                    Exit Sub
                              End If
                          '  cs = LTrim(Mid(lstServidores.ListItems.Item(w), 1, 2))
                            cl = lstServidores.ListItems.Item(w)
                            sector = lstServidores.ListItems.Item(w).SubItems(1)
                            llocal = lstServidores.ListItems.Item(w).SubItems(2)
                       Set RsDatos = oComando.GetSP()
                     
                         If Not (RsDatos.EOF Or RsDatos.BOF) Then
                            RsDatos.MoveFirst
                                For z = 0 To RsDatos.RecordCount - 1
                                        Isql = "insert into " & sNombre & " values ( '" & sector & "','" & cl & "','" & llocal & "','" & LTrim(RsDatos.Fields(0)) & "', '" & LTrim(RsDatos.Fields(1)) & "', " & IIf(IsNull(RsDatos.Fields(2)), 0, RsDatos.Fields(2)) & " )"
        '                                 Debug.Print Isql
                                        Cn.Execute Isql
                                    RsDatos.MoveNext
                                Next z
                         End If
                End If
sigue:
            Next w
        End If
    
    If chkForma(0).value = 1 And chkForma(1).value = 1 And chkForma(2).value = 1 Then
    
            Isql = "select sector,codlocal,local,grupo, nventa1 from " & sNombre
    Else
            Isql = "select sector,codlocal,local,grupo, nventa1 from " & sNombre & " WHERE etiqueta in(" & filtro & ")"
    End If
    
    Set RsPrinter = Lib.OpenRecordset(Isql, Cn)
    If optOpcion(0).value = True Then
    
                  ConsolidadoSector.DiscardSavedData
                  ConsolidadoSector.Database.SetDataSource RsPrinter
                  ConsolidadoSector.Text1.SetText "Agrupado Por: " & IIf(Me.optOpcion(0).value = True, "Sector", "Local") & ". Analisís del Intervalo del: " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") + " (Monto Expresado en Soles)"
                  ConsolidadoSector.Text20.SetText sRazonSocial
                  If Me.chkFHoraria.value = 1 Then
                    ConsolidadoSector.Text4.Suppress = False
                    ConsolidadoSector.Text4.SetText "Franja Horaria: " & Format(Me.dtpHorIni.value, "HH:MM") & " - " & Format(Me.dtpHorFin.value, "HH:MM") & " horas"
                    Else
                    ConsolidadoSector.Text4.Suppress = True
                  End If
                  ConsolidadoSector.Text2.SetText "Forma de Ingreso: " & filtro
                  frmEmite.CRViewer.DisplayGroupTree = False
                  frmEmite.CRViewer.ReportSource = ConsolidadoSector
    Else
                  Consolidado.DiscardSavedData
                  Consolidado.Database.SetDataSource RsPrinter
                  Consolidado.Text1.SetText "Agrupado Por: " & IIf(Me.optOpcion(0).value = True, "Sector", "Local") & ". Analisís del Intervalo del: " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") + " (Monto Expresado en Soles)"
                  Consolidado.Text20.SetText sRazonSocial
                  If Me.chkFHoraria.value = 1 Then
                    Consolidado.Text4.Suppress = False
                    Consolidado.Text4.SetText "Franja Horaria: " & Format(Me.dtpHorIni.value, "HH:MM") & " - " & Format(Me.dtpHorFin.value, "HH:MM") & " horas"
                    Else
                    Consolidado.Text4.Suppress = True
                  End If
                  Consolidado.Text2.SetText "Forma de Ingreso: " & filtro
                  frmEmite.CRViewer.DisplayGroupTree = False
                  frmEmite.CRViewer.ReportSource = Consolidado
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsPrinter = Nothing
   Set frmRepRegistroVentaIntegrado = Nothing
   Set RsCliente = Nothing
   Set RsEstado = Nothing
   Set RsTipoDocumento = Nothing
   If sReporte <> "" Then
      Cn.Execute "Drop Table " & sNombre
      Cn.Execute "Drop Table " & sReporte
   End If
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
   Dim p, w As Integer
    Dim s As Integer
    If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
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
       ' validaCantidadSeleccion
End Sub

