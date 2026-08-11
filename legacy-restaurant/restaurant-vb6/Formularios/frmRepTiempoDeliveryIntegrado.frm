VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepTiempoDeliveryIntegrado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Tiempos en Delivery Integrado"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmRepTiempoDeliveryIntegrado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
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
      Left            =   5520
      TabIndex        =   15
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox chkTodo 
         Caption         =   "Seleccionar Todo"
         Height          =   360
         Left            =   2760
         TabIndex        =   11
         Top             =   3960
         Width           =   1695
      End
      Begin MSComctlLib.ListView lstServidores 
         Height          =   3735
         Left            =   120
         TabIndex        =   10
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
      Left            =   5040
      Picture         =   "frmRepTiempoDeliveryIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4500
      UseMaskColor    =   -1  'True
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
      Height          =   4395
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   "Para Análisis:"
         Height          =   1575
         Left            =   600
         TabIndex        =   16
         Top             =   1560
         Width           =   4335
         Begin VB.TextBox txtRango1 
            Height          =   375
            Left            =   1155
            TabIndex        =   4
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtRango2 
            Height          =   375
            Left            =   1155
            TabIndex        =   5
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rango 1 :"
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
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rango 2 :"
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
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "minutos"
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
            Left            =   3075
            TabIndex        =   18
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "minutos"
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
            Left            =   3075
            TabIndex        =   17
            Top             =   480
            Width           =   660
         End
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1710
         TabIndex        =   2
         Top             =   930
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   51970049
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1710
         TabIndex        =   0
         Top             =   390
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   51970049
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   420
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   51970051
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   930
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   51970051
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
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
         TabIndex        =   14
         Top             =   1020
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
         Left            =   270
         TabIndex        =   13
         Top             =   480
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
      Index           =   2
      Left            =   6495
      Picture         =   "frmRepTiempoDeliveryIntegrado.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4500
      UseMaskColor    =   -1  'True
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
      Left            =   3585
      Picture         =   "frmRepTiempoDeliveryIntegrado.frx":11DE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4500
      UseMaskColor    =   -1  'True
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
      Left            =   2130
      Picture         =   "frmRepTiempoDeliveryIntegrado.frx":1710
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4500
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   750
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepTiempoDeliveryIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsReporte       As ADODB.Recordset

Dim reporteDelivery  As New dsrTiempoDeliveryIntegrado

Dim sCriterio       As String
Dim sTitulo         As String
Dim sNombre         As String

Dim fInicio As Date
Dim fFinal As Date
 
Dim rango1 As Double
Dim rango2 As Double

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   Dim r, p As Integer
   r = 0
   p = 0
   sCriterio = ""
   sTitulo = ""
   Cn.Execute "delete from " & sNombre
   
   fInicio = Format(dtpFecIni, "YYYY/mm/dd") & " " & Format(dtpHorIni, "HH:mm:ss")
   fFinal = Format(dtpFecFin, "YYYY/mm/dd") & " " & Format(dtpHorFin, "HH:mm:ss")

   If fInicio > fFinal Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
   End If
      
   If Val(Me.txtRango1.Text) = Val(Me.txtRango2.Text) Then
         MsgBox "Error en Rangos para Análisis", vbInformation, sMensaje
         txtRango1.SetFocus
         Exit Sub
  End If
  
   If Val(txtRango1.Text) > Val(txtRango2.Text) Then
         MsgBox "Error en Rangos para Análisis", vbCritical, sMensaje
         txtRango1.SetFocus
         Exit Sub
  End If
  
  If Val(Me.txtRango1.Text) = 0 Or Val(Me.txtRango2.Text) = 0 Then
        MsgBox "Faltan Datos", vbInformation, sMensaje
        txtRango1.SetFocus
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
       
   sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
  
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
                frmEmite.CRViewer.DisplayGroupTree = True
                reporteDelivery.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                reporteDelivery.PaperOrientation = crPortrait
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
        
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
               
                reporteDelivery.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                reporteDelivery.PaperOrientation = crPortrait
                reporteDelivery.PrintOut

          Case Is = 2 ' Salir
               Unload Me
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
              reporteDelivery.ExportOptions.FormatType = 21
               reporteDelivery.ExportOptions.DestinationType = 1
               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
               cmdSave.ShowSave
               If cmdSave.FileName = "" Then
                  Exit Sub
               End If
               reporteDelivery.ExportOptions.DiskFileName = cmdSave.FileName
               reporteDelivery.Export False
   End Select
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    sNombre = ""
    sNombre = dbTemporal(sCaja, 11, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "pedido", "Float", "total", "Float", "promedio", "Float", "minimo", "Float", "maximo", "Float", "ta", "Float", "tb", "Float", "tc", "Float")
    dtpFecIni.value = Date
    dtpFecFin.value = Date
    obtieneListaServidores
End Sub
Private Sub chkTodo_Click()
     activaCheck IIf(chkTodo.value = 1, True, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Cn.Execute "Drop Table " & sNombre
End Sub

Public Sub Genera()
        Screen.MousePointer = vbHourglass
       On Error GoTo sigue
        Dim oComando As New clsComando
        Dim conectado As Boolean
       
        Dim RsDatos As New ADODB.Recordset
        Dim RsCabecera As New ADODB.Recordset
        Dim w, z As Long
   
        Dim cs As String
        Dim cl As String
        Dim Cabecera As String
        Dim sector As String
        Dim llocal As String
        Set oComando = New clsComando
        rango1 = Val(Me.txtRango1.Text)
        rango2 = Val(Me.txtRango2.Text)
        
        fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
        fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
   
        For w = 1 To lstServidores.ListItems.Count
        
        If lstServidores.ListItems.Item(w).Checked = True Then
               
                conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                If conectado = False Then
                    MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                    lstServidores.ListItems.Item(w).Checked = False
                    GoTo sigue
                End If
 
                Set oComando = New clsComando
                If Not oComando.CreateCmdSp("spRep_TiempoDeliveryIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                   Set oComando = Nothing
                   Exit Sub
                End If
 
                oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
                oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                oComando.CreateParameter "@r1", adDouble, adParamInput, 8, Val(rango1)
                oComando.CreateParameter "@r2", adDouble, adParamInput, 8, Val(rango2)
                
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
                                          Isql = "insert into " & sNombre & " values ( '" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "'," & RsDatos.Fields(1) & "," & RsDatos.Fields(2) & "," & RsDatos.Fields(3) & "," & RsDatos.Fields(4) & "," & RsDatos.Fields(5) & "," & RsDatos.Fields(6) & "," & RsDatos.Fields(7) & ")"
                                          Cn.Execute Isql
                          RsDatos.MoveNext
                          Next z
                End If
        End If
sigue:
    Next w
        Isql = "select * from " & sNombre & " order by 3"
        Set rsReporte = Lib.OpenRecordset(Isql, Cn)
           reporteDelivery.DiscardSavedData
            reporteDelivery.Database.SetDataSource rsReporte
            reporteDelivery.Text13.SetText sRazonSocial
            reporteDelivery.Text3.SetText "Análisis Rango 1: " & Me.txtRango1.Text & " -   Rango 2: " & Me.txtRango2.Text
            reporteDelivery.ReportTitle = sTitulo
            frmEmite.CRViewer.ReportSource = reporteDelivery
 
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

Private Sub lstServidores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
       End If
   ' validaCantidadSeleccion
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

 
Private Sub txtRango1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRango2.SetFocus
    End If
End Sub

Private Sub txtRango1_LostFocus()
  If Val(txtRango1.Text) > 0 Then
      txtRango1.Text = Format(Val(txtRango1.Text), "#,##0")
   Else
     txtRango1.Text = "0.00"
   End If
 End Sub
 
Private Sub txtRango2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOpcion(0).SetFocus
    End If
End Sub

 Private Sub txtRango2_LostFocus()
  If Val(txtRango2.Text) > 0 Then
      txtRango2.Text = Format(Val(txtRango2.Text), "#,##0")
   Else
     txtRango2.Text = "0.00"
   End If
 End Sub
