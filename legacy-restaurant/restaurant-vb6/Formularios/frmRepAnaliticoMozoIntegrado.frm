VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepAnaliticoMozoIntegrado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas Por Mozos"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12570
   Icon            =   "frmRepAnaliticoMozoIntegrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7860
      Begin VB.Frame Frame3 
         Caption         =   " Origen de Datos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   5520
         TabIndex        =   32
         Top             =   1560
         Width           =   2205
         Begin VB.OptionButton optOrigen 
            Caption         =   "Producción (a+b+c)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(a) Ventas"
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
            Top             =   572
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(b) Cortesías"
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
            Left            =   135
            TabIndex        =   15
            Top             =   844
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "(c) Ctas. Corrientes"
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
            Left            =   135
            TabIndex        =   16
            Top             =   1116
            Width           =   2025
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Combinación"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   18
            Top             =   1660
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Cargos"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   19
            Top             =   1935
            Width           =   1935
         End
         Begin VB.OptionButton optOrigen 
            Caption         =   "Pedidos Facturados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   17
            Top             =   1388
            Width           =   2025
         End
      End
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
         Height          =   1305
         Left            =   3720
         TabIndex        =   31
         Top             =   1560
         Width           =   1700
         Begin VB.OptionButton optOpcion 
            Caption         =   "Resumido"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   585
            Width           =   1305
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Consolidado"
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   900
            Width           =   1305
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Valores "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1700
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Costo"
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   6
            Top             =   900
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   5
            Top             =   585
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
      End
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
         Height          =   1305
         Left            =   1920
         TabIndex        =   29
         Top             =   1560
         Width           =   1700
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Sector"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Local"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   11
            Top             =   585
            Width           =   1035
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Mozos"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   12
            Top             =   900
            Value           =   -1  'True
            Width           =   1395
         End
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1515
         TabIndex        =   2
         Top             =   915
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   56885249
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1515
         TabIndex        =   0
         Top             =   480
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   56885249
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3690
         TabIndex        =   3
         Top             =   900
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
         Format          =   56885251
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3690
         TabIndex        =   1
         Top             =   480
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
         Format          =   56885251
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin VB.Label Label2 
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
         Left            =   300
         TabIndex        =   34
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label Label3 
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
         Left            =   390
         TabIndex        =   33
         Top             =   975
         Width           =   990
      End
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
      Height          =   4935
      Left            =   7920
      TabIndex        =   25
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox chkTodo 
         Caption         =   "Seleccionar Todo"
         Height          =   360
         Left            =   2760
         TabIndex        =   26
         Top             =   4440
         Width           =   1695
      End
      Begin MSComctlLib.ListView lstServidores 
         Height          =   4215
         Left            =   120
         TabIndex        =   27
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
      Left            =   2640
      Picture         =   "frmRepAnaliticoMozoIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   24
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
      Left            =   4110
      Picture         =   "frmRepAnaliticoMozoIntegrado.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
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
      Left            =   8490
      Picture         =   "frmRepAnaliticoMozoIntegrado.frx":132E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5040
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
      Left            =   5565
      Picture         =   "frmRepAnaliticoMozoIntegrado.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   21
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
      Left            =   7020
      Picture         =   "frmRepAnaliticoMozoIntegrado.frx":1C42
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepAnaliticoMozoIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim RsReporte As Recordset
Dim RsTipo As Recordset
Dim RsTipoPedido As Recordset
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsMotorizado As Recordset
Dim RsTurno As Recordset
Dim Detallado As New dsrAnaliticoMotorizadoIntegradoDetallado
Dim Consolidado As New dsrAnaliticoMotorizadoIntegradoConsolidado
Dim xMotorizado As New dsrAnaliticoMotorizadoIntegradoDetalladoXM
Dim cMotorizado As New dsrAnaliticoMotorizadoIntegradoConsolidadoXM
Dim crMotorizado As New dsrAnaliticoMotorizadoIntegradoConsolidadoRES
Dim sCriterio As String
Dim xCriterio As String
Dim sTitulo As String
Dim sProducto As String
Dim sTurno As String

Dim fInicio As Date
Dim fFinal As Date


Private Sub chkTodo_Click()
 activaCheck IIf(chkTodo.value = 1, True, False)
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

Private Sub CmdActualiza_Click()
If MsgBox("Esta operación puede tardar varios minutos" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
    Screen.MousePointer = vbHourglass
    configuraListaServidores
    Me.chkTodo.value = 0
    MsgBox "Lista De Servidores Actualizada", vbInformation, sMensaje
    Screen.MousePointer = vbDefault
End If
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


Private Sub cmdOpcion_Click(Index As Integer)
    Cn.Execute " delete from " & sNombre
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
   
    If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
        MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
        Exit Sub
        dtpFecFin.SetFocus
    End If
    sTitulo = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
      
   
   
    Select Case Index
        Case Is = 0 ' Preview
            Genera
            If RsReporte.EOF = True Then
               Screen.MousePointer = vbDefault
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            frmEmite.CRViewer.DisplayGroupTree = True
            
            If optOpcion(0).value Then
                If optTipo(2).value = True Then
                    xMotorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    xMotorizado.PaperOrientation = crPortrait
                Else
                
                    Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Detallado.PaperOrientation = crPortrait
                End If
            ElseIf optOpcion(1).value = True Then
                If optTipo(2).value = True Then
                    cMotorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    cMotorizado.PaperOrientation = crPortrait
                Else
                    Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Consolidado.PaperOrientation = crPortrait
                End If
            Else
                    crMotorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    crMotorizado.PaperOrientation = crPortrait
            End If
            
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
               
        Case Is = 1 ' Imprimir
            Genera
            Screen.MousePointer = vbDefault
            If RsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            If optOpcion(0).value Then
                If optTipo(2).value = True Then
                    xMotorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    xMotorizado.PaperOrientation = crPortrait
                    xMotorizado.PrintOut
                Else
                
                    Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Detallado.PaperOrientation = crPortrait
                    Detallado.PrintOut
                End If
            ElseIf optOpcion(1).value = True Then
                If optTipo(2).value = True Then
                    cMotorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    cMotorizado.PaperOrientation = crPortrait
                    cMotorizado.PrintOut
                Else
                    Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    Consolidado.PaperOrientation = crPortrait
                    Consolidado.PrintOut
                End If
            Else
                    crMotorizado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                    crMotorizado.PaperOrientation = crPortrait
                    crMotorizado.PrintOut
            End If
        Case Is = 2 ' Salir
            Unload Me
          
        Case Is = 3 ' Exportar
            Genera
            Screen.MousePointer = vbDefault
            If RsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            If optOpcion(0).value Then
                If optTipo(2).value = True Then
                    xMotorizado.ExportOptions.FormatType = 21
                    xMotorizado.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                    Exit Sub
                    End If
                    xMotorizado.ExportOptions.DiskFileName = cmdSave.FileName
                    xMotorizado.Export False
                
    
                Else
                    Detallado.ExportOptions.FormatType = 21
                    Detallado.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                        Exit Sub
                    End If
                    Detallado.ExportOptions.DiskFileName = cmdSave.FileName
                    Detallado.Export False
                End If
            ElseIf optOpcion(1).value = True Then
                If optTipo(2).value = True Then
                    cMotorizado.ExportOptions.FormatType = 21
                    cMotorizado.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                    Exit Sub
                    End If
                    cMotorizado.ExportOptions.DiskFileName = cmdSave.FileName
                    cMotorizado.Export False
                Else
                    Consolidado.ExportOptions.FormatType = 21
                    Consolidado.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                       Exit Sub
                    End If
                    Consolidado.ExportOptions.DiskFileName = cmdSave.FileName
                    Consolidado.Export False
                    End If
            Else
            
                    crMotorizado.ExportOptions.FormatType = 21
                    crMotorizado.ExportOptions.DestinationType = 1
                    cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                    cmdSave.ShowSave
                    If cmdSave.FileName = "" Then
                       Exit Sub
                    End If
                    crMotorizado.ExportOptions.DiskFileName = cmdSave.FileName
                    crMotorizado.Export False
                    
            End If
           
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
    sNombre = ""
    sNombre = dbTemporal(sCaja, 12, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "Mozo", "nvarchar(150)", "TipoProducto", "nvarchar(150)", "grupo", "nvarchar(150)", "subgrupo", "nvarchar(150)", "producto", "nvarchar(150)", "tcodigopedido", "nvarchar(150)", "cantidad", "float", "Ventas", "Float", "npedidos", "Float")
    dtpFecIni.value = Date
    dtpFecFin.value = Date
    obtieneListaServidores
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsReporte = Nothing
   Set RsTurno = Nothing
   Set RsTipoPedido = Nothing
   Set RsGrupo = Nothing
   Set RsSubGrupo = Nothing
   Set RsTipo = Nothing
   Set RsMotorizado = Nothing
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
   Dim sPrecio As String
   Set oComando = New clsComando
   Set Detallado = New dsrAnaliticoMotorizadoIntegradoDetallado
   Set Consolidado = New dsrAnaliticoMotorizadoIntegradoConsolidado
   Set xMotorizado = New dsrAnaliticoMotorizadoIntegradoDetalladoXM
   Set cMotorizado = New dsrAnaliticoMotorizadoIntegradoConsolidadoXM
   Set crMotorizado = New dsrAnaliticoMotorizadoIntegradoConsolidadoRES
   
       ' venta / neto / costo
     If Me.optOrigen(4).value = True Then ' combinacion
        If optValor(0).value = True Then ' venta
             sPrecio = "dbo.CPEDIDO.nVenta * dbo.CPEDIDO.nCantidad"
        ElseIf optValor(1).value = True Then ' neto
            sPrecio = "dbo.CPEDIDO.nPrecioNeto  * dbo.DPEDIDO.nCantidad"
        ElseIf optValor(2).value = True Then ' costo
              sPrecio = "(isnull(dbo.CPEDIDO.nInsumo,0)+isnull(dbo.CPEDIDO.nGasto,0)+isnull(dbo.CPEDIDO.nManoObra,0)) * dbo.CPEDIDO.nCantidad"
        End If
    Else
        If optValor(0).value = True Then '  venta
                sPrecio = "dbo.DPEDIDO.nPrecioVenta * dbo.DPEDIDO.nCantidad"
        ElseIf optValor(1).value = True Then ' neto
                sPrecio = "dbo.DPEDIDO.nPrecioNeto * dbo.DPEDIDO.nCantidad"
        ElseIf optValor(2).value = True Then ' costo
                sPrecio = "(isnull(dbo.DPEDIDO.nInsumo,0)+isnull(dbo.DPEDIDO.nGasto,0)+isnull(dbo.DPEDIDO.nManoObra,0)) * dbo.DPEDIDO.nCantidad"
        End If
    End If
     
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
                If Not oComando.CreateCmdSp("spRep_AnaliticoMozoIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                Exit Sub
                End If
                      oComando.CreateParameter "@flagProduccion", adBoolean, adParamInput, 1, optOrigen(0).value
                      oComando.CreateParameter "@flagVenta", adBoolean, adParamInput, 1, optOrigen(1).value
                      oComando.CreateParameter "@flagCortesia", adBoolean, adParamInput, 1, optOrigen(2).value
                      oComando.CreateParameter "@flagCuentaCte", adBoolean, adParamInput, 1, optOrigen(3).value
                      oComando.CreateParameter "@flagPedidosFacturados", adBoolean, adParamInput, 1, optOrigen(6).value
                      oComando.CreateParameter "@flagCombinacion", adBoolean, adParamInput, 1, optOrigen(4).value
                      oComando.CreateParameter "@flagCargo", adBoolean, adParamInput, 1, optOrigen(5).value
                      oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
                      oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                      oComando.CreateParameter "@sPrecio", adVarChar, adParamInput, 150, sPrecio
                   
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
                                Isql = "insert into " & sNombre & " values ('" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "','" & RsDatos.Fields(2) & "','" & RsDatos.Fields(3) & "','" & RsDatos.Fields(4) & "','" & RsDatos.Fields(5) & "'," & RsDatos.Fields(6) & "," & RsDatos.Fields(7) & "," & RsDatos.Fields(8) & ")"
                              '  Debug.Print Isql
                                Cn.Execute Isql
                            RsDatos.MoveNext
                        Next z
                 End If
        End If
sigue:
    Next w
    ' MsgBox err.Description
    Dim Tipo As String
     If optOpcion(0).value = True Then ' DETALLADO
        
        If optTipo(0).value = True Then ' sector
            Tipo = "Sector"
                Isql = "select sector, mozo,  npedidos , local +' ' + tcodigopedido as tcodigopedido, sum(ventas) as ventas from " & sNombre & " group by  sector, mozo, npedidos,local+' ' + tcodigopedido order by sector, mozo,npedidos "
        ElseIf optTipo(1).value = True Then ' local
            Tipo = "Local"
                Isql = "select local as sector, mozo, npedidos,tcodigopedido, sum(ventas) as ventas from " & sNombre & " group by  local, mozo, npedidos,tcodigopedido order by local, mozo,npedidos "
        ElseIf optTipo(2).value = True Then ' motorizado
             Tipo = "Mozo"
                Isql = "select local as sector, mozo, npedidos,tcodigopedido, sum(ventas) as ventas  from " & sNombre & " group by  local, mozo, npedidos,tcodigopedido order by local, mozo,npedidos "
        End If
          
              Set RsReporte = Lib.OpenRecordset(Isql, Cn)
              If optTipo(2).value = True Then
              
                  xMotorizado.DiscardSavedData
                  xMotorizado.Text10.SetText "Venta Por Mozos"
                  xMotorizado.Text1.SetText "Mozo"
                  xMotorizado.Text8.SetText ""
                  xMotorizado.Database.SetDataSource RsReporte
                  xMotorizado.Text9.SetText "Local"
                  xMotorizado.ReportTitle = "Agrupado por " & Tipo & " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                  xMotorizado.Text13.SetText sRazonSocial
                  frmEmite.CRViewer.DisplayGroupTree = False
                  frmEmite.CRViewer.ReportSource = xMotorizado
              
              Else
                  Detallado.DiscardSavedData
                  Detallado.Database.SetDataSource RsReporte
                  Detallado.Text1.SetText "Mozo"
                  Detallado.Text10.SetText "Ventas Por Mozos"
                  Detallado.Text8.SetText ""
                  Detallado.Text9.SetText Tipo
                  Detallado.ReportTitle = "Agrupado por " & Tipo & " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                  Detallado.Text13.SetText sRazonSocial
                  frmEmite.CRViewer.DisplayGroupTree = False
                  frmEmite.CRViewer.ReportSource = Detallado
              End If
                  
     ElseIf optOpcion(1).value = True Then ' resumido
            If optTipo(0).value = True Then ' sector
            Tipo = "Sector"
                    Isql = "select sector, mozo, count(distinct local+' '+tcodigopedido) as npedidos, sum(ventas) as ventas  from " & sNombre & " group by  sector,  mozo order by sector, mozo,3 "
            ElseIf optTipo(1).value = True Then ' local
             Tipo = "Local"
                    Isql = "select local as sector, mozo, count(distinct local+' '+tcodigopedido) as npedidos, sum(ventas) as ventas  from " & sNombre & " group by  local, mozo, npedidos order by local, mozo,npedidos "
            ElseIf optTipo(2).value = True Then ' motorizado
            Tipo = "Mozo"
                 Isql = "select local as sector, mozo, count(distinct local +' '+tcodigopedido) as npedidos,sum(ventas) as ventas from " & sNombre & " group by  local, mozo order by local, mozo"
            End If
        
                Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                If optTipo(2).value = True Then
                  cMotorizado.DiscardSavedData
                  cMotorizado.Database.SetDataSource RsReporte
                  cMotorizado.Text10.SetText "Venta Por Mozos"
                  cMotorizado.Text1.SetText "Mozo"
                  cMotorizado.Text8.SetText ""
                  cMotorizado.ReportTitle = "Agrupado por " & Tipo & " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                  cMotorizado.Text13.SetText sRazonSocial
                  frmEmite.CRViewer.DisplayGroupTree = False
                  frmEmite.CRViewer.ReportSource = cMotorizado
                
                Else

                  Consolidado.DiscardSavedData
                  Consolidado.Text10.SetText "Venta Por Mozos"
                  Consolidado.Text7.SetText ""
                  Consolidado.Database.SetDataSource RsReporte
                  Consolidado.ReportTitle = "Agrupado por " & Tipo & " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                  Consolidado.Text6.SetText Tipo
                  Consolidado.Text13.SetText sRazonSocial
                  frmEmite.CRViewer.DisplayGroupTree = False
                  frmEmite.CRViewer.ReportSource = Consolidado
                End If
     Else ' consolidado
                 Isql = "select mozo,  count(distinct local+' '+tcodigopedido) as npedidos  , sum(ventas) as ventas  from " & sNombre & " group by  mozo,NPEDIDOS order by mozo "
                 Set RsReporte = Lib.OpenRecordset(Isql, Cn)
                  crMotorizado.DiscardSavedData
                  crMotorizado.Text10.SetText "Venta Por Mozos"
                  crMotorizado.Text6.SetText "Mozo"
                  crMotorizado.Text7.SetText ""
                  crMotorizado.Database.SetDataSource RsReporte
                  crMotorizado.ReportTitle = "Agrupado por " & Tipo & " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                  crMotorizado.Text13.SetText sRazonSocial
                  frmEmite.CRViewer.DisplayGroupTree = False
                  frmEmite.CRViewer.ReportSource = crMotorizado
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

Private Sub optOpcion_Click(Index As Integer)
    Select Case Index
        Case 2
            If Me.optOpcion(2).value = True Then
                Me.optTipo(0).Enabled = False
                Me.optTipo(1).Enabled = False
            Else
                Me.optTipo(0).Enabled = True
                Me.optTipo(1).Enabled = True
            End If
        Case Else
                Me.optTipo(0).Enabled = True
                Me.optTipo(1).Enabled = True
    End Select
End Sub
