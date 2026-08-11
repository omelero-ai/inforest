VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepProductosNoEnlazados 
   Caption         =   "Reporte Productos No Enlazados"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   Icon            =   "frmRepProductosNoEnlazados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   7335
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
      Left            =   3660
      Picture         =   "frmRepProductosNoEnlazados.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1755
      Visible         =   0   'False
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
      Left            =   5130
      Picture         =   "frmRepProductosNoEnlazados.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1755
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
      Left            =   2205
      Picture         =   "frmRepProductosNoEnlazados.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1755
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
      Left            =   750
      Picture         =   "frmRepProductosNoEnlazados.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1755
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   6240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   1725
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7125
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
         Left            =   4845
         TabIndex        =   9
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2145
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
         Left            =   4845
         TabIndex        =   8
         Top             =   735
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.CheckBox chkTipoProducto 
         Caption         =   "Todos Los Tipos"
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
         Left            =   4845
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   380
         Value           =   1  'Checked
         Width           =   2070
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         Top             =   345
         Width           =   2985
         _ExtentX        =   5265
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
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1755
         TabIndex        =   10
         Top             =   720
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   1755
         TabIndex        =   11
         Top             =   1065
         Width           =   2985
         _ExtentX        =   5265
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
         Left            =   720
         TabIndex        =   13
         Top             =   1125
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
         Left            =   1050
         TabIndex        =   12
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Producto :"
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
         Left            =   70
         TabIndex        =   7
         Top             =   380
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmRepProductosNoEnlazados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim rsReporte As Recordset
Dim RsProdNoEnlazados As ADODB.Recordset
Dim RsTipoproductos As ADODB.Recordset
Dim RsGrupos As ADODB.Recordset
Dim RsSubGrupos As ADODB.Recordset

Dim clsrProdNoEnlazados As New dsrRepProductosNoEnlazados
Dim sCriterio As String
Dim sCondicion As String
Dim sProducto As String
Dim sGrupo As String
Dim sSubgrupo As String

Private Sub cboGrupo_Click(Area As Integer)
      cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubGrupo " & IIf(cboGrupo.BoundText = "", "", "where tGrupo = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubGrupos = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupos
    End With
    sProducto = ""
End Sub

Private Sub chkGrupo_Click()
If chkGrupo.Value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
   Else
      cboGrupo.Enabled = True
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

Private Sub chkTipoProducto_Click()
 If chkTipoProducto.Value = 1 Then
      cboTipoProducto.Enabled = False
      cboTipoProducto.Text = ""
   Else
      cboTipoProducto.Enabled = True
   End If
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
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
                  clsrProdNoEnlazados.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  clsrProdNoEnlazados.PaperOrientation = crPortrait
               
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
               
                  clsrProdNoEnlazados.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  clsrProdNoEnlazados.PaperOrientation = crPortrait
                  clsrProdNoEnlazados.PrintOut
             
          
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               
                  clsrProdNoEnlazados.ExportOptions.FormatType = 21
                  clsrProdNoEnlazados.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  clsrProdNoEnlazados.ExportOptions.DiskFileName = cmdSave.FileName
                  clsrProdNoEnlazados.Export False
              
               
   End Select

End Sub



Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    
End Sub

Sub LlenaCombos()
    With cboTipoProducto
         Isql = "Select * from Vtipoproducto where lactivo=1"
         Set RsTipoproductos = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoproductos
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboGrupo
         Isql = "Select * from vGrupo order by Descripcion"
         Set RsGrupos = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupos
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
           
    With cboSubGrupo
         Isql = "Select * from vSubGrupo order by Descripcion"
         Set RsSubGrupos = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupos
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
End Sub
Public Sub Genera()
   Dim oComando As clsComando
   sProducto = ""
   sGrupo = ""
   sSubgrupo = ""
   
   sCondicion = ""
   If chkTipoProducto.Value = 0 Then
      If cboTipoProducto.Text = "" Then
         MsgBox "Debe escoger un Tipo de Prodcuto", vbCritical, sMensaje
         Exit Sub
      End If
      sProducto = cboTipoProducto.BoundText
   End If
    If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger un grupo", vbCritical, sMensaje
         Exit Sub
      End If
      sGrupo = cboGrupo.BoundText
   End If
   If chkSubGrupo.Value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger un SubGrupo", vbCritical, sMensaje
         Exit Sub
      End If
      sSubgrupo = cboSubGrupo.BoundText
   End If
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_ProductosNoEnlazados", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
  
        oComando.CreateParameter "@ttipopro", adVarChar, adParamInput, 20, sProducto
   
        oComando.CreateParameter "@tgrupo", adVarChar, adParamInput, 20, sGrupo
   
        oComando.CreateParameter "@tsubgrupo", adVarChar, adParamInput, 20, sSubgrupo
    
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    rsReporte.Filter = sCondicion
 
       clsrProdNoEnlazados.DiscardSavedData
       clsrProdNoEnlazados.Database.SetDataSource rsReporte
       'clsrProdNoEnlazados.txtTipoProductos.SetText = cboTipoProducto.Text
       'clsrProdNoEnlazados.ReportTitle = sCriterio
       'clsrProdNoEnlazados.Text12.SetText "Propinas Resumido"
       'clsrProdNoEnlazados.Text15.SetText sRazonSocial
       'clsrProdNoEnlazados.Text7.SetText localConectado
       frmEmite.CRViewer.ReportSource = clsrProdNoEnlazados
   
   Screen.MousePointer = vbDefault
End Sub


