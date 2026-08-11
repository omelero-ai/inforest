VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepOcupabilidad 
   Caption         =   "Reporte de Ocupabilidad"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   Icon            =   "frmRepOcupabilidad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7110
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
      Left            =   1785
      Picture         =   "frmRepOcupabilidad.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1260
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
      Left            =   3045
      Picture         =   "frmRepOcupabilidad.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1350
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
      Left            =   5730
      Picture         =   "frmRepOcupabilidad.frx":132E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1350
   End
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
      Height          =   4410
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   7065
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
         Height          =   1065
         Left            =   1635
         TabIndex        =   22
         Top             =   3195
         Width           =   1905
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
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   24
            Top             =   225
            Value           =   -1  'True
            Width           =   1290
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
            Height          =   300
            Index           =   1
            Left            =   135
            TabIndex        =   23
            Top             =   585
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3690
         TabIndex        =   18
         Top             =   3195
         Width           =   3165
         Begin VB.OptionButton optOpcion 
            Caption         =   "Solo Horas Ocupadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   20
            Top             =   225
            Value           =   -1  'True
            Width           =   2355
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Todas las horas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   19
            Top             =   585
            Width           =   1875
         End
      End
      Begin VB.CheckBox chkMesa 
         Caption         =   "Todas las Mesas"
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
         Left            =   4965
         TabIndex        =   16
         Top             =   1115
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkSalon 
         Caption         =   "Todos los Salones"
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
         Left            =   4965
         TabIndex        =   13
         Top             =   715
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Todos los Locales"
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
         Left            =   4965
         TabIndex        =   12
         Top             =   315
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CommandButton cmdBusca 
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
         Index           =   1
         Left            =   3990
         Picture         =   "frmRepOcupabilidad.frx":1420
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1470
         Width           =   765
      End
      Begin VB.TextBox txtTurno 
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
         Left            =   1635
         TabIndex        =   0
         Top             =   1500
         Width           =   2265
      End
      Begin VB.CheckBox chkTurno 
         Caption         =   "Todos los Turnos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4965
         TabIndex        =   2
         Top             =   1515
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1635
         TabIndex        =   5
         Top             =   2330
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
         Format          =   20971521
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1635
         TabIndex        =   3
         Top             =   1900
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
         Format          =   20971521
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3660
         TabIndex        =   4
         Top             =   1890
         Width           =   1095
         _ExtentX        =   1931
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
         Format          =   20971523
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3660
         TabIndex        =   6
         Top             =   2310
         Width           =   1095
         _ExtentX        =   1931
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
         Format          =   20971523
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboLocal 
         Height          =   315
         Left            =   1635
         TabIndex        =   14
         Top             =   300
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
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   1635
         TabIndex        =   15
         Top             =   700
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
      Begin MSDataListLib.DataCombo cboMesa 
         Height          =   315
         Left            =   1635
         TabIndex        =   17
         Top             =   1100
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
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   345
         Left            =   1635
         TabIndex        =   21
         Top             =   2760
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
         CustomFormat    =   "HH"
         Format          =   20971523
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
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
         Left            =   555
         TabIndex        =   32
         Top             =   2397
         Width           =   990
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
         Left            =   465
         TabIndex        =   31
         Top             =   1967
         Width           =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno :"
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
         Left            =   960
         TabIndex        =   30
         Top             =   1552
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local :"
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
         Index           =   8
         Left            =   945
         TabIndex        =   29
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salón :"
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
         Left            =   930
         TabIndex        =   28
         Top             =   760
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mesa :"
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
         Left            =   960
         TabIndex        =   27
         Top             =   1160
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Reporte :"
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
         Left            =   135
         TabIndex        =   26
         Top             =   3240
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Inicial :"
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
         Left            =   570
         TabIndex        =   25
         Top             =   2827
         Width           =   975
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
      Left            =   4380
      Picture         =   "frmRepOcupabilidad.frx":1522
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1350
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1125
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepOcupabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim sTemporal As String
Dim RsReporte As Recordset
Dim RsTurno As Recordset
Dim RsLocal As Recordset
Dim RsSalon As Recordset
Dim RsMesa As Recordset
Dim ReporteR As New dsrOcupabilidadR
Dim ReporteD As New dsrOcupabilidadD
Dim sCriterio As String
Dim sTitulo As String
Dim sTurno As String
Dim sHora(48) As String
Dim HoraInicial(48) As Integer
Dim HoraFinal(48) As Integer
Dim sOrden(48) As Integer
Dim i As Integer
Dim nMesa As Double
Dim nAdulto As Double

Sub LlenaCombos()
    With cboLocal
         Isql = "Select * from vLocal"
         Set RsLocal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsLocal
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
   
    With cboSalon
         Isql = "Select * from vSalon"
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboMesa
         Isql = "Select * from tMesa order by tDetallado"
         Set RsMesa = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMesa
         .DataField = "tDetallado"
         .ListField = "tDetallado"
         .BoundColumn = "tCodigoMesa"
    End With
End Sub

Private Sub chkLocal_Click()
   If chkLocal.Value = 1 Then
      cboLocal.Enabled = False
      cboLocal.Text = ""
      RsSalon.Filter = adFilterNone
   Else
      cboLocal.Enabled = True
   End If
End Sub

Private Sub chkMesa_Click()
   If chkMesa.Value = 1 Then
      cboMesa.Enabled = False
      cboMesa.Text = ""
   Else
      cboMesa.Enabled = True
   End If
End Sub

Private Sub chkSalon_Click()
   If chkSalon.Value = 1 Then
      cboSalon.Enabled = False
      cboSalon.Text = ""
   Else
      cboSalon.Enabled = True
   End If
End Sub

Private Sub cmdBusca_Click(Index As Integer)
Dim xCriterio As String
    Isql = "Select tTurno as Codigo, tCaja, convert(nVarChar,fInicial, 120) as Descripcion, convert(nvarchar,fFinal,120) as fFinal from MTURNO order by tTurno DESC"
    frmBusca.nPredeterm = 0
    Call ConfGrilla(4, frmBusca.grdGrilla, "Turno", 2, "Codigo", 1400, 0, 0, "", _
                                           "Caja", 2, "tCaja", 1100, 0, 0, "", _
                                           "Fecha Inicial", 2, "Descripcion", 2400, 0, 0, "", _
                                           "Fecha Final", 2, "fFinal", 2400, 0, 0, "")
    frmBusca.Show vbModal
    If Not wEnter Then
       Exit Sub
    End If
    sTurno = sCodigo
    txtTurno.Text = sCodigo
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 2 Then
        Unload Me
        Exit Sub
    End If
   
    sCriterio = ""
    sTitulo = ""
    
    If chkTurno.Value = 0 Then
        If sTurno = "" Then
           MsgBox "Debe escoger un Turno", vbCritical, sMensaje
           Exit Sub
        End If
      
        Set RsTurno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
        If RsTurno.RecordCount > 0 Then
            sTitulo = " Del " & Format(RsTurno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(RsTurno!fFinal), Format(RsTurno!finicial, "dd/mmm/yyyy") & " 23:59", RsTurno!fFinal), "dd/mmm/yyyy HH:mm")
            sCriterio = " MPEDIDO.tTurno ='" & sTurno & "'"
        Else
            MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
        End If
    Else
        If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
            MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
            Exit Sub
            dtpFecFin.SetFocus
        End If
        sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
        sCriterio = " MPEDIDO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "' "
    End If
      
    If chkLocal.Value = 0 Then
       If cboLocal.Text = "" Then
          MsgBox "Debe escoger un Local", vbCritical, sMensaje
          Exit Sub
       End If
       sCriterio = sCriterio & " and vSalon.tLocal ='" & cboLocal.BoundText & "'"
    End If
    
    If chkSalon.Value = 0 Then
       If cboSalon.Text = "" Then
          MsgBox "Debe escoger un Salón", vbCritical, sMensaje
          Exit Sub
       End If
       sCriterio = sCriterio & " and MPEDIDO.tSalon ='" & cboSalon.BoundText & "'"
    End If
    
    If chkMesa.Value = 0 Then
       If cboMesa.Text = "" Then
          MsgBox "Debe escoger una Mesa", vbCritical, sMensaje
          Exit Sub
       End If
       sCriterio = sCriterio & " and MPEDIDO.tMEsa ='" & cboMesa.BoundText & "'"
    End If
          
    Select Case Index
        Case Is = 0 ' Preview
            Genera
            If RsReporte.EOF = True Then
               Screen.MousePointer = vbDefault
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            frmEmite.CRViewer.DisplayGroupTree = False
            
            If optOpcion(0).Value Then
               ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteR.PaperOrientation = crPortrait
            Else
               ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteD.PaperOrientation = crPortrait
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
            If optOpcion(0).Value Then
               ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteR.PaperOrientation = crPortrait
               ReporteR.PrintOut
            Else
               ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteD.PaperOrientation = crPortrait
               ReporteD.PrintOut
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
            
            If optOpcion(0).Value Then
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
               ReporteD.ExportOptions.FormatType = 21
               ReporteD.ExportOptions.DestinationType = 1
               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
               cmdSave.ShowSave
               If cmdSave.FileName = "" Then
                  Exit Sub
               End If
               ReporteD.ExportOptions.DiskFileName = cmdSave.FileName
               ReporteD.Export False
            End If
   End Select

End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub chkTurno_Click()
   If chkTurno.Value = 1 Then
      cmdBusca(1).Enabled = False
      txtTurno.Text = ""
      sTurno = ""
      txtTurno.Text = ""
      dtpFecIni.Enabled = True
      dtpFecFin.Enabled = True
      dtpHorIni.Enabled = True
      dtpHorFin.Enabled = True
   Else
      cmdBusca(1).Enabled = True
      dtpFecIni.Enabled = False
      dtpFecFin.Enabled = False
      dtpHorIni.Enabled = False
      dtpHorFin.Enabled = False
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    sTemporal = dbTemporal(sCaja, 14, "Orden", "Int", "Hora", "nvarchar(20)", "Mesa", "nVarchar(4)", "Adicional", "float", "Adulto", "float", "Nino", "float", "Venta", "float", "tMesa", "float", "tAdulto", "float", "nDias", "float", "tCodigoPedido", "nvarchar(10)", "dMesa", "nvarchar(50)", "FechaInicial", "smalldatetime", "FechaFinal", "smalldatetime")
    LlenaCombos
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cmdBusca(1).Enabled = False
    dtpHora.Value = Date
    
    For i = 0 To 23
        HoraInicial((i * 2) + 1) = (i * 60)
        HoraInicial((i * 2) + 2) = (i * 60) + 30
        
        sHora((i * 2) + 1) = IIf(i < 10, "0" & LTrim(Str(i)), LTrim(Str(i))) & ":00"
        sHora((i * 2) + 2) = IIf(i < 10, "0" & LTrim(Str(i)), LTrim(Str(i))) & ":30"
        
        sOrden((i * 2) + 1) = ((i * 2) + 1)
        sOrden((i * 2) + 2) = ((i * 2) + 2)
    Next i
        
    Do While True
       sNombre = "#TMP" & Mid("000000", 1, 6 - Len(Trim(Str(DateDiff("s", Date, Now))))) + Trim(Str(DateDiff("s", Date, Now)))
       If Calcular("select count(name) as Codigo from tempdb.dbo.sysobjects where id = object_id(N'tempdb.[dbo].[" & sNombre & "]')", Cn) < 1 Then
          Exit Do
       End If
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cn.Execute "drop table " & sTemporal
    Set RsReporte = Nothing
    Set RsTurno = Nothing
End Sub

Public Sub Genera()
    Dim nHora As Integer
    Dim nContador As Integer
    Dim nDia As Integer
    Screen.MousePointer = vbHourglass
    nHora = Hour(dtpHora.Value) * 2 + 1
    nContador = 0
    For i = nHora To 48
        nContador = nContador + 1
        sOrden(i) = nContador
    Next i
    
    For i = 1 To nHora - 1
        nContador = nContador + 1
        sOrden(i) = nContador
    Next i
    nDia = DateDiff("h", Format(dtpFecIni.Value, "yyyy/mm/dd") + " " + Format(dtpHorIni.Value, "HH:nn"), Format(dtpFecFin.Value, "yyyy/mm/dd") + " " + Format(dtpHorFin.Value, "HH:nn"))
    nDia = IIf(Round(nDia / 24, 0) = 0, 1, Round(nDia / 24, 0))
        
    nMesa = Calcular("select count(tCodigoMesa) as Codigo from TMESA where lActivo=1 and tcodigoMesa<>'000'", Cn)
    nAdulto = Calcular("Select sum(nPersona) as Codigo from TMESA where lActivo=1 and tcodigoMesa<>'000'", Cn)
    
    Cn.Execute "delete from " & sTemporal
    
    Isql = "SELECT dbo.MPEDIDO.tCodigoPedido, MAX(dbo.MPEDIDO.fRegistro) AS Inicio, MAX(dbo.MDOCUMENTO.fRegistro) AS Final, MAX(dbo.MPEDIDO.tMesa) AS Mesa, MAX(T1.nAdicional) AS nMesa, MAX(dbo.MPEDIDO.nAdulto) AS Adulto, MAX(dbo.MPEDIDO.nNino) AS Nino, SUM(dbo.DPEDIDO.nVenta) AS Venta, dbo.TMESA.tDetallado AS dMesa " & _
           "INTO " & sNombre & " " & _
           "FROM dbo.MPEDIDO INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido INNER JOIN dbo.MDOCUMENTO ON dbo.DPEDIDO.tDocumento = dbo.MDOCUMENTO.tDocumento LEFT OUTER JOIN " & _
           "(SELECT dbo.TPEDIDOMESA.tCodigoPedido, COUNT(dbo.TPEDIDOMESA.tMesa) AS nAdicional FROM dbo.MPEDIDO INNER JOIN dbo.TPEDIDOMESA ON dbo.MPEDIDO.tCodigoPedido = dbo.TPEDIDOMESA.tCodigoPedido WHERE MPEDIDO.tTipoPedido='01' and tEstadoPedido<>'03' and " & sCriterio & " GROUP BY dbo.TPEDIDOMESA.tCodigoPedido) T1 " & _
           "ON dbo.MPEDIDO.tCodigoPedido = T1.tCodigoPedido LEFT OUTER JOIN dbo.TMESA ON dbo.MPEDIDO.tMesa = dbo.TMESA.tCodigoMesa " & _
           "where MPEDIDO.tTipoPedido='01' and tEstadoPedido<>'03' and tMesa <> '' " & _
           "and " & sCriterio & " " & _
           "Group by MPEDIDO.tCodigoPedido, dbo.TMESA.tDetallado"
    Cn.Execute Isql
        
    For i = 1 To 48
        If optOpcion(0).Value Then
           Isql = "Insert into " & sTemporal & " (Orden, Hora, Mesa, Adicional, Adulto, Nino, Venta, tMesa, tAdulto, nDias) " & _
           "select " & sOrden(i) & " as Orden, '" & sHora(i) & "' as Hora, " & _
           "count(Mesa) as Mesa, isnull(sum(nMesa),0) as Adicional, isnull(sum(Adulto),0) as Adulto, isnull(sum(Nino),0) as Nino, isnull(sum(Venta),0) as Venta, " & nMesa & " as tMesa, " & nAdulto & " as tAdulto, " & nDia & " as nDias " & _
           "from " & sNombre & " " & _
           "Where (DatePart(hh, Inicio) * 60 + DatePart(mi, Inicio) <= " & HoraInicial(i) & " And DatePart(hh, Final) * 60 + DatePart(mi, Final) >= " & HoraInicial(i) & ") or " & _
           " (DatePart(hh, Inicio) * 60 + DatePart(mi, Inicio) >= " & HoraInicial(i) & " And DatePart(hh, Final) * 60 + DatePart(mi, Final) >= " & HoraInicial(i) & " and datepart(dayofyear,inicio)+1 = datepart(dayofyear,final)) or " & _
           " (DatePart(hh, Inicio) * 60 + DatePart(mi, Inicio) >= " & HoraInicial(i) & " And DatePart(hh, Final) * 60 + DatePart(mi, Final) <= " & HoraInicial(i) + 30 & " and datepart(dayofyear,inicio) = datepart(dayofyear,final)) "
        Else
           Isql = "Insert into " & sTemporal & " " & _
           "select " & sOrden(i) & " as Orden, '" & sHora(i) & "' as Hora, " & _
           "1 as Mesa, isnull(nMesa,0) as Adicional, isnull(Adulto,0), isnull(Nino,0), isnull(Venta,0), " & nMesa & " as tMesa, " & nAdulto & " as tAdulto, " & nDia & " as nDias, tCodigoPedido, dMesa, Inicio, Final " & _
           "from " & sNombre & " " & _
           "Where (DatePart(hh, Inicio) * 60 + DatePart(mi, Inicio) <= " & HoraInicial(i) & " And DatePart(hh, Final) * 60 + DatePart(mi, Final) >= " & HoraInicial(i) & ") or " & _
           " (DatePart(hh, Inicio) * 60 + DatePart(mi, Inicio) >= " & HoraInicial(i) & " And DatePart(hh, Final) * 60 + DatePart(mi, Final) >= " & HoraInicial(i) & " and datepart(dayofyear,inicio)+1 = datepart(dayofyear,final)) or " & _
           " (DatePart(hh, Inicio) * 60 + DatePart(mi, Inicio) >= " & HoraInicial(i) & " And DatePart(hh, Final) * 60 + DatePart(mi, Final) <= " & HoraInicial(i) + 30 & " and datepart(dayofyear,inicio) = datepart(dayofyear,final)) "
        End If
        Cn.Execute Isql
    Next i
    
    Cn.Execute "drop table " & sNombre
    If optOpcion(0).Value = True Then
       Isql = "select Orden, Hora, Mesa, Adicional, Adulto, Nino, Venta, tMesa, tAdulto, nDias "
    Else
       Isql = "select * "
    End If
    
    If optOpcion(2).Value = True Then
       Isql = Isql & "from " & sTemporal & " where Mesa > 0 order by Orden"
    Else
       Isql = Isql & "from " & sTemporal & " order by Orden"
    End If
    Set RsReporte = Lib.OpenRecordset(Isql, Cn)
    
    If optOpcion(0).Value Then
       ReporteR.DiscardSavedData
       ReporteR.Database.SetDataSource RsReporte
       ReporteR.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
       ReporteR.Text5.SetText sRazonSocial
       frmEmite.CRViewer.ReportSource = ReporteR
    Else
       ReporteD.DiscardSavedData
       ReporteD.Database.SetDataSource RsReporte
       ReporteD.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
          ReporteD.Text15.SetText sRazonSocial
       frmEmite.CRViewer.ReportSource = ReporteD
    End If
End Sub

Private Sub cboLocal_Click(Area As Integer)
    cboSalon.Text = ""
    With cboSalon
         Isql = "Select * from vSalon " & IIf(cboLocal.BoundText = "", "", "where tLocal = '" & cboLocal.BoundText & "'") & " order by Descripcion "
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
    End With
End Sub
