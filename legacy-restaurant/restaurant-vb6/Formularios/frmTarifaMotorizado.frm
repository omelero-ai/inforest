VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTarifaMotorizado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarifa de Motorizado"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Grabar"
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
      Left            =   1440
      Picture         =   "frmTarifaMotorizado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1170
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
      Left            =   4920
      Picture         =   "frmTarifaMotorizado.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo cboTarifa 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
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
   Begin VB.Label lblModificacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Ultima Modificación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Tarifa Actual: "
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmTarifaMotorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTarifaMotorizado As New ADODB.Recordset


Public Sub LlenaCombo()
        With cboTarifa
             
             Isql = "select 'nTarifaLV' as Codigo, 'Tarifa Lun - Sab' as Descripcion  union select 'nTarifaSD' as Codigo ,'Tarifa Dom' as Descripcion  union select 'nTarifaEs' as Codigo, 'Tarifa Especial' as Descripcion"
             Set rsTarifaMotorizado = Lib.OpenRecordset(Isql, Cn)
             Set .RowSource = rsTarifaMotorizado
                 .DataField = "Descripcion"
                 .ListField = "Descripcion"
                 .BoundColumn = "Codigo"
        End With
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
    Case 3
        Unload Me
    Case 1
        If Supervisor("21") = False Then
                MsgBox "Clave no permitida", vbExclamation, sMensaje
                Exit Sub
        End If
        Cn.Execute "update tparametro set tTarifaActualMotorizado ='" & cboTarifa.Text & "', tUsuarioTarifa='" & UCase(sVar1) & "',fRegistroTarifa=getdate()"
        lblModificacion.Caption = Calcular("select isnull(tUsuarioTarifa,'')  as codigo from tparametro", Cn) & "  -  " & Calcular("select isnull(fregistrotarifa,'')  as codigo from tparametro", Cn)
        tTarifaActualMotorizado = cboTarifa.Text
        mdiDespachador.StatusBar.Panels.Item(1).Text = "Caja : " & sCaja & "  /  Tarifa Actual Motorizado : " & tTarifaActualMotorizado
           sVar1 = ""
    End Select
End Sub

Private Sub Form_Initialize()
        Set rsTarifaMotorizado = New ADODB.Recordset
        lblModificacion.Caption = Calcular("select isnull(tUsuarioTarifa,'')  as codigo from tparametro", Cn) & "  -  " & Calcular("select isnull(fregistrotarifa,'')  as codigo from tparametro", Cn)
End Sub

Private Sub Form_Load()
    LlenaCombo
    cboTarifa.Text = IIf(IsNull(tTarifaActualMotorizado), "", tTarifaActualMotorizado)
 
End Sub

Private Sub Form_Terminate()
    Set rsTarifaMotorizado = Nothing
End Sub

