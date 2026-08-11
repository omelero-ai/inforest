VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFechaEmision 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fecha de Emision - Caja Contingencia"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5280
   DrawMode        =   10  'Mask Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Fecha de Emision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   2
         Left            =   3600
         Picture         =   "frmFechaEmision.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dtfechaemision 
         Height          =   465
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   820
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16908289
         CurrentDate     =   37539
      End
   End
End
Attribute VB_Name = "frmFechaEmision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public femision As Date

Private Sub cmdOpcion_Click(Index As Integer)
If Me.dtfechaemision.value > FechaServidor() Then
    MsgBox "No se puede ingresar una fecha mayor a la del dia de Hoy!"
    Exit Sub
End If
Dim dayanul As Double
dayProceso = Calcular("SELECT DATEDIFF(DAY,'" & Format(Me.dtfechaemision.value, "yyyyMMdd") & "',GETDATE()) as Codigo ", Cn)
If dayProceso > DiasContingenciaDocumentos Then
    MsgBox "Tiempo maximo de Registro de Documento: " & DiasContingenciaDocumentos & " Dias", vbInformation, sMensaje
    Exit Sub
End If
femision = Me.dtfechaemision.value
MsgBox "Fecha de Emision Agregada Correctamente " & mensajeOperador, vbInformation, sMensaje
Unload Me
End Sub

Private Sub Form_Load()
femision = FechaServidor()
dtfechaemision.value = FechaServidor()
End Sub
