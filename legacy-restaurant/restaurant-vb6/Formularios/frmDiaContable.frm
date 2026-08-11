VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiaContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje día Contable"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   Icon            =   "frmDiaContable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   5475
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cierre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   120
         Picture         =   "frmDiaContable.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1042
         Width           =   1575
      End
      Begin VB.CommandButton cmdApertura 
         Caption         =   "Apertura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   120
         Picture         =   "frmDiaContable.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1020
         Width           =   1575
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
         Height          =   630
         Left            =   3780
         Picture         =   "frmDiaContable.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1042
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpDiaContable 
         Height          =   495
         Left            =   3300
         TabIndex        =   2
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57737217
         CurrentDate     =   41144
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   1860
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dia Contable Actual :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   450
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmDiaContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDiaContable As clsDiaContable
Dim fdiacontable As Date
Dim modo As String

'botones
Private Sub cmdApertura_Click()
    If validaFechaIngreso = True Then
        oDiaContable.insertarDiaContable Me.dtpDiaContable.value
        lIniciaPorDiaContable = True
        lDiaContableAperturado = True
        Unload Me
    Else
        MsgBox "La fecha seleccionada es menor al último Día Contable registrado", vbCritical, sMensaje
        dtpDiaContable.SetFocus
    End If

End Sub

Private Sub cmdCerrar_Click()
    If MsgBox("¿Desea Cerrar el Dia Contable?", vbYesNo + vbQuestion, sMensaje) = vbYes Then
            Cn.Execute " update tdiacontable set lcierre=1, tusuariocierre='" & sUsuario & "', fregistrocierre=getdate() where fdiacontable='" & Format(Me.dtpDiaContable.value, "yyyyMMdd") & "' "
            Unload Me
    End If
End Sub

Private Sub cmdOpcion_Click()
    If lIniciaPorDiaContable = False Then '
            MsgBox "No puede Iniciar Venta sin Aperturar Dia Contable", vbCritical, sMensaje
    End If
    Unload Me
End Sub

'formularios
Private Sub Form_Initialize()
    Set oDiaContable = New clsDiaContable
End Sub

Private Sub Form_Load()
    lDiaContableAperturado = False
    cmdCerrar.Visible = False
    cmdApertura.Visible = False
    fdiacontable = oDiaContable.obtieneDiaContable
    Me.dtpDiaContable.value = fdiacontable
    If modo = "Apertura" Then
            Me.Label1.Visible = True
            If Format(fdiacontable, "yyyyMMdd") = "19900101" Then
                    Me.Label1.Caption = "Dia Contable a Aperturar :"
                    Me.Label2.Caption = "Proceso de aprtura del día contable"
                    Me.dtpDiaContable.Enabled = True
                    Me.dtpDiaContable.value = oDiaContable.obtieneFechaServidor
                    Me.cmdApertura.Visible = True
                Else
                    Me.dtpDiaContable.Enabled = False
                    Me.Label1.Caption = "Dia Contable Actual :"
                    Me.Label2.Caption = "Para cerrar este Dia Contable se debe cerrar el turno"
                    lIniciaPorDiaContable = True
                    lDiaContableAperturado = True
            End If
    Else
        Me.Label1.Caption = "Dia Contable a Cerrar :"
        Me.Label2.Caption = "Proceso de cierre del día contable"
        Me.cmdCerrar.Visible = True
        Me.cmdCerrar.Enabled = True
        lIniciaPorDiaContable = True
        lDiaContableAperturado = True
    End If
End Sub

Private Sub Form_Terminate()
    Set oDiaContable = Nothing
End Sub

'metodos y procedimientos
Public Sub obtieneModoIngreso(X As String)
    modo = X
End Sub

Public Function validaFechaIngreso() As Boolean
    validaFechaIngreso = True
    Dim fechaMaxima As Date
    fechaMaxima = Calcular("select max(fdiacontable)  as codigo from tdiacontable", Cn)
    If Format(Me.dtpDiaContable.value, "yyyyMMdd") < Format(fechaMaxima, "yyyyMMdd") Then
        validaFechaIngreso = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmDiaContable = Nothing
End Sub

