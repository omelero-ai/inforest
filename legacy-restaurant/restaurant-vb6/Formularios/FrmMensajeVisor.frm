VERSION 5.00
Begin VB.Form FrmMensajeVisor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Timer TimerVisor 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7440
         Top             =   1680
      End
      Begin VB.CommandButton CmdAceptar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CONFIRMACION POR CAJERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CANCELAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3960
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "ESPERANDO CONFIRMACION DEL CLIENTE...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   7695
      End
   End
End
Attribute VB_Name = "FrmMensajeVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RESULTADOVISOR As Boolean
Public OPERADORMENSAJE As Integer
Private Sub CmdAceptar_Click()
    TimerVisor.Enabled = False
    Cn.Execute "update infovisor set estado=3   where tcaja='" & sCaja & "'"
    OPERADORMENSAJE = 0
    RESULTADOVISOR = True
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    TimerVisor.Enabled = False
    OPERADORMENSAJE = 0
    Cn.Execute "update infovisor set estado=4  where tcaja='" & sCaja & "'"
    RESULTADOVISOR = False
    Unload Me
End Sub

Private Sub Form_Load()
    TimerVisor.Enabled = True
    TimerVisor.Interval = 1000
End Sub

Private Sub TimerVisor_Timer()

  If Calcular("select top 1 estado as codigo from infovisor   where tcaja='" & sCaja & "'", Cn) = 3 Then
    RESULTADOVISOR = True
    OPERADORMENSAJE = 0
    TimerVisor.Enabled = False
    Unload Me
    
  Else
    If Calcular("select top 1 estado as codigo from infovisor   where tcaja='" & sCaja & "'", Cn) = 4 Then
        RESULTADOVISOR = False
        OPERADORMENSAJE = 1
        TimerVisor.Enabled = False
         Cn.Execute "update infovisor set estado=1  where tcaja='" & sCaja & "'"
        Unload Me
    End If
  End If

End Sub
