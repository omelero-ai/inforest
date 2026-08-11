VERSION 5.00
Begin VB.Form frmListaMotivos 
   Caption         =   "Motivos de Anulación"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEliminacion 
      Caption         =   " Motivo de Anulación"
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
      Height          =   6585
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   2505
         Picture         =   "frmListaMotivos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5625
         Width           =   1515
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   2505
         TabIndex        =   38
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   8
         Left            =   1725
         TabIndex        =   37
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   7
         Left            =   930
         TabIndex        =   36
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   150
         TabIndex        =   35
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   5
         Left            =   3300
         TabIndex        =   34
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   4
         Left            =   2505
         TabIndex        =   33
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1725
         TabIndex        =   32
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   930
         TabIndex        =   31
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   150
         TabIndex        =   30
         Top             =   330
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   10
         Left            =   3300
         TabIndex        =   29
         Top             =   1080
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   14
         Left            =   2505
         TabIndex        =   28
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   13
         Left            =   1725
         TabIndex        =   27
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   12
         Left            =   930
         TabIndex        =   26
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   11
         Left            =   150
         TabIndex        =   25
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   15
         Left            =   3300
         TabIndex        =   24
         Top             =   1845
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   19
         Left            =   2505
         TabIndex        =   23
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   18
         Left            =   1725
         TabIndex        =   22
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   17
         Left            =   930
         TabIndex        =   21
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   16
         Left            =   150
         TabIndex        =   20
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   20
         Left            =   3300
         TabIndex        =   19
         Top             =   2595
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   24
         Left            =   2505
         TabIndex        =   18
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   23
         Left            =   1725
         TabIndex        =   17
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   22
         Left            =   930
         TabIndex        =   16
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   21
         Left            =   150
         TabIndex        =   15
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   25
         Left            =   3300
         TabIndex        =   14
         Top             =   3360
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   29
         Left            =   2505
         TabIndex        =   13
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   28
         Left            =   1725
         TabIndex        =   12
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   27
         Left            =   930
         TabIndex        =   11
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   26
         Left            =   150
         TabIndex        =   10
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   30
         Left            =   3300
         TabIndex        =   9
         Top             =   4110
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   34
         Left            =   2505
         TabIndex        =   8
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   33
         Left            =   1710
         TabIndex        =   7
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   32
         Left            =   930
         TabIndex        =   6
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   31
         Left            =   150
         TabIndex        =   5
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   35
         Left            =   3300
         TabIndex        =   4
         Top             =   4860
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   36
         Left            =   150
         TabIndex        =   3
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   37
         Left            =   930
         TabIndex        =   2
         Top             =   5625
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminacion 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   38
         Left            =   1710
         TabIndex        =   1
         Top             =   5625
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmListaMotivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMotivoEliminacion As Recordset

Private Sub cmdEliminacion_Click(Index As Integer)
  RsMotivoEliminacion.MoveFirst
  RsMotivoEliminacion.Find "Descripcion='" & cmdEliminacion(Index).Caption & "'"
  If RsMotivoEliminacion.EOF Then
     RsMotivoEliminacion.MoveFirst
  End If
  sCodigo = RsMotivoEliminacion!codigo
  sDescrip = RsMotivoEliminacion!Descripcion
  wEnter = True
  Unload Me
End Sub

Private Sub CmdOpcion_Click(Index As Integer)
    Select Case Index
         Case Is = 16 'Nuevo
         wEnter = False
         Unload Me
    End Select
    

End Sub

Private Sub Form_Load()
  Centrar Me
   wEnter = False
   Set RsMotivoEliminacion = Lib.OpenRecordset("select * from vMotivoANULACION where lActivo = 1 order by Codigo", Cn)
   If RsMotivoEliminacion.RecordCount > 0 Then
      AsignaComando 38, RsMotivoEliminacion, cmdEliminacion()
   Else
      MsgBox "No existe Motivos configuradas para esta caja", vbExclamation, sMensaje
      Exit Sub
   End If
End Sub
