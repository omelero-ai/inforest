VERSION 5.00
Begin VB.Form frmPrecuentaImpresora 
   BorderStyle     =   0  'None
   Caption         =   "Impresoras del Sistema"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   Icon            =   "frmPrecuentaImpresora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   5475
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   1440
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   2760
         TabIndex        =   8
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   4080
         TabIndex        =   7
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   6
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   7
         Left            =   2760
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdImpresora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   8
         Left            =   4080
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
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
         Height          =   555
         Index           =   0
         Left            =   4080
         Picture         =   "frmPrecuentaImpresora.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1530
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "&Predeterminada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   1950
         Picture         =   "frmPrecuentaImpresora.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1530
         Width           =   1545
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "  Impresoras del Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5595
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2220
      TabIndex        =   11
      Top             =   1140
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrecuentaImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsImpresora As Recordset

Private Sub cmdImpresora_Click(Index As Integer)
  RsImpresora.MoveFirst
  RsImpresora.Find "Descripcion='" & cmdImpresora(Index).Caption & "'"
  If RsImpresora.EOF Then
     RsImpresora.MoveFirst
  End If
  sCodigo = RsImpresora!codigo
  wEnter = True
  Unload Me
End Sub

Private Sub CmdOpcion_Click(Index As Integer)
   If Index = 1 Then
      sCodigo = sPreCuenta
      wEnter = True
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   Centrar Me
   wEnter = False
   Set RsImpresora = Lib.OpenRecordset("select tImpresora as Codigo, tDescripcion as Descripcion from TIMPRESORA where tcaja='" & sCaja & "'", Cn)
   If RsImpresora.RecordCount > 0 Then
      AsignaComando 8, RsImpresora, cmdImpresora()
   Else
      MsgBox "No existe impresoras configuradas para esta caja", vbExclamation, sMensaje
      Exit Sub
   End If
End Sub

