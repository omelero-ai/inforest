VERSION 5.00
Begin VB.Form frmAreasImpresora 
   BorderStyle     =   0  'None
   Caption         =   "Areas de producción"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   Icon            =   "frmAreasImpresora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   5415
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
         Left            =   4080
         Picture         =   "frmAreasImpresora.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1530
         Width           =   1215
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
      TabIndex        =   0
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "frmAreasImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsArea As Recordset

Private Sub cmdImpresora_Click(Index As Integer)
  RsArea.MoveFirst
  RsArea.Find "Descripcion='" & cmdImpresora(Index).Caption & "'"
  If RsArea.EOF Then
     RsArea.MoveFirst
  End If
  sCodigo = RsArea!Codigo
  wEnter = True
  Unload Me
End Sub

Private Sub cmdOpcion_Click()
   wEnter = False
   Unload Me
End Sub

Private Sub Form_Load()
   Centrar Me
   wEnter = False
   Set RsArea = Lib.OpenRecordset("select Codigo, Descripcion from vArea where lActivo=1", Cn)
   If RsArea.RecordCount > 0 Then
      AsignaComando 8, RsArea, cmdImpresora()
   Else
      MsgBox "No existe áreas configuradas para esta caja", vbExclamation, sMensaje
      Exit Sub
   End If
End Sub
