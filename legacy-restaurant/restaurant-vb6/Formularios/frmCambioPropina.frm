VERSION 5.00
Begin VB.Form frmCambioPropina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Propina"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmCambioPropina.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
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
      Height          =   585
      Index           =   0
      Left            =   2970
      Picture         =   "frmCambioPropina.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1230
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
      Height          =   585
      Index           =   1
      Left            =   4380
      Picture         =   "frmCambioPropina.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1230
   End
   Begin VB.CommandButton cmdPropina 
      Caption         =   "Propina S/. 0.00"
      Height          =   585
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   1230
   End
   Begin VB.CommandButton cmdPropina 
      Caption         =   "Propina US$ 0.00"
      Height          =   585
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmCambioPropina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wFlag As Boolean
Private Sub cmdOpcion_Click(Index As Integer)
   If wFlag And Index = 0 Then
      wEnter = True
   Else
      wEnter = False
   End If
   Unload Me
End Sub

Private Sub cmdPropina_Click(Index As Integer)
   sTipo = ""
   frmNumPad.Show vbModal
   If Index = 0 Then  ' Propina MN
      If wEnter Then
         wFlag = True
         frmPedidoCorrelativo.nPropina = Val(sDescrip)
         frmPedidoCorrelativo.tPropina = "01"
         cmdPropina(0).Caption = "Propina " & sMonN & " " & Format(Val(sDescrip), "###,##0.00")
         cmdPropina(1).Caption = "Propina " & sMonE & " 0.00"
         cmdPropina(0).FontBold = True
         cmdPropina(1).FontBold = False
      End If
   Else               ' Propina ME
      If wEnter Then
         wFlag = True
         frmPedidoCorrelativo.nPropina = Val(sDescrip)
         frmPedidoCorrelativo.tPropina = "02"
         cmdPropina(0).Caption = "Propina " & sMonN & " 0.00"
         cmdPropina(1).Caption = "Propina " & sMonE & " " & Format(Val(sDescrip), "###,##0.00")
         cmdPropina(0).FontBold = False
         cmdPropina(1).FontBold = True
      End If
   End If
   wEnter = False
End Sub

Private Sub Form_Load()
   Centrar Me
   wFlag = False
   cmdPropina(1).Caption = "Propina " & sMonE & " 0.00"
   cmdPropina(0).Caption = "Propina " & sMonN & " 0.00"
End Sub
