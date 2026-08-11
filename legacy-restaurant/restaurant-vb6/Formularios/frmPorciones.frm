VERSION 5.00
Begin VB.Form frmPorciones 
   Caption         =   "Configuración de Porciones"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmPorciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   5850
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
      Height          =   615
      Index           =   0
      Left            =   4650
      Picture         =   "frmPorciones.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2655
      Width           =   1170
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
      Height          =   615
      Index           =   1
      Left            =   3420
      Picture         =   "frmPorciones.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2655
      Width           =   1170
   End
   Begin VB.TextBox txtFactor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   4860
      TabIndex        =   12
      Top             =   2205
      Width           =   915
   End
   Begin VB.TextBox txtFactor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   4860
      TabIndex        =   9
      Top             =   1755
      Width           =   915
   End
   Begin VB.TextBox txtFactor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   4860
      TabIndex        =   6
      Top             =   1305
      Width           =   915
   End
   Begin VB.TextBox txtFactor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   4860
      TabIndex        =   3
      Top             =   855
      Width           =   915
   End
   Begin VB.TextBox txtPorcion 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   3915
      TabIndex        =   11
      Top             =   2205
      Width           =   825
   End
   Begin VB.TextBox txtPorcion 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   3915
      TabIndex        =   8
      Top             =   1755
      Width           =   825
   End
   Begin VB.TextBox txtPorcion 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3915
      TabIndex        =   5
      Top             =   1305
      Width           =   825
   End
   Begin VB.TextBox txtPorcion 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   3915
      TabIndex        =   2
      Top             =   855
      Width           =   825
   End
   Begin VB.TextBox txtLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   1350
      MaxLength       =   25
      TabIndex        =   10
      Top             =   2205
      Width           =   2445
   End
   Begin VB.TextBox txtLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1350
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1755
      Width           =   2445
   End
   Begin VB.TextBox txtLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1350
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1305
      Width           =   2445
   End
   Begin VB.TextBox txtLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1350
      MaxLength       =   25
      TabIndex        =   1
      Top             =   855
      Width           =   2445
   End
   Begin VB.CheckBox chkPorcion 
      Caption         =   "Activar Tabla de Porciones"
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
      Left            =   495
      TabIndex        =   0
      Top             =   135
      Width           =   2760
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Factor del Precio"
      Height          =   420
      Left            =   4860
      TabIndex        =   21
      Top             =   360
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Porción"
      Height          =   195
      Left            =   3960
      TabIndex        =   20
      Top             =   585
      Width           =   780
   End
   Begin VB.Label Label5 
      Caption         =   "Label de la Porción"
      Height          =   195
      Left            =   1350
      TabIndex        =   19
      Top             =   585
      Width           =   1635
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Porcion (4) :"
      Height          =   240
      Left            =   180
      TabIndex        =   18
      Top             =   2250
      Width           =   1050
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Porcion (3) :"
      Height          =   240
      Left            =   180
      TabIndex        =   17
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Porcion (2) :"
      Height          =   240
      Left            =   180
      TabIndex        =   16
      Top             =   1350
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Porcion (1) :"
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   923
      Width           =   1050
   End
End
Attribute VB_Name = "frmPorciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPorcion As Recordset

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 0 Then
      If RsPorcion.RecordCount > 0 Then
         Isql = "Update TPORCION Set " & _
                 "tLabel1 = '" & txtLabel(1).Text & "', " & _
                 "tLabel2 = '" & txtLabel(2).Text & "', " & _
                 "tLabel3 = '" & txtLabel(3).Text & "', " & _
                 "tLabel4 = '" & txtLabel(4).Text & "', " & _
                 "nPorcion1 = '" & txtPorcion(1).Text & "', " & _
                 "nPorcion2 = '" & txtPorcion(2).Text & "', " & _
                 "nPorcion3 = '" & txtPorcion(3).Text & "', " & _
                 "nPorcion4 = '" & txtPorcion(4).Text & "', " & _
                 "nFactor1 = '" & txtFactor(1).Text & "', " & _
                 "nFactor2 = '" & txtFactor(2).Text & "', " & _
                 "nFactor3 = '" & txtFactor(3).Text & "', " & _
                 "nFactor4 = '" & txtFactor(4).Text & "', " & _
                 "lPorcion =" & IIf(chkPorcion.Value, 1, 0)
      Else
         Isql = "Insert TPORCION values (" & IIf(chkPorcion.Value, 1, 0) & "," & _
                "'" & txtLabel(1).Text & "', " & txtPorcion(1).Text & "," & txtFactor(1).Text & ", " & _
                "'" & txtLabel(2).Text & "', " & txtPorcion(2).Text & "," & txtFactor(2).Text & ", " & _
                "'" & txtLabel(3).Text & "', " & txtPorcion(3).Text & "," & txtFactor(3).Text & ", " & _
                "'" & txtLabel(4).Text & "', " & txtPorcion(4).Text & "," & txtFactor(4).Text & ") "
      End If
      Cn.Execute Isql
      MsgBox "Datos Actualizados", vbInformation, sMensaje
      Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   Isql = "select * from TPORCION"
   Set RsPorcion = Lib.OpenRecordset(Isql, Cn)
   
   If Not RsPorcion.EOF Then
      RsPorcion.MoveFirst
      txtLabel(1).Text = IIf(IsNull(RsPorcion!tLabel1), "", RsPorcion!tLabel1)
      txtLabel(2).Text = IIf(IsNull(RsPorcion!tLabel2), "", RsPorcion!tLabel2)
      txtLabel(3).Text = IIf(IsNull(RsPorcion!tLabel3), "", RsPorcion!tLabel3)
      txtLabel(4).Text = IIf(IsNull(RsPorcion!tLabel4), "", RsPorcion!tLabel4)
      txtPorcion(1).Text = Format(IIf(IsNull(RsPorcion!nPorcion1), 0, RsPorcion!nPorcion1), "##0.000")
      txtPorcion(2).Text = Format(IIf(IsNull(RsPorcion!nPorcion2), 0, RsPorcion!nPorcion2), "##0.000")
      txtPorcion(3).Text = Format(IIf(IsNull(RsPorcion!nPorcion3), 0, RsPorcion!nPorcion3), "##0.000")
      txtPorcion(4).Text = Format(IIf(IsNull(RsPorcion!nPorcion4), 0, RsPorcion!nPorcion4), "##0.000")
      txtFactor(1).Text = Format(IIf(IsNull(RsPorcion!nFactor1), 0, RsPorcion!nFactor1), "##0.000")
      txtFactor(2).Text = Format(IIf(IsNull(RsPorcion!nFactor2), 0, RsPorcion!nFactor2), "##0.000")
      txtFactor(3).Text = Format(IIf(IsNull(RsPorcion!nFactor3), 0, RsPorcion!nFactor3), "##0.000")
      txtFactor(4).Text = Format(IIf(IsNull(RsPorcion!nFactor4), 0, RsPorcion!nFactor4), "##0.000")
      chkPorcion.Value = IIf(IsNull(RsPorcion!lPorcion), 0, IIf(RsPorcion!lPorcion = True, 1, 0))
   Else
      chkPorcion.Value = False
      txtLabel(1).Text = ""
      txtLabel(2).Text = ""
      txtLabel(3).Text = ""
      txtLabel(4).Text = ""
      txtPorcion(1).Text = "0.000"
      txtPorcion(2).Text = "0.000"
      txtPorcion(3).Text = "0.000"
      txtPorcion(4).Text = "0.000"
      txtFactor(1).Text = "0.000"
      txtFactor(2).Text = "0.000"
      txtFactor(3).Text = "0.000"
      txtFactor(4).Text = "0.000"
   End If
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsPorcion = Nothing
   Set frmPorcion = Nothing
End Sub

Private Sub txtFactor_GotFocus(Index As Integer)
   Enfoque txtFactor(Index)
End Sub

Private Sub txtFactor_KeyPress(Index As Integer, KeyAscii As Integer)
   TabNext (KeyAscii)
   Numerico KeyAscii, txtFactor(Index).Text
End Sub

Private Sub txtFactor_LostFocus(Index As Integer)
   txtFactor(Index).Text = Format(txtFactor(Index).Text, "##0.000")
End Sub

Private Sub txtLabel_KeyPress(Index As Integer, KeyAscii As Integer)
   TabNext (KeyAscii)
End Sub

Private Sub txtPorcion_GotFocus(Index As Integer)
   Enfoque txtPorcion(Index)
End Sub

Private Sub txtPorcion_KeyPress(Index As Integer, KeyAscii As Integer)
   TabNext (KeyAscii)
   Numerico KeyAscii, txtPorcion(Index).Text
End Sub

Private Sub txtPorcion_LostFocus(Index As Integer)
   txtPorcion(Index).Text = Format(txtPorcion(Index).Text, "##0.000")
End Sub
