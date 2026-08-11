VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmCierrePeriodo 
   Caption         =   "Cierre de Periodo"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   Icon            =   "frmCierrePeriodo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
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
      Height          =   555
      Index           =   1
      Left            =   60
      Picture         =   "frmCierrePeriodo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   1275
   End
   Begin VB.TextBox txtEstado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   1200
      TabIndex        =   1
      Top             =   630
      Width           =   2625
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Anular Cierre"
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
      Left            =   1335
      Picture         =   "frmCierrePeriodo.frx":0534
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1260
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
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
      Height          =   555
      Index           =   0
      Left            =   2610
      Picture         =   "frmCierrePeriodo.frx":0636
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1260
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker dtpPeriodo 
      Height          =   465
      Left            =   1200
      TabIndex        =   0
      Top             =   60
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM yyyy"
      Format          =   52166659
      UpDown          =   -1  'True
      CurrentDate     =   37932
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Estado :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   697
      Width           =   885
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Periodo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   930
   End
End
Attribute VB_Name = "frmCierrePeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPeriodo As Recordset
Dim sPeriodo As String
Dim wExiste As Boolean

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Cierre
               If txtEstado.Text = "Cerrado" Then
                  MsgBox "Periodo cerrado", vbInformation, sMensaje
                  Exit Sub
               End If
                              
               If MsgBox("Deseas cerrar este periodo?", vbQuestion + vbYesNo, "Cierre de Periodo") = vbNo Then
                  Exit Sub
               End If
               
               If sPeriodo = Format(Date, "YYYYMM") Then
                  MsgBox "No se puede cerra el periodo Actual", vbCritical, sMensaje
                  Exit Sub
               End If
                              
               If wExiste Then
                  Cn.Execute "update MCIERRE set lCierre = 1 where tPeriodo = '" & sPeriodo & "'"
               Else
                  Cn.Execute "insert into MCIERRE (tPeriodo, lCierre) values ('" & sPeriodo & "', 1)"
               End If
               RsPeriodo.Requery
               RsPeriodo.MoveFirst
               RsPeriodo.Find ("tPeriodo='" & sPeriodo & "'")
               txtEstado.Text = IIf(RsPeriodo!lCierre, "Cerrado", "Abierto")
               txtEstado.ForeColor = IIf(RsPeriodo!lCierre, &H800000, &HC0&)
               wExiste = True
               
          Case Is = 1 ' Salir
               Unload Me
               
          Case Is = 2 ' Anular
               If txtEstado.Text = "Abierto" Then
                  MsgBox "Periodo abierto", vbInformation, sMensaje
                  Exit Sub
               End If
               
               If MsgBox("Deseas cerrar este periodo?", vbQuestion + vbYesNo, "Cierre de Periodo") = vbNo Then
                  Exit Sub
               End If
                    
               Cn.Execute "update MCIERRE set lCierre = 0 where tPeriodo = '" & sPeriodo & "'"
               RsPeriodo.Requery
               RsPeriodo.MoveFirst
               RsPeriodo.Find ("tPeriodo='" & sPeriodo & "'")
               txtEstado.Text = IIf(RsPeriodo!lCierre, "Cerrado", "Abierto")
               txtEstado.ForeColor = IIf(RsPeriodo!lCierre, &H800000, &HC0&)
               wExiste = True
   End Select
   
End Sub

Private Sub dtpPeriodo_Change()
   Screen.MousePointer = vbHourglass
   sPeriodo = Format(dtpPeriodo.Value, "YYYYMM")
   If RsPeriodo.RecordCount = 0 Then
      txtEstado.Text = "Abierto"
      txtEstado.ForeColor = &HC0&
      wExiste = False
   Else
      RsPeriodo.MoveFirst
      RsPeriodo.Find ("tPeriodo='" & sPeriodo & "'")
      If Not RsPeriodo.EOF Then
         txtEstado.Text = IIf(RsPeriodo!lCierre, "Cerrado", "Abierto")
         txtEstado.ForeColor = IIf(RsPeriodo!lCierre, &H800000, &HC0&)
         wExiste = True
      Else
         txtEstado.Text = "Abierto"
         txtEstado.ForeColor = &HC0&
         wExiste = False
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Centrar Me
   dtpPeriodo.Value = Date
   sPeriodo = Format(dtpPeriodo.Value, "YYYYMM")
   Set RsPeriodo = Lib.OpenRecordset("select * from MCIERRE", Cn)
   
   If RsPeriodo.RecordCount = 0 Then
      txtEstado.Text = "Abierto"
      wExiste = False
   Else
      RsPeriodo.MoveFirst
      RsPeriodo.Find ("tPeriodo='" & sPeriodo & "'")
      If Not RsPeriodo.EOF Then
         txtEstado.Text = IIf(RsPeriodo!lCierre, "Cerrado", "Abierto")
         txtEstado.ForeColor = IIf(RsPeriodo!lCierre, &H800000, &HC0&)
         wExiste = True
      Else
         txtEstado.Text = "Abierto"
         wExiste = False
      End If
   End If
End Sub
