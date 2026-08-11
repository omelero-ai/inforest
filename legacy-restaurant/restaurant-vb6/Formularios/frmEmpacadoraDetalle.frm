VERSION 5.00
Begin VB.Form frmEmpacadoraDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2640
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9390
   Icon            =   "frmEmpacadoraDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   9390
   Begin VB.Frame fraDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   1410
      TabIndex        =   16
      Top             =   0
      Width           =   7980
      Begin VB.TextBox txtDetallado 
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
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   750
         Width           =   5070
      End
      Begin VB.TextBox txtResumido 
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
         Left            =   1485
         MaxLength       =   24
         TabIndex        =   2
         Text            =   " "
         Top             =   1140
         Width           =   2595
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   870
         TabIndex        =   3
         Top             =   1500
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Detallada :"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   795
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   1185
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   825
         TabIndex        =   17
         Top             =   405
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9330
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1890
      Width           =   9390
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4470
         TabIndex        =   9
         Top             =   60
         Width           =   4530
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3000
            Picture         =   "frmEmpacadoraDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3480
            Picture         =   "frmEmpacadoraDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   3960
            Picture         =   "frmEmpacadoraDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmEmpacadoraDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmEmpacadoraDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmEmpacadoraDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1530
            TabIndex        =   20
            Top             =   150
            Width           =   1395
         End
      End
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
         Left            =   5790
         Picture         =   "frmEmpacadoraDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
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
         Left            =   4620
         Picture         =   "frmEmpacadoraDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
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
         Index           =   2
         Left            =   6960
         Picture         =   "frmEmpacadoraDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
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
         Left            =   8130
         Picture         =   "frmEmpacadoraDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   1905
      Left            =   15
      Picture         =   "frmEmpacadoraDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1380
   End
End
Attribute VB_Name = "frmEmpacadoraDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBoton As Recordset

Sub Asignar()
    With frmEmpacadora.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!Codigo), "", !Codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
End Sub
Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmEmpacadora.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmEmpacadora.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmEmpacadora.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmEmpacadora.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmEmpacadora.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmEmpacadora.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmEmpacadora.RsCabecera.RecordCount = 0, 0, frmEmpacadora.RsCabecera.AbsolutePosition) & " de " & frmEmpacadora.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.Value = 1
                'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  Sw = False
                  'Cambiar el SQL
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='EMPACADOR' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "0001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 4)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, nboton, tValor, lActivo) " & _
                         "values ('EMPACADOR', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', 0, " & _
                                " '', " & _
                                       chkActivo.Value & ") "
            
                  Cn.Execute Isql
                  frmEmpacadora.RsCabecera.Sort = "Codigo ASC"
                  frmEmpacadora.RsCabecera.Requery
                  frmEmpacadora.RsCabecera.MoveLast
                  Asignar
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmEmpacadora.RsCabecera.RecordCount = 0, 0, frmEmpacadora.RsCabecera.AbsolutePosition) & " de " & frmEmpacadora.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "lActivo =" & chkActivo.Value & _
                         " where tTABLA = 'EMPACADOR' and tCodigo = '" & txtCodigo & "'"
                   Cn.Execute Isql
                                      
                   nPos = frmEmpacadora.RsCabecera.Bookmark
                   frmEmpacadora.RsCabecera.Requery
                   
                   If frmEmpacadora.RsCabecera.RecordCount = 0 Then
                      frmEmpacadora.RsCabecera.Filter = adFilterNone
                   End If
                   frmEmpacadora.RsCabecera.Bookmark = nPos
                   Asignar
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmEmpacadora.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Empacadora" & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If Calcular("select count(tEmpacador) as Codigo From MPEDIDO where tEmpacador='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Se tiene movimientos con este Empacador, " & Chr(13) & "No se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'EMPACADOR' and tCodigo = '" & txtCodigo & "'"
    
               frmEmpacadora.RsCabecera.Requery
               If frmEmpacadora.RsCabecera.RecordCount <> 0 Then
                  frmEmpacadora.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmEmpacadora.RsCabecera.RecordCount = 0, 0, frmEmpacadora.RsCabecera.AbsolutePosition) & " de " & frmEmpacadora.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
         
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Empacadores "
    fraDetalle.Caption = Me.Caption
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
    Else
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmEmpacadora.RsCabecera.RecordCount = 0, 0, frmEmpacadora.RsCabecera.AbsolutePosition) & " de " & frmEmpacadora.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmEmpacadoraDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
End Sub

Private Sub txtDetallado_LostFocus()
   Call ValidaStr(txtDetallado)
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub
